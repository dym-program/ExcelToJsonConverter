using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Windows.Forms;  // 需要引入 Windows.Forms 包
using System.Windows.Controls;

namespace ExcelToJsonConverter
{
    using Newtonsoft.Json;
    using System;

    public class CustomJsonConverter : JsonConverter
    {
        // 用于判断字段类型
        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(double) || objectType == typeof(float);
        }

        // 自定义数字序列化规则
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            if (value is double numberValue)
            {
                // 如果是数字且是整数，去掉小数点
                if (numberValue % 1 == 0)
                {
                    writer.WriteValue(Convert.ToInt64(numberValue));  // 写为整数
                }
                else
                {
                    writer.WriteValue(numberValue);  // 否则保持小数
                }
            }
            else
            {
                writer.WriteValue(value);
            }
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }



    public class JsonHelper
    {
        // 自定义JsonConverter来处理数字类型
        public class IntegerConverter : JsonConverter
        {
            public override bool CanConvert(Type objectType)
            {
                return objectType == typeof(double) || objectType == typeof(float);
            }

            public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
            {
                // 如果是整数值，则直接写出整数部分
                if (value is double doubleValue)
                {
                    if (doubleValue == (long)doubleValue)
                    {
                        writer.WriteValue((long)doubleValue); // 作为整数写出
                    }
                    else
                    {
                        writer.WriteValue(doubleValue); // 保持浮动值
                    }
                }
                else if (value is float floatValue)
                {
                    if (floatValue == (long)floatValue)
                    {
                        writer.WriteValue((long)floatValue); // 作为整数写出
                    }
                    else
                    {
                        writer.WriteValue(floatValue); // 保持浮动值
                    }
                }
            }

            public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
            {
                return reader.Value; // 默认行为
            }
        }

        // 清理字符串中的转义字符
        private static string CleanString(string value)
        {
            return value.Replace("\\\"", "\"");
        }

        public static string Serialize(object data)
        {
            // 创建序列化设置
            var settings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                Converters = new List<JsonConverter> { new IntegerConverter(), new CustomJsonConverter() }
            };

            // 使用 JsonConvert.SerializeObject 来序列化数据，传入 settings
            var json = JsonConvert.SerializeObject(data, settings);

            // 这里不需要手动替换转义字符
            return json;
        }

    }

    public partial class MainWindow : Window
    {
        // 存储 config.txt 中的映射关系
        private Dictionary<string, string> ExcelToJsonMapping = new Dictionary<string, string>();

        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }

        // 窗口加载时调用
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadConfigMapping();  // 加载映射关系
            LoadExcelFiles();  // 加载 Excel 文件列表
        }

        // 加载 config.txt 文件中的映射
        private void LoadConfigMapping()
        {
            var configPath = "config.txt";  // 配置文件路径
            if (File.Exists(configPath))
            {
                var lines = File.ReadAllLines(configPath);
                foreach (var line in lines)
                {
                    var parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        var excelFileName = parts[0].Trim();
                        var jsonFileName = parts[1].Trim();

                        // 获取当前应用程序目录，确保 excelFileName 是相对路径
                        var fullExcelFilePath = Path.Combine(Directory.GetCurrentDirectory(), excelFileName);

                        // 将映射关系存储为绝对路径的 Excel 文件路径与目标 JSON 文件的对应关系
                        ExcelToJsonMapping[fullExcelFilePath] = jsonFileName;
                    }
                }
            }
        }


        // 加载 Excel 文件
        private void LoadExcelFiles()
        {
            var excelFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx")
                           .Where(file => !Path.GetFileName(file).StartsWith("~"))
                           .ToList();

            // 这里可以将文件路径或文件名显示在 ListBox 中
            ExcelFileListBox.ItemsSource = excelFiles.ToList();
        }

        // 选择输出目录
        private void SelectOutputDirectory(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                OutputDirectoryTextBox.Text = dialog.SelectedPath;
            }
        }

        // 转换 Excel 文件为 JSON
        private void ConvertExcelToJson(object sender, RoutedEventArgs e)
        {
            var outputDirectory = OutputDirectoryTextBox.Text;
            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                System.Windows.MessageBox.Show("请先选择输出目录");
                return;
            }

            if (!Directory.Exists(outputDirectory))
            {
                System.Windows.MessageBox.Show("输出目录无效");
                return;
            }

            var excelFiles = ExcelFileListBox.ItemsSource.Cast<string>().ToList();
            foreach (var excelFile in excelFiles)
            {
                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(excelFile);

                // 确保对比时的文件名是相对路径，并且去除扩展名
                if (ExcelToJsonMapping.ContainsKey(excelFile))
                {
                    var jsonFileName = ExcelToJsonMapping[excelFile];
                    var jsonFilePath = Path.Combine(outputDirectory, jsonFileName);

                    // 读取 Excel 文件并转换为 JSON
                    ConvertExcelToJsonFile(excelFile, jsonFilePath);
                }
                else
                {
                    System.Windows.MessageBox.Show($"配置中没有找到对应的映射：{fileNameWithoutExtension}");
                }
            }

            System.Windows.MessageBox.Show("转换完成");
        }
        // 转换 Excel 文件为 JSON 文件
        private void ConvertExcelToJsonFile(string excelFilePath, string jsonFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var sheets = package.Workbook.Worksheets;
                var allJsonData = new Dictionary<string, List<Dictionary<string, object>>>();

                foreach (var worksheet in sheets)
                {
                    if (worksheet.Dimension == null)
                    {
                        continue; // 如果没有数据，跳过
                    }

                    var sheetName = worksheet.Name;
                    var rowCount = worksheet.Dimension.Rows;
                    var columnCount = worksheet.Dimension.Columns;

                    // 第一行：字段名称
                    var fieldNames = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        fieldNames.Add(worksheet.Cells[1, col].Text);
                    }

                    // 第二行：字段类型
                    var fieldTypes = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        fieldTypes.Add(worksheet.Cells[2, col].Text.ToLower());
                    }

                    // 第三行：注释
                    var comments = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        comments.Add(worksheet.Cells[3, col].Text);
                    }

                    var items = new List<Dictionary<string, object>>();

                    // 从第四行开始是数据
                    for (int row = 4; row <= rowCount; row++)
                    {
                        var item = new Dictionary<string, object>();
                        for (int col = 1; col <= columnCount; col++)
                        {
                            var fieldName = fieldNames[col - 1];
                            var fieldType = fieldTypes[col - 1];
                            var cellValue = worksheet.Cells[row, col].Text;

                            // 根据字段类型进行不同处理
                            switch (fieldType)
                            {
                                case "number":
                                    item[fieldName] = double.TryParse(cellValue, out var numberValue) ? numberValue : 0;
                                    break;

                                case "string":
                                    item[fieldName] = cellValue;
                                    break;

                                case "array":
                                    item[fieldName] = ParseArray(cellValue);
                                    break;

                                default:
                                    item[fieldName] = cellValue;
                                    break;
                            }
                        }
                        items.Add(item);
                    }

                    allJsonData[sheetName] = items;
                }

                // 序列化之前，清理转义字符
                CleanJsonData(allJsonData);

                // 将转换后的数据写入 JSON 文件
                var json = JsonHelper.Serialize(allJsonData);
                File.WriteAllText(jsonFilePath, json);
            }
        }

        // 清理 JSON 数据中的转义字符
        private void CleanJsonData(Dictionary<string, List<Dictionary<string, object>>> allJsonData)
        {
            foreach (var sheet in allJsonData)
            {
                foreach (var item in sheet.Value)
                {
                    foreach (var key in item.Keys.ToList())
                    {
                        if (item[key] is string strValue)
                        {
                            // 处理字符串，去掉转义字符
                            item[key] = strValue.Replace("\\\"", "\"");
                        }
                        else if (item[key] is List<object> arrayValue)
                        {
                            // 处理数组，确保每个元素都是干净的
                            var cleanArray = new List<object>();
                            foreach (var arrayItem in arrayValue)
                            {
                                if (arrayItem is string arrayStr)
                                {
                                    cleanArray.Add(arrayStr.Replace("\\\"", "\""));
                                }
                                else
                                {
                                    cleanArray.Add(arrayItem);
                                }
                            }
                            item[key] = cleanArray;
                        }
                    }
                }
            }
        }

        private object ParseArray(string cellValue)
        {
            cellValue = cellValue.Trim();

            // 如果是数组的形式，处理里面的元素
            if (cellValue.StartsWith("{") && cellValue.EndsWith("}"))
            {
                var arrayString = cellValue.Substring(1, cellValue.Length - 2).Trim();
                var elements = arrayString.Split(',');

                var array = new List<object>();

                foreach (var element in elements)
                {
                    var trimmedElement = element.Trim();

                    // 1. 去除转义字符
                    trimmedElement = trimmedElement.Replace("\\\"", "\"").Replace("\\\\", "\\");

                    // 2. 处理有引号的字符串
                    if (trimmedElement.StartsWith("\"") && trimmedElement.EndsWith("\""))
                    {
                        // 去除外部的引号
                        trimmedElement = trimmedElement.Trim('"');
                        array.Add(trimmedElement);  // 将清理过的字符串添加到数组中
                    }
                    else if (trimmedElement.StartsWith("{") && trimmedElement.EndsWith("}"))
                    {
                        // 处理嵌套对象
                        var jsonObj = JsonConvert.DeserializeObject<Dictionary<string, object>>(trimmedElement);
                        array.Add(jsonObj);
                    }
                    else
                    {
                        // 普通的值（如数字、字符串）
                        array.Add(trimmedElement);
                    }
                }

                return array;
            }

            return new List<object>();
        }



        private Dictionary<string, object> ParseNestedObject(string nestedObject)
        {
            var result = new Dictionary<string, object>();
            nestedObject = nestedObject.TrimStart('{').TrimEnd('}').Trim();

            var keyValuePairs = nestedObject.Split(',');

            foreach (var pair in keyValuePairs)
            {
                var parts = pair.Split('=');
                if (parts.Length == 2)
                {
                    result[parts[0].Trim()] = parts[1].Trim();
                }
            }

            return result;
        }

    }
}
