using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace ExcelToJsonConverter
{
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

        public static string Serialize(object data)
        {
            var settings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                Converters = new List<JsonConverter> { new IntegerConverter() }
            };

            string json = JsonConvert.SerializeObject(data, settings);

            // 清理所有字符串中的转义字符
            json = json.Replace("\\\"", "\"");

            return json;
        }
    }

    public class ExcelFileInfo
    {
        public string ExcelFileName { get; set; }
        public string JsonFileName { get; set; }
    }

    public partial class MainWindow : Window
    {
        public ObservableCollection<ExcelFileInfo> FileList { get; set; }

        public string OutputDirectory
        {
            get { return (string)GetValue(OutputDirectoryProperty); }
            set { SetValue(OutputDirectoryProperty, value); }
        }

        public static readonly DependencyProperty OutputDirectoryProperty =
            DependencyProperty.Register("OutputDirectory", typeof(string), typeof(MainWindow), new PropertyMetadata("..\\..\\config"));

        private Dictionary<string, string> ExcelToJsonMapping = new Dictionary<string, string>();
        private FileSystemWatcher fileWatcher;

        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            this.DataContext = this;
            FileList = new ObservableCollection<ExcelFileInfo>();

            // 设置文件监控器
            fileWatcher = new FileSystemWatcher(Directory.GetCurrentDirectory(), "config.txt")
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName,
                EnableRaisingEvents = true
            };

            fileWatcher.Changed += FileWatcher_Changed;
            fileWatcher.Created += FileWatcher_Changed;
            fileWatcher.Deleted += FileWatcher_Changed;
        }

        // 文件变动时刷新加载配置
        private void FileWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            Dispatcher.Invoke(() => LoadConfigMapping());
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadConfigMapping();  // Load file mappings
        }

        // Load mappings from config.txt
        private void LoadConfigMapping()
        {
            FileList.Clear(); // 清空现有列表
            var configPath = "config.txt";
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

                        if (File.Exists(excelFileName)) // Check if file exists
                        {
                            FileList.Add(new ExcelFileInfo
                            {
                                ExcelFileName = excelFileName,
                                JsonFileName = jsonFileName
                            });

                            ExcelToJsonMapping[excelFileName] = jsonFileName; // Store mapping
                        }
                        else
                        {
                            FileList.Add(new ExcelFileInfo
                            {
                                ExcelFileName = excelFileName,
                                JsonFileName = "文件不存在"
                            });

                            ExcelToJsonMapping[excelFileName] = "文件不存在"; // Store mapping
                        }
                    }
                }
            }
            else
            {
                System.Windows.MessageBox.Show("config.txt 文件不存在");
            }
        }

        // Double-click ListView item to trigger conversion
        private void ExcelFileListView_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // 通过 sender 获取 ListView 控件
            var listView = sender as System.Windows.Controls.ListView;
            if (listView == null) return;

            // 获取当前选中的项，类型为 ExcelFileInfo
            var selectedFile = listView.SelectedItem as ExcelFileInfo;
            if (selectedFile != null)
            {
                // 清空日志输出窗口
                LogOutputTextBox.Clear();

                // 检查文件是否存在
                if (!File.Exists(selectedFile.ExcelFileName))
                {
                    LogMessage($"错误: 文件 {selectedFile.ExcelFileName} 不存在", "Red");
                    return;
                }

                try
                {
                    // 转换文件
                    ConvertExcelToJson(selectedFile.ExcelFileName, selectedFile.JsonFileName);
                    // 记录成功日志（绿色文字）
                    LogMessage($"转换成功: {selectedFile.ExcelFileName} -> {selectedFile.JsonFileName}", "Green");
                }
                catch (Exception ex)
                {
                    // 记录错误日志（红色文字）
                    LogMessage($"错误: {selectedFile.ExcelFileName} -> {selectedFile.JsonFileName} - {ex.Message}", "Red");
                }
            }
            else
            {
                System.Windows.MessageBox.Show("请选择一个文件进行转换");
            }
        }


        // Handle conversion when the "转换为JSON" button is clicked (convert all files)
        private void ConvertExcelToJsonButton_Click(object sender, RoutedEventArgs e)
        {
            // 清空日志输出窗口
            LogOutputTextBox.Clear();

            // Loop through all files in FileList and convert them
            foreach (var file in FileList)
            {
                if (file.JsonFileName != "文件不存在")
                {
                    try
                    {
                        ConvertExcelToJson(file.ExcelFileName, file.JsonFileName);
                        // 记录成功日志
                        LogMessage($"转换成功: {file.ExcelFileName} -> {file.JsonFileName}", "Green");
                    }
                    catch (Exception ex)
                    {
                        // 记录错误日志
                        LogMessage($"错误: {file.ExcelFileName} -> {file.JsonFileName} - {ex.Message}", "Red");
                    }
                }
                else
                {
                    LogMessage($"文件不存在: {file.ExcelFileName} - 跳过转换", "Red");
                }
            }
        }

        private void ConvertExcelToJson(string excelFile, string jsonFileName)
        {
            //var outputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "config");
            var outputDirectory = OutputDirectory; // 直接获取最新的编辑框值
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            var jsonFilePath = Path.Combine(outputDirectory, jsonFileName);
            ConvertExcelToJsonFile(excelFile, jsonFilePath);
        }

        private void ConvertExcelToJsonFile(string excelFilePath, string jsonFilePath)
        {
            var excelFile = new FileInfo(excelFilePath);
            using (var package = new ExcelPackage(excelFile))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    var rowCount = worksheet.Dimension.Rows;
                    var columnCount = worksheet.Dimension.Columns;

                    var data = new List<Dictionary<string, object>>();

                    var header = new List<string>();
                    for (var col = 1; col <= columnCount; col++)
                    {
                        header.Add(worksheet.Cells[1, col].Text);
                    }

                    for (var row = 2; row <= rowCount; row++)
                    {
                        var rowDict = new Dictionary<string, object>();
                        for (var col = 1; col <= columnCount; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Text;
                            rowDict[header[col - 1]] = cellValue;
                        }
                        data.Add(rowDict);
                    }

                    var json = JsonHelper.Serialize(data);
                    File.WriteAllText(jsonFilePath, json);
                }
            }
        }

        // Log message with color output in the output window
        private void LogMessage(string message, string color)
        {
            LogOutputTextBox.AppendText($"{DateTime.Now}: {message}\n");
            if (color == "Red")
            {
                LogOutputTextBox.Foreground = System.Windows.Media.Brushes.Red;
            }
            else
            {
                LogOutputTextBox.Foreground = System.Windows.Media.Brushes.Green;
            }
            LogOutputTextBox.ScrollToEnd();
        }

        private void SelectOutputDirectory(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                OutputDirectoryTextBox.Text = dialog.SelectedPath;
            }
        }
    }
}
