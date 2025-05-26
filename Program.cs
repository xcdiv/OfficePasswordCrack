using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace OfficePasswordCrack
{
    public class Program
    {
        static async Task Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("使用方法: OfficePasswordCrack.exe [Excel文件路径] [字典文件路径] [TaskNum]");
                return;
            }

            string excelFilePath = args[0];
            string dictFilePath = args[1];
            string foundPassword = "";
            string TaskNum = args[2];

            // 生成Excel关联的任务记录文件路径（Excel名称+_task.txt）
            string taskRecordPath = Path.Combine(
                Path.GetDirectoryName(excelFilePath),
                $"{Path.GetFileNameWithoutExtension(excelFilePath)}_{TaskNum}_task.txt"
            );
            int startLine = 0;

            // 加载上次记录的行数（从Excel关联的task文件）
            if (File.Exists(taskRecordPath))
            {
                try
                {
                    string savedLineText = File.ReadAllText(taskRecordPath).Trim();
                    if (int.TryParse(savedLineText, out int savedLine))
                    {
                        startLine = savedLine;
                        Console.WriteLine($"检测到续传文件，从字典第 {startLine + 1} 行开始处理");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"加载续传记录失败: {ex.Message}，将从第1行开始");
                }
            }

            try
            {
                int currentLine = 0;
                using (var fileStream = new FileStream(dictFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var reader = new StreamReader(fileStream, Encoding.UTF8))
                {
                    string line;
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        currentLine++;
                        // 跳过已处理的行
                        if (currentLine <= startLine) continue;

                        if (string.IsNullOrWhiteSpace(line))
                        {
                            File.WriteAllText(taskRecordPath, currentLine.ToString()); // 记录空行行数
                            continue;
                        }

                        string password = line.Trim();
                        try
                        {
                            using (var package = new ExcelPackage(new FileInfo(excelFilePath), password))
                            {
                                if (package.Workbook.Worksheets.Count > 0)
                                {
                                    foundPassword = password;
                                    Console.WriteLine($"{password} pass!");
                                    break;
                                }
                            }
                        }
                        catch (Exception err)
                        {
                            // password skip Please set the license using one of the methods on the static property ExcelPackage.License. See https://epplussoftware.com/developers/licensenotsetexception for more information
                            if (err.Message.ToString().IndexOf("Please set the license using one of the methods on the static property ExcelPackage.License") > -1)
                            {
                                foundPassword = password;
                                Console.WriteLine($"{password} find ！");
                                break;
                            }
                            else
                            {
                                Console.WriteLine($"{password} skip {err.Message}");
                            }
                        }
                        finally
                        {
                            File.WriteAllText(taskRecordPath, currentLine.ToString()); // 实时更新处理行数
                        }
                    }
                }

                if (!string.IsNullOrEmpty(foundPassword))
                {
                    // 找到密码后删除续传记录
                    if (File.Exists(taskRecordPath)) File.Delete(taskRecordPath);
                    string outputPath = $"{excelFilePath}.txt";
                    File.WriteAllText(outputPath, foundPassword);
                    Console.WriteLine($"找到密码: {foundPassword}，已保存至 {outputPath}");
                }
                else
                {
                    Console.WriteLine($"本次处理至字典第 {currentLine} 行，续传记录保存在 {taskRecordPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"程序异常: {ex.Message}");
            }
        }

        // 异步尝试Excel密码（EPPlus实现，仅支持.xlsx）
        private static Task<bool> TryExcelPasswordAsync(string filePath, string password)
        {
            return Task.Run(() =>
            {
                if (!Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("仅支持 .xlsx 格式文件");
                    return false;
                }

                try
                {
                    using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                    using (var package = new ExcelPackage(fileStream, password))
                    {
                        return package.Workbook.Worksheets.Count > 0;
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    Console.WriteLine($"权限不足: {ex.Message}");
                    return false;
                }
                catch (Exception ex) when (
                    ex is InvalidDataException ||
                    ex.Message.Contains("password") ||
                    ex.Message.Contains("加密"))
                {
                    return false;
                }
            });
        }
    }
}
