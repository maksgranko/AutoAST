using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using System.Diagnostics;
using System.Threading;
using System.Data.OleDb;
using System.IO;
using System.Data;
using System.Drawing.Imaging;
using System.Drawing;
using System.Text.RegularExpressions;

namespace AutoAST
{
    public static class ScreenshotHelper
    {
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, ref Rectangle rectangle);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdc, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, uint dwRop);

        private const uint SRCCOPY = 0x00CC0020;

        public static Bitmap CaptureWindow(IntPtr handle)
        {
            try
            {
                // Получаем размеры окна
                Rectangle rect = new Rectangle();
                GetWindowRect(handle, ref rect);
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;

                // Создаем битмап для сохранения изображения
                Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);

                // Получаем DC окна
                IntPtr hdcWindow = GetWindowDC(handle);

                // Создаем совместимый DC
                IntPtr hdcMemDC = CreateCompatibleDC(hdcWindow);

                // Создаем совместимый битмап
                IntPtr hBitmap = CreateCompatibleBitmap(hdcWindow, width, height);

                // Выбираем битмап в DC
                IntPtr hOld = SelectObject(hdcMemDC, hBitmap);

                // Копируем изображение
                BitBlt(hdcMemDC, 0, 0, width, height, hdcWindow, 0, 0, SRCCOPY);

                // Восстанавливаем DC
                SelectObject(hdcMemDC, hOld);

                // Конвертируем в Bitmap
                bmp = Image.FromHbitmap(hBitmap);

                // Очищаем ресурсы
                DeleteObject(hBitmap);
                DeleteDC(hdcMemDC);
                ReleaseDC(handle, hdcWindow);

                return bmp;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при создании скриншота: {ex.Message}");
                return null;
            }
        }

        public static void SaveWindowScreenshot(IntPtr handle, string filename)
        {
            try
            {
                using (Bitmap bmp = CaptureWindow(handle))
                {
                    if (bmp != null)
                    {
                        string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Screenshots");
                        Directory.CreateDirectory(directory);
                        string fullPath = Path.Combine(directory, filename);
                        bmp.Save(fullPath, ImageFormat.Png);
                        Console.WriteLine($"Скриншот сохранен: {fullPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении скриншота: {ex.Message}");
            }
        }
    }

    public static class Win32Helper
    {
        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        public const int GWL_STYLE = -16;
        public const uint WM_SETTEXT = 0x000C;
    }

    public class TextBoxAutomation
    {
        private readonly AutomationElement _element;
        private readonly IntPtr _hwnd;

        public TextBoxAutomation(AutomationElement element)
        {
            _element = element ?? throw new ArgumentNullException(nameof(element));
            _hwnd = new IntPtr(element.Current.NativeWindowHandle);
        }

        public string GetCurrentValue()
        {
            try
            {
                // Пробуем получить значение через UI Automation
                object pattern;
                if (_element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                {
                    var valuePattern = pattern as ValuePattern;
                    if (valuePattern != null)
                    {
                        return valuePattern.Current.Value?.Trim() ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not get value via UI Automation: {ex.Message}");
            }

            return string.Empty;
        }

        public bool SetValue(string value)
        {
            if (value == null)
                return false;

            // Очищаем значение от лишних пробелов
            value = value.Trim();

            // Проверяем текущее значение
            string currentValue = GetCurrentValue();
            if (currentValue == value)
            {
                Console.WriteLine("Value already matches, skipping update");
                return true;
            }

            try
            {
                // Пробуем разные способы установить фокус
                if (!SetFocus())
                {
                    Console.WriteLine("Warning: Could not set focus");
                }

                // Пробуем установить текст через UI Automation
                if (SetValueViaAutomation(value))
                    return true;

                // Если не получилось через UI Automation, пробуем через Win32
                if (SetValueViaWin32(value))
                    return true;

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting value: {ex.Message}");
                return false;
            }
        }

        private bool SetFocus()
        {
            bool focusSet = false;

            // Пробуем через UI Automation
            try
            {
                _element.SetFocus();
                Thread.Sleep(100);
                focusSet = true;
            }
            catch
            {
                Console.WriteLine("Warning: Could not set focus via UI Automation");
            }

            // Пробуем через Win32
            try
            {
                Win32Helper.SetForegroundWindow(_hwnd);
                Thread.Sleep(100);
                Win32Helper.SetFocus(_hwnd);
                Thread.Sleep(100);
                focusSet = true;
            }
            catch
            {
                Console.WriteLine("Warning: Could not set focus via Win32");
            }

            return focusSet;
        }

        private bool SetValueViaAutomation(string value)
        {
            try
            {
                object pattern;
                if (_element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                {
                    var valuePattern = pattern as ValuePattern;
                    if (valuePattern != null)
                    {
                        valuePattern.SetValue(value);
                        Console.WriteLine("Value set via UI Automation pattern");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not set value via UI Automation: {ex.Message}");
            }
            return false;
        }

        private bool SetValueViaWin32(string value)
        {
            try
            {
                IntPtr textPtr = Marshal.StringToHGlobalUni(value);
                try
                {
                    bool result = Win32Helper.PostMessage(_hwnd, Win32Helper.WM_SETTEXT, IntPtr.Zero, textPtr);
                    if (result)
                    {
                        Console.WriteLine("Value set via Win32 PostMessage");
                        return true;
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(textPtr);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not set value via Win32: {ex.Message}");
            }
            return false;
        }
    }

    public class JetDbExporter
    {
        private readonly string _dbPath;
        private readonly string _outputDir;

        public JetDbExporter(string dbPath, string outputDir)
        {
            _dbPath = dbPath ?? throw new ArgumentNullException(nameof(dbPath));
            _outputDir = outputDir ?? throw new ArgumentNullException(nameof(outputDir));

            if (!File.Exists(dbPath))
                throw new FileNotFoundException("Database file not found", dbPath);

            Directory.CreateDirectory(outputDir);
        }

        public void ExportAllTablesToCSV()
        {
            // Пробуем разные провайдеры
            var providers = new[]
            {
                "Microsoft.Jet.OLEDB.4.0",
                "Microsoft.ACE.OLEDB.12.0"
            };

            Exception lastException = null;
            foreach (var provider in providers)
            {
                try
                {
                    string connectionString = $"Provider={provider};Data Source={_dbPath};Persist Security Info=False;";
                    using (var connection = new OleDbConnection(connectionString))
                    {
                        connection.Open();
                        Console.WriteLine($"Успешное подключение через провайдер {provider}");

                        // Получаем список всех таблиц
                        DataTable schema = connection.GetSchema("Tables");
                        foreach (DataRow row in schema.Rows)
                        {
                            string tableName = row["TABLE_NAME"].ToString();
                            if (!tableName.StartsWith("MSys")) // Пропускаем системные таблицы
                            {
                                ExportTableToCSV(connection, tableName);
                            }
                        }
                        return; // Если успешно, выходим из метода
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    Console.WriteLine($"Не удалось подключиться через провайдер {provider}: {ex.Message}");
                }
            }

            // Если ни один провайдер не сработал, выводим инструкции
            Console.WriteLine("\nНе удалось подключиться к базе данных. Пожалуйста, выполните следующие действия:");
            Console.WriteLine("1. Установите Microsoft Access Database Engine 2010 Redistributable:");
            Console.WriteLine("   https://www.microsoft.com/en-us/download/details.aspx?id=13255");
            Console.WriteLine("2. Если вы используете 64-битную версию Windows, убедитесь, что:");
            Console.WriteLine("   - Установлена 64-битная версия Access Database Engine");
            Console.WriteLine("   - Проект скомпилирован для платформы x64");
            Console.WriteLine("\nТехническая информация:");
            if (lastException != null)
            {
                Console.WriteLine($"Последняя ошибка: {lastException.Message}");
            }

            throw new Exception("Не удалось подключиться к базе данных. Установите необходимые компоненты.");
        }

        private void ExportTableToCSV(OleDbConnection connection, string tableName)
        {
            string outputPath = Path.Combine(_outputDir, $"{tableName}.csv");
            Console.WriteLine($"Экспорт таблицы {tableName} в {outputPath}");

            using (var command = new OleDbCommand($"SELECT * FROM [{tableName}]", connection))
            using (var reader = command.ExecuteReader())
            using (var writer = new StreamWriter(outputPath, false, Encoding.UTF8))
            {
                // Записываем заголовки
                var headers = new List<string>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    headers.Add(reader.GetName(i));
                }
                writer.WriteLine(string.Join(";", headers));

                // Записываем данные
                while (reader.Read())
                {
                    var values = new List<string>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        string value = reader[i]?.ToString() ?? "";
                        // Экранируем специальные символы
                        value = value.Replace("\"", "\"\"");
                        if (value.Contains(";") || value.Contains("\"") || value.Contains("\n"))
                        {
                            value = $"\"{value}\"";
                        }
                        values.Add(value);
                    }
                    writer.WriteLine(string.Join(";", values));
                }
            }
        }
    }

    public class UiAutomationHelper
    {
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // Получает заголовок окна
        public static string GetWindowTitle(IntPtr hWnd)
        {
            StringBuilder sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, 256);
            return sb.ToString();
        }

        // Находит все окна процесса
        private static List<IntPtr> GetProcessWindows(int processId)
        {
            List<IntPtr> result = new List<IntPtr>();

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                int windowProcessId;
                GetWindowThreadProcessId(hWnd, out windowProcessId);

                if (windowProcessId == processId && IsWindowVisible(hWnd))
                {
                    string title = GetWindowTitle(hWnd);
                    if (!string.IsNullOrWhiteSpace(title))
                    {
                        result.Add(hWnd);
                    }
                }
                return true;
            }, IntPtr.Zero);

            return result;
        }

        // Находит все окна по шаблону заголовка "{ число }"
        public static List<IntPtr> FindTestWindows()
        {
            List<IntPtr> result = new List<IntPtr>();

            // Ищем процесс KT_ATE
            Process[] processes = Process.GetProcessesByName("KT_ATE-12");
            if (processes.Length == 0)
            {
                processes = Process.GetProcesses()
                    .Where(p => p.ProcessName.StartsWith("KT_ATE-"))
                    .ToArray();
            }

            foreach (var process in processes)
            {
                var windows = GetProcessWindows(process.Id);
                foreach (var window in windows)
                {
                    string title = GetWindowTitle(window);
                    // Ищем окна с заголовком { число }
                    if (title.Contains("{") && title.Contains("}"))
                    {
                        result.Add(window);
                    }
                }
            }

            return result;
        }

        public static void PrintElementProperties(AutomationElement element, string indent = "")
        {
            if (element == null) return;

            try
            {
                Console.WriteLine($"{indent}Element:");
                Console.WriteLine($"{indent}  ClassName: {element.Current.ClassName}");
                Console.WriteLine($"{indent}  AutomationId: {element.Current.AutomationId}");
                Console.WriteLine($"{indent}  Name: {element.Current.Name}");
                Console.WriteLine($"{indent}  ControlType: {element.Current.ControlType.ProgrammaticName}");
                Console.WriteLine($"{indent}  BoundingRectangle: {element.Current.BoundingRectangle}");
                Console.WriteLine($"{indent}  NativeWindowHandle: 0x{element.Current.NativeWindowHandle:X}");
                Console.WriteLine($"{indent}  IsEnabled: {element.Current.IsEnabled}");
                Console.WriteLine($"{indent}  IsKeyboardFocusable: {element.Current.IsKeyboardFocusable}");
                Console.WriteLine($"{indent}  HasKeyboardFocus: {element.Current.HasKeyboardFocus}");

                if (element.Current.NativeWindowHandle != 0)
                {
                    int style = Win32Helper.GetWindowLong(new IntPtr(element.Current.NativeWindowHandle), Win32Helper.GWL_STYLE);
                    Console.WriteLine($"{indent}  Window Style: 0x{style:X8}");
                }

                var patterns = element.GetSupportedPatterns();
                if (patterns.Length > 0)
                {
                    Console.WriteLine($"{indent}  Supported Patterns:");
                    foreach (var pattern in patterns)
                    {
                        Console.WriteLine($"{indent}    - {pattern.ProgrammaticName}");
                    }
                }
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{indent}Error getting properties: {ex.Message}");
            }
        }

        // Находит текстовое поле по специфичным характеристикам
        public static AutomationElement FindTargetTextBox(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero)
                return null;

            AutomationElement window = AutomationElement.FromHandle(windowHandle);
            if (window == null)
                return null;

            Console.WriteLine($"Searching in window: {window.Current.Name}");

            try
            {
                // Сначала найдем все элементы окна для отладки
                AutomationElementCollection allElements = window.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                Console.WriteLine($"Total elements found: {allElements.Count}");

                foreach (AutomationElement element in allElements)
                {
                    PrintElementProperties(element, "  ");
                }

                // Теперь попробуем найти наш TextBox
                PropertyCondition classCondition = new PropertyCondition(
                    AutomationElement.ClassNameProperty,
                    "ThunderRT6TextBox"
                );

                AutomationElementCollection textBoxes = window.FindAll(TreeScope.Descendants, classCondition);
                Console.WriteLine($"\nFound {textBoxes.Count} ThunderRT6TextBox elements");

                foreach (AutomationElement textBox in textBoxes)
                {
                    PrintElementProperties(textBox, "  TextBox: ");

                    if (VerifyElementCharacteristics(textBox))
                    {
                        Console.WriteLine("  Found matching TextBox!");
                        return textBox;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during search: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            return null;
        }

        private static bool VerifyElementCharacteristics(AutomationElement element)
        {
            try
            {
                if (!element.Current.IsEnabled || element.Current.IsOffscreen)
                {
                    Console.WriteLine("  Failed: Element not enabled or is offscreen");
                    return false;
                }

                int style = Win32Helper.GetWindowLong(new IntPtr(element.Current.NativeWindowHandle), Win32Helper.GWL_STYLE);
                if (style != 0x540100C0)
                {
                    Console.WriteLine($"  Failed: Wrong style (0x{style:X8} != 0x540100C0)");
                    return false;
                }

                var rect = element.Current.BoundingRectangle;
                if (rect.Left != 155 || rect.Top != 1020 ||
                    rect.Right != 1897 || rect.Bottom != 1050)
                {
                    Console.WriteLine($"  Failed: Wrong position ({rect})");
                    return false;
                }

                // Проверяем иерархию предков
                AutomationElement parent = TreeWalker.ControlViewWalker.GetParent(element);
                bool foundCorrectWindow = false;
                while (parent != null)
                {
                    string name = parent.Current.Name ?? "";
                    Console.WriteLine($"  Checking parent: {name}");

                    if (parent.Current.ControlType == ControlType.Window &&
                        name.Contains("ТЗ №") &&
                        name.Contains("экспертная мера трудности = 0,6") &&
                        name.Contains("мера трудности АСТ = 1"))
                    {
                        foundCorrectWindow = true;
                        break;
                    }

                    if (name.Contains("Конструктор Тестов Адаптивной Среды Тестирования АСТ"))
                    {
                        break;
                    }

                    parent = TreeWalker.ControlViewWalker.GetParent(parent);
                }

                if (!foundCorrectWindow)
                {
                    Console.WriteLine("  Failed: Correct parent window not found");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error in verification: {ex.Message}");
                return false;
            }
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Находим все .ast файлы в директории программы
                string[] astFiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.ast");
                if (!astFiles.Any())
                {
                    Console.WriteLine("Не найдены файлы .ast в директории программы!");
                    Console.ReadKey();
                    return;
                }

                string dbPath = astFiles.First();
                Console.WriteLine($"Найдена база данных: {Path.GetFileName(dbPath)}");

                // Инициализируем парсер базы данных
                using (var parser = new AstParser(dbPath))
                {
                    // Ищем окна тестирования
                    var windows = UiAutomationHelper.FindTestWindows();
                    Console.WriteLine($"\nНайдено окон тестирования: {windows.Count}");

                    foreach (var window in windows)
                    {
                        string title = UiAutomationHelper.GetWindowTitle(window);
                        Console.WriteLine($"\nОбработка окна: {title}");

                        // Делаем скриншот окна до изменений
                        ScreenshotHelper.SaveWindowScreenshot(window, $"before_{DateTime.Now:yyyyMMdd_HHmmss}.png");

                        // Извлекаем номер задания из заголовка
                        var match = Regex.Match(title, @"\{\s*(\d+)\s*\}");
                        if (!match.Success)
                        {
                            // Пробуем альтернативный формат
                            match = Regex.Match(title, @"ТЗ\s*№\s*(\d+)");
                            if (!match.Success)
                            {
                                Console.WriteLine("Не удалось найти номер задания в заголовке окна");
                                continue;
                            }
                        }

                        int taskId = int.Parse(match.Groups[1].Value);
                        Console.WriteLine($"Номер задания: {taskId}");

                        // Получаем информацию о задании из базы
                        var questions = parser.ParseQuestions();
                        var currentQuestion = questions.FirstOrDefault(q => q.Id == taskId);

                        if (currentQuestion == null)
                        {
                            Console.WriteLine($"Задание #{taskId} не найдено в базе данных");
                            continue;
                        }

                        // Выводим информацию о задании
                        Console.WriteLine("\nИнформация о задании:");
                        Console.WriteLine($"Тип: {(currentQuestion.IsOpenQuestion ? "открытый вопрос" : "закрытый вопрос")}");
                        Console.WriteLine($"Текст: {currentQuestion.Text}");
                        Console.WriteLine($"Правильные ответы: {string.Join(", ", currentQuestion.CorrectAnswers)}");
                        if (!currentQuestion.IsOpenQuestion)
                        {
                            Console.WriteLine($"Неправильные ответы: {string.Join(", ", currentQuestion.WrongAnswers)}");
                        }

                        // Ищем поле ввода
                        AutomationElement textBox = UiAutomationHelper.FindTargetTextBox(window);
                        if (textBox != null)
                        {
                            Console.WriteLine("\nНайдено поле ввода, устанавливаем правильный ответ...");
                            var automation = new TextBoxAutomation(textBox);

                            // Для открытых вопросов берем первый правильный ответ
                            // Для закрытых вопросов можно использовать любой правильный ответ
                            string answer = currentQuestion.CorrectAnswers.FirstOrDefault();

                            if (string.IsNullOrEmpty(answer))
                            {
                                Console.WriteLine("Не найден правильный ответ в базе данных");
                                continue;
                            }

                            if (automation.SetValue(answer))
                            {
                                Console.WriteLine($"Установлен ответ: {answer}");
                                // Делаем скриншот окна после изменений
                                Thread.Sleep(500); // Даем время на обновление UI
                                ScreenshotHelper.SaveWindowScreenshot(window, $"after_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                            }
                            else
                            {
                                Console.WriteLine("Не удалось установить значение");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Поле ввода не найдено");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Подробности: {ex.InnerException.Message}");
                }
            }

            Console.WriteLine("\nНажмите любую клавишу для выхода...");
            Console.ReadKey();
        }
    }
}
