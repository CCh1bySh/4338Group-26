using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using OfficeOpenXml;
using Microsoft.Win32;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;

namespace Group4338
{
    public partial class _4338_Чибышев : Window
    {
        // Класс для хранения данных сотрудника
        public class Employee
        {
            public int Id { get; set; }
            public string Login { get; set; }
            public string Password { get; set; }
            public string Role { get; set; }
        }

        // Класс для десериализации JSON (соответствует структуре 5.json)
        public class JsonEmployee
        {
            public string login { get; set; }
            public string password { get; set; }
            public string role { get; set; }
        }

        // Список сотрудников
        private List<Employee> employees = new List<Employee>();

        // Конструктор
        public _4338_Чибышев()
        {
            InitializeComponent();

            // Настройка лицензии EPPlus
            ExcelPackage.License.SetNonCommercialOrganization("Group4338");

            // Загружаем тестовые данные
            LoadSampleData();
        }

        // Загрузка демо-данных
        private void LoadSampleData()
        {
            employees = new List<Employee>
            {
                new Employee { Id = 1, Login = "admin", Password = "12345", Role = "Администратор" },
                new Employee { Id = 2, Login = "user1", Password = "qwerty", Role = "Пользователь" },
                new Employee { Id = 3, Login = "manager", Password = "11111", Role = "Менеджер" },
                new Employee { Id = 4, Login = "petrov", Password = "pass123", Role = "Пользователь" },
                new Employee { Id = 5, Login = "ivanova", Password = "54321", Role = "Менеджер" },
                new Employee { Id = 6, Login = "guest", Password = "guest", Role = "Гость" }
            };

            UpdateDataGrid();
            UpdateStatus("Загружены демо-данные");
        }

        // Обновление DataGrid
        private void UpdateDataGrid()
        {
            dgEmployees.ItemsSource = null;
            dgEmployees.ItemsSource = employees;
            txtRecordCount.Text = $"Записей: {employees.Count}";
        }

        // Обновление статуса
        private void UpdateStatus(string message)
        {
            txtStatus.Text = $"🕒 {DateTime.Now:HH:mm:ss} - {message}";
        }

        // Очистка данных
        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            employees.Clear();
            UpdateDataGrid();
            UpdateStatus("Данные очищены");
        }

        // ============ ЛАБОРАТОРНАЯ РАБОТА №3 (Excel) ============

        // ИМПОРТ ИЗ EXCEL
        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Выберите файл Excel для импорта";
                openFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";

                if (openFileDialog.ShowDialog() == true)
                {
                    employees.Clear();

                    using (var package = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            string login = worksheet.Cells[row, 1].Text;
                            string password = worksheet.Cells[row, 2].Text;
                            string role = worksheet.Cells[row, 3].Text;

                            if (!string.IsNullOrWhiteSpace(login))
                            {
                                employees.Add(new Employee
                                {
                                    Id = employees.Count + 1,
                                    Login = login,
                                    Password = password,
                                    Role = string.IsNullOrEmpty(role) ? "Не указана" : role
                                });
                            }
                        }
                    }

                    UpdateDataGrid();
                    UpdateStatus($"Импортировано {employees.Count} записей из Excel");
                    MessageBox.Show($"Импорт завершен! Загружено {employees.Count} записей.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при импорте Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ЭКСПОРТ В EXCEL
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (employees.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "Сохранить файл Excel";
                saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
                saveFileDialog.FileName = $"Employees_by_Role_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (var package = new ExcelPackage())
                    {
                        var groupedByRole = employees.GroupBy(emp => emp.Role);

                        foreach (var group in groupedByRole)
                        {
                            string sheetName = CleanSheetName(group.Key);
                            var worksheet = package.Workbook.Worksheets.Add(sheetName);

                            // Заголовки
                            worksheet.Cells[1, 1].Value = "Логин";
                            worksheet.Cells[1, 2].Value = "Пароль";
                            worksheet.Cells[1, 1, 1, 2].Style.Font.Bold = true;

                            // Данные
                            int row = 2;
                            foreach (var employee in group)
                            {
                                worksheet.Cells[row, 1].Value = employee.Login;
                                worksheet.Cells[row, 2].Value = employee.Password;
                                row++;
                            }

                            worksheet.Cells.AutoFitColumns();
                        }

                        package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    }

                    UpdateStatus($"Экспортировано в Excel: {System.IO.Path.GetFileName(saveFileDialog.FileName)}");
                    MessageBox.Show("Экспорт в Excel завершен!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Очистка имени листа
        private string CleanSheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Без_роли";

            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            if (name.Length > 30)
                name = name.Substring(0, 30);

            return name;
        }

        // ============ ЛАБОРАТОРНАЯ РАБОТА №4 (JSON и Word) ============

        // ИМПОРТ ИЗ JSON
        private void btnImportJson_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Выберите JSON файл для импорта";
                openFileDialog.Filter = "JSON файлы (*.json)|*.json";

                if (openFileDialog.ShowDialog() == true)
                {
                    // Читаем JSON файл
                    string jsonString = File.ReadAllText(openFileDialog.FileName, Encoding.UTF8);

                    // Десериализуем JSON в список объектов
                    var jsonEmployees = JsonSerializer.Deserialize<List<JsonEmployee>>(jsonString);

                    if (jsonEmployees != null)
                    {
                        employees.Clear();

                        foreach (var jsonEmp in jsonEmployees)
                        {
                            employees.Add(new Employee
                            {
                                Id = employees.Count + 1,
                                Login = jsonEmp.login ?? "",
                                Password = jsonEmp.password ?? "",
                                Role = jsonEmp.role ?? "Не указана"
                            });
                        }

                        UpdateDataGrid();
                        UpdateStatus($"Импортировано {employees.Count} записей из JSON");
                        MessageBox.Show($"JSON импорт завершен! Загружено {employees.Count} записей.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при импорте JSON: {ex.Message}\n\nУбедитесь, что файл имеет правильный формат JSON.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ЭКСПОРТ В WORD
        private void btnExportWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (employees.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "Сохранить документ Word";
                saveFileDialog.Filter = "Word документы (*.docx)|*.docx";
                saveFileDialog.FileName = $"Employees_by_Role_{DateTime.Now:yyyyMMdd_HHmmss}.docx";

                if (saveFileDialog.ShowDialog() == true)
                {
                    // Группируем по ролям
                    var groupedByRole = employees.GroupBy(emp => emp.Role).OrderBy(g => g.Key);

                    // Создаем Word документ
                    using (var document = WordprocessingDocument.Create(saveFileDialog.FileName, WordprocessingDocumentType.Document))
                    {
                        // Добавляем основную часть документа
                        var mainPart = document.AddMainDocumentPart();
                        mainPart.Document = new Document();
                        var body = mainPart.Document.AppendChild(new Body());

                        // Заголовок документа
                        body.AppendChild(CreateParagraph("Отчет по сотрудникам", true, 28));
                        body.AppendChild(CreateParagraph($"Сгенерировано: {DateTime.Now:dd.MM.yyyy HH:mm:ss}", false, 20));
                        body.AppendChild(CreateParagraph("", false, 0)); // Пустая строка

                        // Для каждой роли создаем отдельную страницу
                        int pageCount = 0;
                        foreach (var group in groupedByRole)
                        {
                            pageCount++;

                            // Заголовок роли
                            body.AppendChild(CreateParagraph($"Роль: {group.Key}", true, 24));
                            body.AppendChild(CreateParagraph($"Количество сотрудников: {group.Count()}", false, 20));
                            body.AppendChild(CreateParagraph("", false, 0));

                            // Создаем таблицу для сотрудников
                            Table table = new Table();

                            // Добавляем границы таблицы
                            TableProperties tableProperties = new TableProperties();
                            TableBorders tableBorders = new TableBorders();

                            tableBorders.TopBorder = new TopBorder() { Val = BorderValues.Single, Size = 1 };
                            tableBorders.BottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 1 };
                            tableBorders.LeftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 1 };
                            tableBorders.RightBorder = new RightBorder() { Val = BorderValues.Single, Size = 1 };
                            tableBorders.InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 1 };
                            tableBorders.InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 1 };

                            tableProperties.AppendChild(tableBorders);
                            table.AppendChild(tableProperties);

                            // Добавляем заголовки таблицы
                            TableRow headerRow = new TableRow();

                            // Заголовки с серым фоном
                            headerRow.AppendChild(CreateTableCell("№", true, true));
                            headerRow.AppendChild(CreateTableCell("Логин", true, true));
                            headerRow.AppendChild(CreateTableCell("Пароль", true, true));
                            table.AppendChild(headerRow);

                            // Добавляем данные
                            int rowNum = 1;
                            foreach (var emp in group)
                            {
                                TableRow dataRow = new TableRow();
                                dataRow.AppendChild(CreateTableCell(rowNum.ToString(), false, false));
                                dataRow.AppendChild(CreateTableCell(emp.Login, false, false));
                                dataRow.AppendChild(CreateTableCell(emp.Password, false, false));
                                table.AppendChild(dataRow);
                                rowNum++;
                            }

                            // Добавляем таблицу в документ
                            body.AppendChild(table);

                            // Добавляем пустую строку
                            body.AppendChild(CreateParagraph("", false, 0));

                            // Добавляем разрыв страницы (кроме последней страницы)
                            if (pageCount < groupedByRole.Count())
                            {
                                body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                            }
                        }

                        // Сохраняем документ
                        mainPart.Document.Save();
                    }

                    UpdateStatus($"Экспортировано в Word: {System.IO.Path.GetFileName(saveFileDialog.FileName)}");
                    MessageBox.Show($"Экспорт в Word завершен!\n\nСоздано страниц: {groupedByRole.Count()}\nВсего записей: {employees.Count}", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Вспомогательный метод для создания параграфа в Word
        private Paragraph CreateParagraph(string text, bool bold, int fontSize)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();

            if (!string.IsNullOrEmpty(text))
            {
                run.AppendChild(new Text(text));
            }

            RunProperties runProperties = new RunProperties();

            if (bold)
            {
                runProperties.AppendChild(new Bold());
            }

            if (fontSize > 0)
            {
                // Используем полное имя с пространством имен
                var size = new DocumentFormat.OpenXml.Wordprocessing.FontSize();
                size.Val = (fontSize * 2).ToString(); // В половинных пунктах
                runProperties.AppendChild(size);
            }

            run.PrependChild(runProperties);
            paragraph.AppendChild(run);

            return paragraph;
        }

        // Вспомогательный метод для создания ячейки таблицы в Word
        private TableCell CreateTableCell(string text, bool isHeader, bool hasBackground)
        {
            TableCell cell = new TableCell();
            Paragraph paragraph = new Paragraph();
            Run run = new Run(new Text(text ?? ""));

            // Настройки параграфа (отступы)
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.SpacingBetweenLines = new SpacingBetweenLines()
            {
                Before = "0",
                After = "0",
                Line = "240",
                LineRule = LineSpacingRuleValues.Auto
            };
            paragraph.AppendChild(paragraphProperties);

            // Настройки текста
            RunProperties runProperties = new RunProperties();

            if (isHeader)
            {
                runProperties.AppendChild(new Bold());
            }

            // Размер текста
            var fontSize = new DocumentFormat.OpenXml.Wordprocessing.FontSize();
            fontSize.Val = "24"; // 12pt
            runProperties.AppendChild(fontSize);

            run.PrependChild(runProperties);
            paragraph.AppendChild(run);

            // Настройки ячейки
            TableCellProperties cellProperties = new TableCellProperties();

            // Границы ячейки
            cellProperties.AppendChild(new TableCellWidth() { Type = TableWidthUnitValues.Auto });

            // Отступы внутри ячейки
            cellProperties.AppendChild(new TableCellMargin()
            {
                TopMargin = new TopMargin() { Width = "100" },
                BottomMargin = new BottomMargin() { Width = "100" },
                LeftMargin = new LeftMargin() { Width = "100" },
                RightMargin = new RightMargin() { Width = "100" }
            });

            // Фон для заголовков
            if (hasBackground)
            {
                Shading shading = new Shading()
                {
                    Val = ShadingPatternValues.Clear,
                    Fill = "D9D9D9" // Светло-серый
                };
                cellProperties.AppendChild(shading);
            }

            cell.AppendChild(cellProperties);
            cell.AppendChild(paragraph);

            return cell;
        }
    }
}