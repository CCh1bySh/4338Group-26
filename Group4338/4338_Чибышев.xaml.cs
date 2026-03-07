using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using OfficeOpenXml;
using Microsoft.Win32;

namespace Group4338
{
    public partial class _4338_Чибышев : Window
    {
        // Простая модель данных
        public class Employee
        {
            public int Id { get; set; }
            public string Login { get; set; }
            public string Password { get; set; }
            public string Role { get; set; }
        }

        private List<Employee> employees = new List<Employee>();

        public _4338_Чибышев()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialOrganization("Study");
        }

        // ИМПОРТ ДАННЫХ
        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel файлы|*.xlsx";

                if (openFile.ShowDialog() == true)
                {
                    employees.Clear();

                    using (var package = new ExcelPackage(new FileInfo(openFile.FileName)))
                    {
                        var sheet = package.Workbook.Worksheets[0];
                        int rows = sheet.Dimension.Rows;

                        // Читаем данные (начиная со 2 строки, т.к. 1 - заголовки)
                        for (int row = 2; row <= rows; row++)
                        {
                            var login = sheet.Cells[row, 1].Text;
                            var password = sheet.Cells[row, 2].Text;
                            var role = sheet.Cells[row, 3].Text;

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

                    // Показываем данные
                    dgEmployees.ItemsSource = null;
                    dgEmployees.ItemsSource = employees;
                    txtStatus.Text = $"Загружено {employees.Count} записей";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        // ЭКСПОРТ В EXCEL
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (employees.Count == 0)
                {
                    MessageBox.Show("Сначала импортируйте данные");
                    return;
                }

                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel файлы|*.xlsx";
                saveFile.FileName = "Employees_by_Role.xlsx";

                if (saveFile.ShowDialog() == true)
                {
                    using (var package = new ExcelPackage())
                    {
                        // Группируем по ролям
                        var groups = employees.GroupBy(x => x.Role);

                        foreach (var group in groups)
                        {
                            // Создаем лист для каждой роли
                            string sheetName = group.Key.Length > 30 ? group.Key.Substring(0, 30) : group.Key;
                            var sheet = package.Workbook.Worksheets.Add(sheetName);

                            // Заголовки
                            sheet.Cells[1, 1].Value = "Логин";
                            sheet.Cells[1, 2].Value = "Пароль";

                            // Данные
                            int row = 2;
                            foreach (var emp in group)
                            {
                                sheet.Cells[row, 1].Value = emp.Login;
                                sheet.Cells[row, 2].Value = emp.Password;
                                row++;
                            }

                            // Автоподбор ширины
                            sheet.Cells.AutoFitColumns();
                        }

                        // Сохраняем
                        package.SaveAs(new FileInfo(saveFile.FileName));
                    }

                    txtStatus.Text = $"Экспортировано в {saveFile.FileName}";
                    MessageBox.Show("Готово!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
    }
}