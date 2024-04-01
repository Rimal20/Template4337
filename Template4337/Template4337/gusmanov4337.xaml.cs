using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.IO;
using System.Globalization;

namespace Template4337
{
    /// <summary>
    /// Логика взаимодействия для gusmanov4337.xaml
    /// </summary>
    public partial class gusmanov4337 : Window
    {
        public gusmanov4337()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Excel.Application ObjWorkExcel = new
            Excel.Application();

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);

            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (ISRPO2Entities usersEntities = new ISRPO2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.Users.Add(new Users()
                    {
                        FullName = list[i, 0],
                        CodeClient = list[i, 1],
                        BirthDate = list[i, 2],
                        Index = list[i, 3],
                        City = list[i, 4],
                        Street = list[i, 5],
                        Home = list[i, 6],
                        Kvartira = list[i, 7],
                        E_mail = list[i, 8]
                    });
                }
                usersEntities.SaveChanges();
            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<Users> users;
            using (ISRPO2Entities usersEntities = new ISRPO2Entities())
            {
                users = usersEntities.Users.ToList().OrderBy(s => s.E_mail).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            Excel.Worksheet category1Sheet = app.Worksheets.Item[1];
            category1Sheet.Name = "Категория 1 (от 20 до 29)";

            Excel.Worksheet category2Sheet = app.Worksheets.Item[2];
            category2Sheet.Name = "Категория 2 (от 30 до 39)";

            Excel.Worksheet category3Sheet = app.Worksheets.Item[3];
            category3Sheet.Name = "Категория 3 (от 40 и выше)";

            foreach (var user in users)
            {
                Excel.Worksheet worksheet = null;

                // Преобразовываем строку birthdate в тип DateTime
                DateTime birthDate;
                if (!DateTime.TryParse(user.BirthDate, out birthDate))
                {
                    // Если преобразование не удалось, пропускаем этого пользователя
                    continue;
                }

                // Вычисляем возраст пользователя по дате рождения
                DateTime currentDate = DateTime.Now;
                int age = currentDate.Year - birthDate.Year;
                if (currentDate.Month < birthDate.Month || (currentDate.Month == birthDate.Month && currentDate.Day < birthDate.Day))
                {
                    age--;
                }

                // Определяем категорию возраста для каждого пользователя
                if (age >= 20 && age <= 29)
                    worksheet = category1Sheet;
                else if (age >= 30 && age <= 39)
                    worksheet = category2Sheet;
                else if (age >= 40)
                    worksheet = category3Sheet;

                // Если не удалось определить категорию, пропускаем этого пользователя
                if (worksheet == null)
                    continue;

                // Находим следующую доступную строку в листе
                int lastRow = worksheet.Cells[worksheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
                int newRow = lastRow + 1;

                // Записываем данные о пользователе в соответствующие столбцы
                worksheet.Cells[newRow, 1] = user.ID;
                worksheet.Cells[newRow, 2] = user.FullName;
                worksheet.Cells[newRow, 3] = user.E_mail;
            }

            // Автоматически подгоняем ширину столбцов
            category1Sheet.Columns.AutoFit();
            category2Sheet.Columns.AutoFit();
            category3Sheet.Columns.AutoFit();

            // Делаем Excel видимым
            app.Visible = true;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "JSON файлы (*.json)|*.json|Все файлы (*.*)|*.*",
                Title = "Выберите файл JSON для добавления в базу данных"
            };
            if (!(openFileDialog.ShowDialog() == true))
                return;
            {
                string jsonFilePath = openFileDialog.FileName;
                List<Users> users = JsonConvert.DeserializeObject<List<Users>>(File.ReadAllText(jsonFilePath));
                using (ISRPO2Entities isrpoEntities = new ISRPO2Entities())
                {
                    foreach (var Users in users)
                    {
                        isrpoEntities.Users.Add(new Users()
                        {
                            FullName = Users.FullName,
                            CodeClient = Users.CodeClient,
                            BirthDate = Users.BirthDate,
                            Index = Users.Index,
                            City = Users.City,
                            Street = Users.Street,
                            Home = Users.Home,
                            Kvartira = Users.Kvartira,
                            E_mail = Users.E_mail
                        });
                    }

                    isrpoEntities.SaveChanges();
                    this.Close();
                    MessageBox.Show("Импорт завершен", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog()
            {
                DefaultExt = "*.docx",
                Filter = "Документ Word (*.docx)|*.docx|Все файлы (*.*)|*.*",
                Title = "Выберите место сохранения файла Word"
            };

            try
            {
                if (sfd.ShowDialog() == true)
                {
                    string outputFilePath = sfd.FileName;

                    using (DocX document = DocX.Create(outputFilePath))
                    {
                        using (ISRPO2Entities isrpoEntities = new ISRPO2Entities())
                        {
                            // Получаем список всех пользователей
                            var allUsers = isrpoEntities.Users.ToList();

                            // Группируем пользователей по возрастным категориям
                            var category1 = allUsers.Where(u => CalculateAge(u.BirthDate) >= 20 && CalculateAge(u.BirthDate) <= 29).ToList();
                            var category2 = allUsers.Where(u => CalculateAge(u.BirthDate) >= 30 && CalculateAge(u.BirthDate) <= 39).ToList();
                            var category3 = allUsers.Where(u => CalculateAge(u.BirthDate) >= 40).ToList();

                            // Вставляем данные на первую страницу
                            InsertDataIntoWordSheet(document, category1, "Категория 1 (от 20 до 29)");

                            // Вставляем данные на вторую страницу
                            document.InsertSectionPageBreak();
                            InsertDataIntoWordSheet(document, category2, "Категория 2 (от 30 до 39)");

                            // Вставляем данные на третью страницу
                            document.InsertSectionPageBreak();
                            InsertDataIntoWordSheet(document, category3, "Категория 3 (от 40 и выше)");
                        }

                        document.Save();
                    }

                    this.Close();
                    MessageBox.Show("Данные успешно сохранены в файл Word.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int CalculateAge(string birthDate)
        {
            DateTime today = DateTime.Today;
            if (DateTime.TryParseExact(birthDate, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime birth))
            {
                int age = today.Year - birth.Year;
                if (today.Month < birth.Month || (today.Month == birth.Month && today.Day < birth.Day))
                {
                    age--;
                }
                return age;
            }
            else
            {
                // Если не удалось распознать дату, возвращаем некорректный возраст
                return -1;
            }
        }

        // Метод для вставки данных в документ Word
        private void InsertDataIntoWordSheet(DocX document, List<Users> data, string categoryTitle)
        {
            if (data.Count == 0)
                return;

            document.InsertParagraph(categoryTitle).FontSize(14).Bold().Alignment = Alignment.center;

            // Создаем таблицу с тремя столбцами
            Xceed.Document.NET.Table table = document.AddTable(data.Count + 1, 3);

            // Заполняем заголовки столбцов
            table.Rows[0].Cells[0].Paragraphs.First().Append("Код клиента");
            table.Rows[0].Cells[1].Paragraphs.First().Append("ФИО");
            table.Rows[0].Cells[2].Paragraphs.First().Append("Email");

            // Заполняем данные
            for (int i = 0; i < data.Count; i++)
            {
                table.Rows[i + 1].Cells[0].Paragraphs.First().Append(data[i].CodeClient.ToString());
                table.Rows[i + 1].Cells[1].Paragraphs.First().Append(data[i].FullName);
                table.Rows[i + 1].Cells[2].Paragraphs.First().Append(data[i].E_mail);
            }

            document.InsertTable(table);
        }


    }
}
