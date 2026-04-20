using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.EntityFrameworkCore;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace Group4337
{
    /// <summary>
    /// Логика взаимодействия для _4337_Соловьев.xaml
    /// Вариант 7: Клиенты, категории по возрасту
    /// </summary>
    public partial class _4337_Соловьев : Window
    {
        public _4337_Соловьев()
        {
            InitializeComponent();
            LoadFromDatabase();
        }

        // ==================== ЗАГРУЗКА ИЗ БД ====================

        private void LoadFromDatabase()
        {
            try
            {
                using var ctx = new ClientDbContext();
                ctx.Database.EnsureCreated();
                var clients = ctx.Clients.ToList();
                ClientsDataGrid.ItemsSource = clients;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки из БД: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadFromDb_Click(object sender, RoutedEventArgs e)
        {
            LoadFromDatabase();
            MessageBox.Show("Данные загружены из базы данных!", "Успех",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // ==================== ИМПОРТ ИЗ EXCEL ====================

        /// <summary>
        /// Парсит дату из ячейки Excel (поддержка DateTime, числового serial и строковых форматов)
        /// </summary>
        private DateTime ParseDate(IXLCell cell)
        {
            var val = cell.Value;
            if (val.IsDateTime)
                return val.GetDateTime();
            if (val.IsNumber)
                return DateTime.FromOADate(val.GetNumber());
            var str = val.ToString().Trim();
            var formats = new[] { "dd.MM.yyyy", "d.M.yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MM-yyyy" };
            if (DateTime.TryParseExact(str, formats,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var dtExact))
                return dtExact;
            if (DateTime.TryParse(str, out var dt))
                return dt;
            throw new FormatException($"Не удалось распознать дату: '{str}'");
        }

        /// <summary>
        /// Безопасно читает строковое значение из ячейки Excel (любой тип ячейки)
        /// </summary>
        private string SafeGetString(IXLCell cell)
        {
            var val = cell.Value;
            if (val.IsNumber)
            {
                // Целые числа (Индекс, Дом, Квартира) — без десятичной части
                var num = val.GetNumber();
                return ((long)num).ToString();
            }
            return val.ToString().Trim();
        }

        private void ImportExcel(string path)
        {
            using var ctx = new ClientDbContext();
            ctx.Database.EnsureCreated();

            using var workbook = new XLWorkbook(path);
            var sheet = workbook.Worksheet(1);
            var range = sheet.RangeUsed();
            if (range == null) return;

            // Читаем заголовки для автоопределения формата
            var headerRow = range.Row(1);
            var firstHeader = SafeGetString(headerRow.Cell(1)).ToLower();

            // Определяем смещение колонок
            // Формат import.xlsx: Id | ClientCode | FullName | BirthDate | PostalCode | City | Street | House | Apartment | Email
            // Формат 3.xlsx:      ФИО | Код клиента | Дата рождения | Индекс | Город | Улица | Дом | Квартира | E-mail
            bool hasIdColumn = firstHeader == "id" || firstHeader.Contains("id");
            int offset = hasIdColumn ? 1 : 0; // Сдвиг если есть колонка Id

            var rows = range.RowsUsed().Skip(1); // Пропускаем заголовок

            foreach (var row in rows)
            {
                Client client;
                if (hasIdColumn)
                {
                    // Формат: Id, ClientCode, FullName, BirthDate, PostalCode, City, Street, House, Apartment, Email
                    client = new Client
                    {
                        ClientCode = SafeGetString(row.Cell(2)),
                        FullName = SafeGetString(row.Cell(3)),
                        BirthDate = ParseDate(row.Cell(4)),
                        PostalCode = SafeGetString(row.Cell(5)),
                        City = SafeGetString(row.Cell(6)),
                        Street = SafeGetString(row.Cell(7)),
                        House = SafeGetString(row.Cell(8)),
                        Apartment = SafeGetString(row.Cell(9)),
                        Email = SafeGetString(row.Cell(10))
                    };
                }
                else
                {
                    // Формат 3.xlsx: ФИО, Код клиента, Дата рождения, Индекс, Город, Улица, Дом, Квартира, E-mail
                    client = new Client
                    {
                        FullName = SafeGetString(row.Cell(1)),
                        ClientCode = SafeGetString(row.Cell(2)),
                        BirthDate = ParseDate(row.Cell(3)),
                        PostalCode = SafeGetString(row.Cell(4)),
                        City = SafeGetString(row.Cell(5)),
                        Street = SafeGetString(row.Cell(6)),
                        House = SafeGetString(row.Cell(7)),
                        Apartment = SafeGetString(row.Cell(8)),
                        Email = SafeGetString(row.Cell(9))
                    };
                }
                ctx.Clients.Add(client);
            }
            ctx.SaveChanges();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx"
            };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    ImportExcel(dialog.FileName);
                    LoadFromDatabase();
                    MessageBox.Show("Данные успешно импортированы из Excel!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка импорта: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // ==================== ЭКСПОРТ В EXCEL ====================

        /// <summary>
        /// Записывает лист Excel для одной категории клиентов.
        /// Формат: Код клиента, ФИО, E-mail, Возраст
        /// </summary>
        private void WriteSheet(XLWorkbook workbook, string name, List<Client> data)
        {
            var ws = workbook.Worksheets.Add(name);
            ws.Cell(1, 1).Value = "Код клиента";
            ws.Cell(1, 2).Value = "ФИО";
            ws.Cell(1, 3).Value = "E-mail";
            ws.Cell(1, 4).Value = "Возраст";
            ws.Cell(1, 5).Value = "Дата рождения";
            ws.Cell(1, 6).Value = "Город";
            ws.Cell(1, 7).Value = "Улица";

            // Жирный заголовок
            var headerRange = ws.Range(1, 1, 1, 7);
            headerRange.Style.Font.Bold = true;

            int row = 2;
            foreach (var c in data)
            {
                ws.Cell(row, 1).Value = c.ClientCode;
                ws.Cell(row, 2).Value = c.FullName;
                ws.Cell(row, 3).Value = c.Email;
                ws.Cell(row, 4).Value = c.Age;
                ws.Cell(row, 5).Value = c.BirthDate.ToString("dd.MM.yyyy");
                ws.Cell(row, 6).Value = c.City;
                ws.Cell(row, 7).Value = c.Street;
                row++;
            }

            ws.Columns().AdjustToContents();
        }

        private void ExportExcel(string path)
        {
            using var ctx = new ClientDbContext();
            var clients = ctx.Clients.ToList();

            // Категории по возрасту (Вариант 7)
            var cat1 = clients.Where(c => c.Age >= 20 && c.Age <= 29).ToList();
            var cat2 = clients.Where(c => c.Age >= 30 && c.Age <= 39).ToList();
            var cat3 = clients.Where(c => c.Age >= 40).ToList();

            using var wb = new XLWorkbook();
            WriteSheet(wb, "Категория 1 (20–29)", cat1);
            WriteSheet(wb, "Категория 2 (30–39)", cat2);
            WriteSheet(wb, "Категория 3 (40+)", cat3);
            wb.SaveAs(path);
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = "Клиенты_по_возрасту"
            };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    ExportExcel(dialog.FileName);
                    MessageBox.Show("Экспорт в Excel завершён!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // ==================== ИМПОРТ ИЗ JSON ====================

        private void ImportJson(string path)
        {
            using var ctx = new ClientDbContext();
            ctx.Database.EnsureCreated();

            var json = File.ReadAllText(path);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            var clients = JsonSerializer.Deserialize<List<Client>>(json, options);
            if (clients == null) return;

            foreach (var c in clients)
            {
                ctx.Clients.Add(new Client
                {
                    ClientCode = c.ClientCode,
                    FullName = c.FullName,
                    BirthDate = c.BirthDate,
                    PostalCode = c.PostalCode,
                    City = c.City,
                    Street = c.Street,
                    House = c.House,
                    Apartment = c.Apartment,
                    Email = c.Email
                });
            }
            ctx.SaveChanges();
        }

        private void Import_Json_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON (*.json)|*.json"
            };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    ImportJson(dialog.FileName);
                    LoadFromDatabase();
                    MessageBox.Show("Данные успешно импортированы из JSON!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка импорта JSON: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // ==================== ЭКСПОРТ В WORD ====================

        private TableCell MakeCell(string text, bool bold = false)
        {
            var run = bold
                ? new Run(new RunProperties(new Bold()), new Text(text))
                : new Run(new Text(text));
            return new TableCell(new Paragraph(run));
        }

        /// <summary>
        /// Добавляет страницу Word с таблицей клиентов одной категории.
        /// Формат: Код клиента, ФИО, E-mail, Возраст
        /// </summary>
        private void AddWordPage(Body body, string title, List<Client> data, bool addPageBreak)
        {
            // Заголовок категории
            body.AppendChild(new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(new RunProperties(new Bold(), new FontSize { Val = "28" }), new Text(title))
            ));

            var table = new Table();
            table.AppendChild(new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                )
            ));

            // Заголовки таблицы
            var headerRow = new TableRow();
            foreach (var h in new[] { "Код клиента", "ФИО", "E-mail", "Возраст", "Дата рождения" })
                headerRow.AppendChild(MakeCell(h, bold: true));
            table.AppendChild(headerRow);

            // Данные
            foreach (var c in data)
            {
                var row = new TableRow();
                foreach (var val in new[]
                {
                    c.ClientCode,
                    c.FullName,
                    c.Email,
                    c.Age.ToString(),
                    c.BirthDate.ToString("dd.MM.yyyy")
                })
                    row.AppendChild(MakeCell(val));
                table.AppendChild(row);
            }

            body.AppendChild(table);

            if (addPageBreak)
                body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private void ExportWord(string path)
        {
            using var ctx = new ClientDbContext();
            var clients = ctx.Clients.ToList();

            // Категории по возрасту (Вариант 7)
            var cat1 = clients.Where(c => c.Age >= 20 && c.Age <= 29).ToList();
            var cat2 = clients.Where(c => c.Age >= 30 && c.Age <= 39).ToList();
            var cat3 = clients.Where(c => c.Age >= 40).ToList();

            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            AddWordPage(body, "Категория 1 — возраст от 20 до 29", cat1, addPageBreak: true);
            AddWordPage(body, "Категория 2 — возраст от 30 до 39", cat2, addPageBreak: true);
            AddWordPage(body, "Категория 3 — возраст от 40", cat3, addPageBreak: false);

            mainPart.Document.Save();
        }

        private void Export_Word_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Word (*.docx)|*.docx",
                FileName = "Клиенты_по_возрасту"
            };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    ExportWord(dialog.FileName);
                    MessageBox.Show("Экспорт в Word завершён!", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка экспорта в Word: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
