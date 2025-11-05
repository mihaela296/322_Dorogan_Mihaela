using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class UsersManagementPage : Page
    {
        private User _currentAdmin;

        public UsersManagementPage(User admin)
        {
            InitializeComponent();
            _currentAdmin = admin;
            this.Loaded += (s, e) => LoadUsers();
        }

        private void LoadUsers()
        {
            try
            {
                using (var db = new DEntities())
                {
                    var users = db.Users.AsQueryable();
                    users = ApplyFilters(users);
                    users = ApplySorting(users);

                    if (DgUsers != null)
                    {
                        DgUsers.ItemsSource = users.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки пользователей: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private IQueryable<User> ApplyFilters(IQueryable<User> users)
        {
            if (TbSearch != null && !string.IsNullOrWhiteSpace(TbSearch.Text))
            {
                var searchText = TbSearch.Text.ToLower();
                users = users.Where(u =>
                    u.FIO.ToLower().Contains(searchText) ||
                    u.Login.ToLower().Contains(searchText));
            }

            if (CbRoleFilter?.SelectedItem is ComboBoxItem roleItem && roleItem.Content.ToString() != "Все")
            {
                users = users.Where(u => u.Role == roleItem.Content.ToString());
            }

            return users;
        }

        private IQueryable<User> ApplySorting(IQueryable<User> users)
        {
            if (CbSort == null || CbSort.SelectedIndex < 0)
                return users.OrderBy(u => u.FIO);

            return CbSort.SelectedIndex switch
            {
                0 => users.OrderBy(u => u.FIO),
                1 => users.OrderByDescending(u => u.FIO),
                2 => users.OrderByDescending(u => u.ID),
                _ => users.OrderBy(u => u.FIO)
            };
        }

        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadUsers();
        }

        private void CbRoleFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadUsers();
        }

        private void CbSort_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadUsers();
        }

        private void BtnClearFilters_Click(object sender, RoutedEventArgs e)
        {
            if (TbSearch != null) TbSearch.Clear();
            if (CbRoleFilter != null) CbRoleFilter.SelectedIndex = 0;
            if (CbSort != null) CbSort.SelectedIndex = 0;
            LoadUsers();
        }

        private void BtnAddUser_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddEditUserPage(null, _currentAdmin));
        }

        private void BtnEditUser_Click(object sender, RoutedEventArgs e)
        {
            var user = (sender as Button)?.DataContext as User;
            if (user != null)
            {
                NavigationService.Navigate(new AddEditUserPage(user, _currentAdmin));
            }
        }

        private void BtnDeleteUser_Click(object sender, RoutedEventArgs e)
        {
            var user = (sender as Button)?.DataContext as User;
            if (user != null)
            {
                if (user.ID == _currentAdmin.ID)
                {
                    MessageBox.Show("Нельзя удалить собственный аккаунт!", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var result = MessageBox.Show(
                    $"Вы уверены, что хотите удалить пользователя:\n{user.FIO} ({user.Login})?",
                    "Подтверждение удаления",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        using (var db = new DEntities())
                        {
                            var hasPayments = db.Payments.Any(p => p.UserID == user.ID);
                            if (hasPayments)
                            {
                                MessageBox.Show("Нельзя удалить пользователя с существующими платежами!", "Ошибка",
                                    MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }

                            var userToDelete = db.Users.Find(user.ID);
                            if (userToDelete != null)
                            {
                                db.Users.Remove(userToDelete);
                                db.SaveChanges();
                                LoadUsers();
                                MessageBox.Show("Пользователь успешно удален!", "Успех",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка удаления пользователя: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void BtnResetPassword_Click(object sender, RoutedEventArgs e)
        {
            var user = (sender as Button)?.DataContext as User;
            if (user != null)
            {
                var result = MessageBox.Show(
                    $"Сбросить пароль пользователя {user.FIO} на '123456'?",
                    "Сброс пароля",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        using (var db = new DEntities())
                        {
                            var userToUpdate = db.Users.Find(user.ID);
                            if (userToUpdate != null)
                            {
                                userToUpdate.Password = GetHash("123456");
                                db.SaveChanges();
                                MessageBox.Show("Пароль успешно сброшен на '123456'", "Успех",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка сброса пароля: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private string GetHash(string input)
        {
            using (var sha1 = System.Security.Cryptography.SHA1.Create())
            {
                var hash = sha1.ComputeHash(System.Text.Encoding.UTF8.GetBytes(input));
                return string.Concat(hash.Select(b => b.ToString("X2")));
            }
        }

        private void BtnExportUsersExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel();
        }

        private void BtnExportUsersWord_Click(object sender, RoutedEventArgs e)
        {
            ExportToWord();
        }

        private void ExportToExcel()
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Пользователи_{DateTime.Now:yyyyMMdd_HHmmss}",
                    DefaultExt = ".csv"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // Всегда сохраняем как CSV для надежности
                    string csvFilePath = Path.ChangeExtension(saveDialog.FileName, ".csv");
                    ExportToCsv(csvFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка");
            }
        }

        private void ExportToCsv(string filePath)
        {
            try
            {
                var users = GetUsersForExport();
                if (users.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта", "Информация");
                    return;
                }

                var csvLines = new List<string>
                {
                    "ID;Логин;ФИО;Роль"
                };

                foreach (var user in users)
                {
                    csvLines.Add(
                        $"{user.ID};" +
                        $"{EscapeCsvField(user.Login)};" +
                        $"{EscapeCsvField(user.FIO)};" +
                        $"{user.Role}"
                    );
                }

                File.WriteAllLines(filePath, csvLines, System.Text.Encoding.UTF8);

                MessageBox.Show($"Данные экспортированы успешно!\nФайл: {filePath}\n\nФайл откроется в Excel автоматически.",
                    "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);

                // Открываем файл в ассоциированной программе (Excel)
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}\n\nПроверьте:\n• Доступ к папке\n• Закрыт ли файл в другой программе",
                    "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string EscapeCsvField(string field)
        {
            if (field.Contains(";") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        private void ExportToWord()
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Word documents (*.docx)|*.docx",
                    FileName = $"Пользователи_{DateTime.Now:yyyyMMdd_HHmmss}",
                    DefaultExt = ".docx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    Word.Application wordApp = null;
                    Word.Document wordDoc = null;

                    try
                    {
                        var users = GetUsersForExport();
                        if (users.Count == 0)
                        {
                            MessageBox.Show("Нет данных для экспорта", "Информация");
                            return;
                        }

                        wordApp = new Word.Application();
                        wordDoc = wordApp.Documents.Add();
                        wordApp.Visible = false;

                        // Заголовок
                        Word.Paragraph title = wordDoc.Paragraphs.Add();
                        title.Range.Text = "ОТЧЕТ ПО ПОЛЬЗОВАТЕЛЯМ";
                        title.Range.Font.Bold = 1;
                        title.Range.Font.Size = 16;
                        title.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        title.Range.InsertParagraphAfter();

                        // Информация о фильтрах
                        Word.Paragraph info = wordDoc.Paragraphs.Add();
                        info.Range.Text = GetFilterInfo();
                        info.Range.Font.Size = 12;
                        info.Range.InsertParagraphAfter();

                        wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                        // Таблица
                        if (users.Count > 0)
                        {
                            Word.Table table = wordDoc.Tables.Add(
                                wordDoc.Paragraphs.Add().Range,
                                users.Count + 1,
                                4);

                            table.Borders.Enable = 1;
                            table.Rows[1].Range.Font.Bold = 1;

                            // Заголовки таблицы
                            string[] headers = { "ID", "Логин", "ФИО", "Роль" };
                            for (int i = 0; i < headers.Length; i++)
                            {
                                table.Cell(1, i + 1).Range.Text = headers[i];
                                table.Cell(1, i + 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }

                            // Данные
                            int row = 2;
                            foreach (var user in users)
                            {
                                table.Cell(row, 1).Range.Text = user.ID.ToString();
                                table.Cell(row, 2).Range.Text = user.Login;
                                table.Cell(row, 3).Range.Text = user.FIO;
                                table.Cell(row, 4).Range.Text = user.Role;
                                row++;
                            }

                            // Статистика
                            wordDoc.Paragraphs.Add().Range.InsertParagraphAfter();
                            Word.Paragraph stats = wordDoc.Paragraphs.Add();
                            var adminCount = users.Count(u => u.Role == "Admin");
                            var userCount = users.Count(u => u.Role == "User");
                            stats.Range.Text = $"СТАТИСТИКА:\nВсего пользователей: {users.Count}\nАдминистраторов: {adminCount}\nОбычных пользователей: {userCount}";
                            stats.Range.Font.Bold = 1;
                            stats.Range.Font.Size = 12;

                            // Сохраняем
                            wordDoc.SaveAs2(saveDialog.FileName);
                        }

                        MessageBox.Show($"Отчет успешно создан!\n\nФайл: {saveDialog.FileName}",
                            "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);

                        // Показываем документ пользователю
                        wordApp.Visible = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при создании отчета: {ex.Message}", "Ошибка экспорта");
                    }
                    finally
                    {
                        // Не закрываем Word, чтобы пользователь увидел документ
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
            }
        }

        private string GetFilterInfo()
        {
            var filters = new List<string>();

            if (TbSearch != null && !string.IsNullOrWhiteSpace(TbSearch.Text))
            {
                filters.Add($"Поиск: {TbSearch.Text}");
            }

            if (CbRoleFilter?.SelectedItem is ComboBoxItem roleItem && roleItem.Content.ToString() != "Все")
            {
                filters.Add($"Роль: {roleItem.Content.ToString()}");
            }

            if (CbSort != null && CbSort.SelectedIndex >= 0)
            {
                var sortText = CbSort.SelectedItem.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "");
                filters.Add($"Сортировка: {sortText}");
            }

            filters.Add($"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}");

            return string.Join(" | ", filters);
        }

        private List<User> GetUsersForExport()
        {
            try
            {
                using (var db = new DEntities())
                {
                    var usersQuery = db.Users.AsQueryable();

                    // Применяем фильтры
                    if (TbSearch != null && !string.IsNullOrWhiteSpace(TbSearch.Text))
                    {
                        var searchText = TbSearch.Text.ToLower();
                        usersQuery = usersQuery.Where(u =>
                            u.FIO.ToLower().Contains(searchText) ||
                            u.Login.ToLower().Contains(searchText));
                    }

                    // Фильтр по роли
                    if (CbRoleFilter?.SelectedItem is ComboBoxItem roleItem && roleItem.Content.ToString() != "Все")
                    {
                        usersQuery = usersQuery.Where(u => u.Role == roleItem.Content.ToString());
                    }

                    // Применяем сортировку
                    if (CbSort == null || CbSort.SelectedIndex < 0)
                        usersQuery = usersQuery.OrderBy(u => u.FIO);
                    else
                    {
                        switch (CbSort.SelectedIndex)
                        {
                            case 0: usersQuery = usersQuery.OrderBy(u => u.FIO); break;
                            case 1: usersQuery = usersQuery.OrderByDescending(u => u.FIO); break;
                            case 2: usersQuery = usersQuery.OrderByDescending(u => u.ID); break;
                            default: usersQuery = usersQuery.OrderBy(u => u.FIO); break;
                        }
                    }

                    return usersQuery.ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка получения данных для экспорта: {ex.Message}");
                return new List<User>();
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadUsers();
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}