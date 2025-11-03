using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class UsersManagementPage : Page
    {
        private User _currentAdmin;

        public UsersManagementPage(User admin)
        {
            InitializeComponent();
            _currentAdmin = admin;

            // Загружаем данные после полной инициализации страницы
            this.Loaded += (s, e) => LoadUsers();
        }

        private void LoadUsers()
        {
            try
            {
                using (var db = new Entities())
                {
                    var users = db.Users.AsQueryable();

                    // Применение фильтров
                    users = ApplyFilters(users);

                    // Применение сортировки
                    users = ApplySorting(users);

                    // Проверяем, что DgUsers инициализирован
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
            // Проверяем, что элементы управления инициализированы
            if (TbSearch != null && !string.IsNullOrWhiteSpace(TbSearch.Text))
            {
                var searchText = TbSearch.Text.ToLower();
                users = users.Where(u =>
                    u.FIO.ToLower().Contains(searchText) ||
                    u.Login.ToLower().Contains(searchText));
            }

            // Фильтр по роли
            if (CbRoleFilter?.SelectedItem is ComboBoxItem roleItem && roleItem.Content.ToString() != "Все")
            {
                users = users.Where(u => u.Role == roleItem.Content.ToString());
            }

            return users;
        }

        private IQueryable<User> ApplySorting(IQueryable<User> users)
        {
            // Проверяем, что ComboBox инициализирован
            if (CbSort == null || CbSort.SelectedIndex < 0)
                return users.OrderBy(u => u.FIO);

            return CbSort.SelectedIndex switch
            {
                0 => users.OrderBy(u => u.FIO), // По ФИО (А-Я)
                1 => users.OrderByDescending(u => u.FIO), // По ФИО (Я-А)
                2 => users.OrderByDescending(u => u.ID), // По дате регистрации
                _ => users.OrderBy(u => u.FIO)
            };
        }

        // Остальные методы остаются без изменений...
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
                // Нельзя удалить самого себя
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
                        using (var db = new Entities())
                        {
                            // Проверяем есть ли связанные платежи
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
                        using (var db = new Entities())
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

        private void BtnExportUsers_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Пользователи_{DateTime.Now:yyyyMMdd_HHmmss}",
                    DefaultExt = ".xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // Простая визуальная обратная связь
                    string originalText = BtnExportUsers.Content.ToString();
                    BtnExportUsers.Content = "Экспорт...";
                    BtnExportUsers.IsEnabled = false;

                    try
                    {
                        // Получаем данные
                        var users = GetUsersForExport();
                        if (users == null || users.Count == 0)
                        {
                            MessageBox.Show("Нет данных для экспорта", "Информация");
                            return;
                        }

                        // Создаем Excel пакет
                        using (var excelPackage = new OfficeOpenXml.ExcelPackage())
                        {
                            var worksheet = excelPackage.Workbook.Worksheets.Add("Пользователи");

                            // Простые заголовки
                            worksheet.Cells[1, 1].Value = "ID";
                            worksheet.Cells[1, 2].Value = "Логин";
                            worksheet.Cells[1, 3].Value = "ФИО";
                            worksheet.Cells[1, 4].Value = "Роль";

                            // Делаем заголовки жирными
                            for (int i = 1; i <= 4; i++)
                            {
                                worksheet.Cells[1, i].Style.Font.Bold = true;
                            }

                            // Заполняем данные
                            int row = 2;
                            foreach (var user in users)
                            {
                                worksheet.Cells[row, 1].Value = user.ID;
                                worksheet.Cells[row, 2].Value = user.Login;
                                worksheet.Cells[row, 3].Value = user.FIO;
                                worksheet.Cells[row, 4].Value = user.Role;
                                row++;
                            }

                            // Автоподбор ширины столбцов
                            worksheet.Cells[1, 1, row, 4].AutoFitColumns();

                            // Сохраняем файл
                            FileInfo fileInfo = new FileInfo(saveDialog.FileName);
                            excelPackage.SaveAs(fileInfo);
                        }

                        MessageBox.Show($"Данные экспортированы успешно!\n\nФайл: {saveDialog.FileName}",
                            "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}\n\n" +
                            "Убедитесь, что:\n" +
                            "• Файл не открыт в другой программе\n" +
                            "• У вас есть права на запись в выбранную папку\n" +
                            "• Имя файла не содержит запрещенных символов",
                            "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    finally
                    {
                        // Восстанавливаем кнопку
                        BtnExportUsers.Content = originalText;
                        BtnExportUsers.IsEnabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка");
                BtnExportUsers.Content = "Экспорт в Excel";
                BtnExportUsers.IsEnabled = true;
            }
        }

        // Новый метод для получения пользователей для экспорта
        private List<User> GetUsersForExport()
{
    try
    {
        using (var db = new Entities())
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