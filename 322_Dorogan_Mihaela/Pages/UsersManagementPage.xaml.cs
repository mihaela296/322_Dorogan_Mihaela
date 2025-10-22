using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Data.Entity;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class UsersManagementPage : Page
    {
        private User _currentAdmin;

        public UsersManagementPage(User admin)
        {
            InitializeComponent();
            _currentAdmin = admin;
            LoadUsers();
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

                    DgUsers.ItemsSource = users.ToList();
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
            // Поиск
            if (!string.IsNullOrWhiteSpace(TbSearch.Text))
            {
                var searchText = TbSearch.Text.ToLower();
                users = users.Where(u =>
                    u.FIO.ToLower().Contains(searchText) ||
                    u.Login.ToLower().Contains(searchText));
            }

            // Фильтр по роли
            if (CbRoleFilter.SelectedItem is ComboBoxItem roleItem && roleItem.Content.ToString() != "Все")
            {
                users = users.Where(u => u.Role == roleItem.Content.ToString());
            }

            return users;
        }

        private IQueryable<User> ApplySorting(IQueryable<User> users)
        {
            return CbSort.SelectedIndex switch
            {
                0 => users.OrderBy(u => u.FIO), // По ФИО (А-Я)
                1 => users.OrderByDescending(u => u.FIO), // По ФИО (Я-А)
                2 => users.OrderByDescending(u => u.ID), // По дате регистрации (предполагаем, что ID автоинкремент)
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
            TbSearch.Clear();
            CbRoleFilter.SelectedIndex = 0;
            CbSort.SelectedIndex = 0;
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
                    FileName = $"Пользователи_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // Здесь должна быть реализация экспорта в Excel
                    // Временная заглушка
                    MessageBox.Show("Функция экспорта в Excel будет реализована позже", "Информация",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
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