using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class CategoriesManagementPage : Page
    {
        private User _currentAdmin;

        public CategoriesManagementPage(User admin)
        {
            InitializeComponent();
            _currentAdmin = admin;
            LoadCategories();
        }

        private void LoadCategories()
        {
            try
            {
                using (var db = new Entities())
                {
                    var categories = db.Categories
                        .Include(c => c.Payments)
                        .AsQueryable();

                    // Применение поиска
                    if (!string.IsNullOrWhiteSpace(TbSearch.Text))
                    {
                        var searchText = TbSearch.Text.ToLower();
                        categories = categories.Where(c => c.Name.ToLower().Contains(searchText));
                    }

                    var categoryList = categories
                        .OrderBy(c => c.Name)
                        .ToList()
                        .Select(c => new
                        {
                            c.ID,
                            c.Name,
                            PaymentCount = c.Payments.Count
                        });

                    DgCategories.ItemsSource = categoryList;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки категорий: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadCategories();
        }

        private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
        {
            TbSearch.Clear();
            LoadCategories();
        }

        private void BtnAddCategory_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddEditCategoryPage(null, _currentAdmin));
        }

        private void BtnEditCategory_Click(object sender, RoutedEventArgs e)
        {
            var category = (sender as Button)?.DataContext;
            if (category != null)
            {
                dynamic cat = category;
                using (var db = new Entities())
                {
                    var categoryToEdit = db.Categories.Find(cat.ID);
                    if (categoryToEdit != null)
                    {
                        NavigationService.Navigate(new AddEditCategoryPage(categoryToEdit, _currentAdmin));
                    }
                }
            }
        }

        private void BtnDeleteCategory_Click(object sender, RoutedEventArgs e)
        {
            var category = (sender as Button)?.DataContext;
            if (category != null)
            {
                dynamic cat = category;

                // Проверяем есть ли связанные платежи
                if (cat.PaymentCount > 0)
                {
                    MessageBox.Show("Нельзя удалить категорию с привязанными платежами!", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var result = MessageBox.Show(
                    $"Вы уверены, что хотите удалить категорию:\n\"{cat.Name}\"?",
                    "Подтверждение удаления",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        using (var db = new Entities())
                        {
                            var categoryToDelete = db.Categories.Find(cat.ID);
                            if (categoryToDelete != null)
                            {
                                db.Categories.Remove(categoryToDelete);
                                db.SaveChanges();
                                LoadCategories();
                                MessageBox.Show("Категория успешно удалена!", "Успех",
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка удаления категории: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadCategories();
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}