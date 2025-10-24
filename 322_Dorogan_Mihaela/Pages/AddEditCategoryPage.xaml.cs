using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class AddEditCategoryPage : Page
    {
        private User _currentUser;
        private Category _editingCategory;
        private bool _isNewCategory;

        public AddEditCategoryPage(Category category, User currentUser)
        {
            InitializeComponent();
            _currentUser = currentUser;
            _editingCategory = category ?? new Category();
            _isNewCategory = category == null;

            InitializeForm();
        }

        private void InitializeForm()
        {
            if (_isNewCategory)
            {
                TbTitle.Text = "ДОБАВЛЕНИЕ НОВОЙ КАТЕГОРИИ";
            }
            else
            {
                TbTitle.Text = $"РЕДАКТИРОВАНИЕ КАТЕГОРИИ: {_editingCategory.Name}";
                LoadCategoryInfo();
            }

            DataContext = _editingCategory;
        }

        private void LoadCategoryInfo()
        {
            try
            {
                using (var db = new Entities())
                {
                    var category = db.Categories
                        .Include(c => c.Payments) // Исправлено на Payments
                        .FirstOrDefault(c => c.ID == _editingCategory.ID);

                    if (category != null)
                    {
                        var paymentCount = category.Payments.Count();
                        var totalAmount = category.Payments.Sum(p => p.Num * p.Price);
                        var lastPayment = category.Payments.OrderByDescending(p => p.Date).FirstOrDefault();

                        var info = new System.Text.StringBuilder();
                        info.AppendLine($"Количество платежей: {paymentCount}");
                        info.AppendLine($"Общая сумма: {totalAmount:N2} руб.");

                        if (lastPayment != null)
                        {
                            info.AppendLine($"Последний платеж: {lastPayment.Date:dd.MM.yyyy}");
                        }
                        else
                        {
                            info.AppendLine("Платежей нет");
                        }

                        TbCategoryInfo.Text = info.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                TbCategoryInfo.Text = $"Ошибка загрузки информации: {ex.Message}";
            }
        }

        private void ShowError(string message)
        {
            TbError.Text = message;
            TbError.Visibility = Visibility.Visible;
        }

        private void ClearError()
        {
            TbError.Text = string.Empty;
            TbError.Visibility = Visibility.Collapsed;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateForm())
                return;

            try
            {
                using (var db = new Entities())
                {
                    if (_isNewCategory)
                    {
                        // Проверка уникальности названия
                        if (db.Categories.Any(c => c.Name == _editingCategory.Name.Trim()))
                        {
                            ShowError("Категория с таким названием уже существует!");
                            TbName.Focus();
                            TbName.SelectAll();
                            return;
                        }

                        db.Categories.Add(_editingCategory);
                    }
                    else
                    {
                        var existingCategory = db.Categories.Find(_editingCategory.ID);
                        if (existingCategory != null)
                        {
                            // Проверка уникальности названия (исключая текущую категорию)
                            if (db.Categories.Any(c => c.Name == _editingCategory.Name.Trim() && c.ID != _editingCategory.ID))
                            {
                                ShowError("Категория с таким названием уже существует!");
                                TbName.Focus();
                                TbName.SelectAll();
                                return;
                            }

                            existingCategory.Name = _editingCategory.Name.Trim();
                        }
                    }

                    db.SaveChanges();

                    MessageBox.Show(_isNewCategory ? "Категория успешно добавлена!" : "Категория успешно обновлена!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                    NavigationService.GoBack();
                }
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка сохранения: {ex.Message}");
            }
        }

        private bool ValidateForm()
        {
            ClearError();

            if (string.IsNullOrWhiteSpace(_editingCategory.Name))
            {
                ShowError("Введите название категории!");
                TbName.Focus();
                return false;
            }

            if (_editingCategory.Name.Trim().Length < 2)
            {
                ShowError("Название категории должно содержать минимум 2 символа!");
                TbName.Focus();
                return false;
            }

            if (_editingCategory.Name.Trim().Length > 50)
            {
                ShowError("Название категории не должно превышать 50 символов!");
                TbName.Focus();
                return false;
            }

            return true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}