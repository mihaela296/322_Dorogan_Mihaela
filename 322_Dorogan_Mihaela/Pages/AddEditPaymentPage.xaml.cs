using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class AddEditPaymentPage : Page
    {
        private User _currentUser;
        private Payment _editingPayment;
        private bool _isNewPayment;

        public AddEditPaymentPage(Payment payment, User currentUser)
        {
            InitializeComponent();
            _currentUser = currentUser;
            _editingPayment = payment ?? new Payment();
            _isNewPayment = payment == null;

            InitializeForm();
        }

        private void InitializeForm()
        {
            if (_isNewPayment)
            {
                TbTitle.Text = "ДОБАВЛЕНИЕ НОВОГО ПЛАТЕЖА";
                _editingPayment.Date = DateTime.Now;
            }
            else
            {
                TbTitle.Text = $"РЕДАКТИРОВАНИЕ ПЛАТЕЖА: {_editingPayment.Name}";
            }

            DataContext = _editingPayment;
            LoadComboBoxData();
            UpdateTotalAmount();
        }

        private void LoadComboBoxData()
        {
            try
            {
                using (var db = new Entities())
                {
                    // Загрузка пользователей
                    var users = db.Users.OrderBy(u => u.FIO).ToList();
                    CbUser.ItemsSource = users;

                    // Загрузка категорий
                    var categories = db.Categories.OrderBy(c => c.Name).ToList();
                    CbCategory.ItemsSource = categories;

                    // Для нового платежа устанавливаем текущего пользователя по умолчанию
                    if (_isNewPayment && _currentUser != null)
                    {
                        CbUser.SelectedValue = _currentUser.ID;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Ошибка загрузки данных: {ex.Message}");
            }
        }

        private void UpdateTotalAmount()
        {
            if (_editingPayment.Num > 0 && _editingPayment.Price > 0)
            {
                var total = _editingPayment.Num * _editingPayment.Price;
                TbTotal.Text = $"{total:N2} руб.";
            }
            else
            {
                TbTotal.Text = "0.00 руб.";
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
                    if (_isNewPayment)
                    {
                        db.Payments.Add(_editingPayment);
                    }
                    else
                    {
                        var existingPayment = db.Payments.Find(_editingPayment.ID);
                        if (existingPayment != null)
                        {
                            existingPayment.Date = _editingPayment.Date;
                            existingPayment.UserID = _editingPayment.UserID;
                            existingPayment.CategoryID = _editingPayment.CategoryID;
                            existingPayment.Name = _editingPayment.Name;
                            existingPayment.Num = _editingPayment.Num;
                            existingPayment.Price = _editingPayment.Price;
                        }
                    }

                    db.SaveChanges();

                    MessageBox.Show(_isNewPayment ? "Платеж успешно добавлен!" : "Платеж успешно обновлен!",
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

            if (_editingPayment.Date == null)
            {
                ShowError("Выберите дату платежа!");
                DpDate.Focus();
                return false;
            }

            if (_editingPayment.UserID == 0)
            {
                ShowError("Выберите пользователя!");
                CbUser.Focus();
                return false;
            }

            if (_editingPayment.CategoryID == 0)
            {
                ShowError("Выберите категорию!");
                CbCategory.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(_editingPayment.Name))
            {
                ShowError("Введите название платежа!");
                TbName.Focus();
                return false;
            }

            if (_editingPayment.Num <= 0)
            {
                ShowError("Количество должно быть больше 0!");
                TbNum.Focus();
                TbNum.SelectAll();
                return false;
            }

            if (_editingPayment.Price <= 0)
            {
                ShowError("Цена должна быть больше 0!");
                TbPrice.Focus();
                TbPrice.SelectAll();
                return false;
            }

            // Проверка на разумные пределы
            if (_editingPayment.Num > 1000000)
            {
                ShowError("Количество слишком большое!");
                TbNum.Focus();
                TbNum.SelectAll();
                return false;
            }

            if (_editingPayment.Price > 1000000000)
            {
                ShowError("Цена слишком большая!");
                TbPrice.Focus();
                TbPrice.SelectAll();
                return false;
            }

            return true;
        }

        private void TbNum_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateTotalAmount();
        }

        private void TbPrice_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateTotalAmount();
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