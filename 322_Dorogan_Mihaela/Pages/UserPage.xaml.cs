using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.IO;
using System.Collections.Generic;
using System.Globalization;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class UserPage : Page
    {
        private User _currentUser;
        private bool _isInitialized = false;

        public UserPage(User user)
        {
            InitializeComponent();
            _currentUser = user;
            Loaded += UserPage_Loaded;
        }

        private void UserPage_Loaded(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized)
            {
                InitializePage();
                _isInitialized = true;
            }
        }

        private void InitializePage()
        {
            try
            {
                TbWelcome.Text = $"Добро пожаловать, {_currentUser.FIO}!";

                // Установка дат по умолчанию (текущий месяц)
                var startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                var endDate = DateTime.Now;

                DpStartDate.SelectedDate = startDate;
                DpEndDate.SelectedDate = endDate;

                // Сначала загружаем основные данные
                LoadCategories();
                LoadPayments();
                LoadStatistics();

                // Затем инициализируем диаграмму
                InitializeChart();

                // И только потом загружаем данные в диаграмму
                LoadChart();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации страницы: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void InitializeChart()
        {
            try
            {
                // Проверяем, что компонент диаграммы существует
                if (UserChart == null)
                {
                    return;
                }

                // Очистка предыдущей диаграммы
                UserChart.Series.Clear();
                UserChart.ChartAreas.Clear();
                UserChart.Titles.Clear();
                UserChart.Legends.Clear();

                // Создание области диаграммы
                ChartArea chartArea = new ChartArea();
                chartArea.AxisX.LabelStyle = new LabelStyle()
                {
                    Angle = -45,
                    Font = new System.Drawing.Font("Arial", 8)
                };
                chartArea.AxisY.LabelStyle = new LabelStyle()
                {
                    Font = new System.Drawing.Font("Arial", 8)
                };
                chartArea.AxisX.Interval = 1;
                UserChart.ChartAreas.Add(chartArea);

                // Легенда
                Legend legend = new Legend();
                legend.Font = new System.Drawing.Font("Arial", 9);
                legend.Docking = Docking.Bottom;
                UserChart.Legends.Add(legend);

                // Заголовок
                Title title = new Title("Мои расходы по категориям");
                title.Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold);
                UserChart.Titles.Add(title);
            }
            catch (Exception ex)
            {
                // Не показываем сообщение об ошибке здесь, т.к. это может быть из-за timing issue
                System.Diagnostics.Debug.WriteLine($"Ошибка инициализации диаграммы: {ex.Message}");
            }
        }

        private void LoadCategories()
        {
            try
            {
                using (var db = new Entities())
                {
                    var categories = db.Categories.OrderBy(c => c.Name).ToList();
                    CbCategory.ItemsSource = categories;
                    CbCategory.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки категорий: {ex.Message}");
            }
        }

        private void LoadPayments()
        {
            try
            {
                using (var db = new Entities())
                {
                    var paymentsQuery = db.Payments
                        .Include(p => p.Category)
                        .Where(p => p.UserID == _currentUser.ID);

                    // Фильтрация по дате
                    if (DpStartDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date >= DpStartDate.SelectedDate);
                    }

                    if (DpEndDate.SelectedDate != null)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.Date <= DpEndDate.SelectedDate);
                    }

                    // Фильтрация по категории - только если выбрана конкретная категория
                    if (CbCategory.SelectedItem != null && CbCategory.SelectedItem is Category selectedCategory)
                    {
                        paymentsQuery = paymentsQuery.Where(p => p.CategoryID == selectedCategory.ID);
                    }

                    var payments = paymentsQuery
                        .OrderByDescending(p => p.Date)
                        .ToList()
                        .Select(p => new
                        {
                            p.ID,
                            p.Date,
                            p.Name,
                            p.Category,
                            p.Num,
                            p.Price,
                            TotalAmount = p.Num * p.Price
                        });

                    DgPayments.ItemsSource = payments;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки платежей: {ex.Message}");
            }
        }

        private void LoadStatistics()
        {
            try
            {
                using (var db = new Entities())
                {
                    var totalSum = db.Payments
                        .Where(p => p.UserID == _currentUser.ID)
                        .Sum(p => (decimal?)(p.Num * p.Price)) ?? 0;

                    var statistics = db.Payments
                        .Include(p => p.Category)
                        .Where(p => p.UserID == _currentUser.ID)
                        .GroupBy(p => p.Category.Name)
                        .Select(g => new
                        {
                            CategoryName = g.Key,
                            PaymentCount = g.Count(),
                            TotalAmount = g.Sum(p => p.Num * p.Price),
                            AverageAmount = g.Average(p => p.Num * p.Price),
                            Percentage = totalSum > 0 ? g.Sum(p => p.Num * p.Price) / totalSum : 0
                        })
                        .OrderByDescending(s => s.TotalAmount)
                        .ToList();

                    DgStatistics.ItemsSource = statistics;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки статистики: {ex.Message}");
            }
        }

        private void LoadChart()
        {
            try
            {
                // Проверяем инициализацию всех необходимых компонентов
                if (CbGroupBy == null || CbChartType == null || UserChart == null)
                {
                    System.Diagnostics.Debug.WriteLine("Компоненты диаграммы еще не инициализированы");
                    return;
                }

                // Проверяем, что ChartAreas и Titles созданы
                if (UserChart.ChartAreas.Count == 0 || UserChart.Titles.Count == 0)
                {
                    InitializeChart();
                }

                using (var db = new Entities())
                {
                    // Проверяем контекст базы данных
                    if (db == null)
                    {
                        MessageBox.Show("Ошибка подключения к базе данных");
                        return;
                    }

                    // Определяем тип группировки
                    string groupBy = "По категориям";
                    if (CbGroupBy.SelectedItem is ComboBoxItem groupByItem && groupByItem.Content != null)
                    {
                        groupBy = groupByItem.Content.ToString();
                    }

                    // Создаем серию
                    Series series = new Series("Расходы");
                    series.IsValueShownAsLabel = true;
                    series.LabelFormat = "{0:N0} руб.";
                    series.Font = new System.Drawing.Font("Arial", 8);
                    series.Color = System.Drawing.Color.SteelBlue;

                    // Загружаем данные в зависимости от типа группировки
                    switch (groupBy)
                    {
                        case "По месяцам":
                            LoadMonthlyChart(series, db);
                            break;
                        default: // По категориям
                            LoadCategoryChart(series, db);
                            break;
                    }

                    // Очищаем и добавляем новую серию
                    UserChart.Series.Clear();

                    // Проверяем, есть ли данные для отображения
                    if (series.Points.Count > 0)
                    {
                        UserChart.Series.Add(series);

                        // Обновляем заголовок
                        if (UserChart.Titles.Count > 0)
                        {
                            UserChart.Titles[0].Text = $"Мои расходы ({groupBy.ToLower()})";
                        }
                    }
                    else
                    {
                        // Если нет данных, показываем сообщение
                        if (UserChart.Titles.Count > 0)
                        {
                            UserChart.Titles[0].Text = "Нет данных для отображения";
                        }
                    }
                }
            }
            catch (System.NullReferenceException nre)
            {
                // Не показываем сообщение пользователю, просто логируем
                System.Diagnostics.Debug.WriteLine($"NullReference в LoadChart: {nre.Message}");
            }
            catch (Exception ex)
            {
                // Не показываем сообщение пользователю при первой загрузке
                System.Diagnostics.Debug.WriteLine($"Ошибка в LoadChart: {ex.Message}");
            }
        }

        private void LoadCategoryChart(Series series, Entities db)
        {
            try
            {
                var chartType = GetChartType();
                series.ChartType = chartType;

                // Получаем данные по категориям с фильтрацией по датам
                var paymentsQuery = db.Payments
                    .Include(p => p.Category)
                    .Where(p => p.UserID == _currentUser.ID);

                // Применяем фильтры по дате
                if (DpStartDate.SelectedDate != null)
                {
                    paymentsQuery = paymentsQuery.Where(p => p.Date >= DpStartDate.SelectedDate);
                }

                if (DpEndDate.SelectedDate != null)
                {
                    paymentsQuery = paymentsQuery.Where(p => p.Date <= DpEndDate.SelectedDate);
                }

                var categoryData = paymentsQuery
                    .GroupBy(p => p.Category.Name)
                    .Select(g => new
                    {
                        Category = g.Key,
                        Total = g.Sum(p => p.Num * p.Price)
                    })
                    .Where(x => x.Total > 0) // Только категории с расходами
                    .OrderByDescending(x => x.Total)
                    .Take(10) // Берем топ-10 категорий
                    .ToList();

                // Цвета для диаграммы
                System.Drawing.Color[] colors = {
                    System.Drawing.Color.SteelBlue,
                    System.Drawing.Color.LightSeaGreen,
                    System.Drawing.Color.Tomato,
                    System.Drawing.Color.Gold,
                    System.Drawing.Color.MediumOrchid,
                    System.Drawing.Color.LightSkyBlue,
                    System.Drawing.Color.LightCoral,
                    System.Drawing.Color.MediumSeaGreen,
                    System.Drawing.Color.SandyBrown,
                    System.Drawing.Color.Plum
                };

                int colorIndex = 0;
                foreach (var item in categoryData)
                {
                    var dataPoint = new DataPoint();
                    dataPoint.SetValueXY(item.Category, (double)item.Total);
                    dataPoint.AxisLabel = item.Category;
                    dataPoint.LegendText = item.Category;
                    dataPoint.Label = item.Total.ToString("N0") + " руб.";
                    dataPoint.Color = colors[colorIndex % colors.Length];
                    dataPoint.LabelForeColor = System.Drawing.Color.Black;

                    series.Points.Add(dataPoint);
                    colorIndex++;
                }

                // Для круговой диаграммы настраиваем отображение
                if (chartType == SeriesChartType.Pie || chartType == SeriesChartType.Doughnut)
                {
                    series["PieLabelStyle"] = "Outside";
                    series["PieLineColor"] = "Black";
                    series.Label = "#VALX: #VALY руб.";
                }

                // Для столбчатой и линейной диаграмм настраиваем подписи
                if (chartType == SeriesChartType.Column || chartType == SeriesChartType.Line)
                {
                    series.Label = "#VALY руб.";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка загрузки диаграммы по категориям: {ex.Message}");
            }
        }

        private void LoadMonthlyChart(Series series, Entities db)
        {
            try
            {
                series.ChartType = GetChartType();

                // Базовый запрос с фильтрацией по пользователю
                var paymentsQuery = db.Payments
                    .Where(p => p.UserID == _currentUser.ID);

                // Применяем фильтры по дате
                if (DpStartDate.SelectedDate != null)
                {
                    paymentsQuery = paymentsQuery.Where(p => p.Date >= DpStartDate.SelectedDate);
                }

                if (DpEndDate.SelectedDate != null)
                {
                    paymentsQuery = paymentsQuery.Where(p => p.Date <= DpEndDate.SelectedDate);
                }

                // Получаем данные по месяцам
                var monthlyData = paymentsQuery
                    .AsEnumerable() // Переключаемся на клиентскую сторону для работы с DateTime
                    .GroupBy(p => new { p.Date.Year, p.Date.Month })
                    .Select(g => new
                    {
                        Period = new DateTime(g.Key.Year, g.Key.Month, 1),
                        Total = g.Sum(p => p.Num * p.Price)
                    })
                    .Where(x => x.Total > 0) // Только месяцы с расходами
                    .OrderBy(x => x.Period)
                    .ToList();

                foreach (var item in monthlyData)
                {
                    var dataPoint = new DataPoint();
                    string periodLabel = item.Period.ToString("MMM yyyy", CultureInfo.GetCultureInfo("ru-RU"));

                    dataPoint.SetValueXY(periodLabel, (double)item.Total);
                    dataPoint.AxisLabel = periodLabel;
                    dataPoint.LegendText = periodLabel;
                    dataPoint.Label = item.Total.ToString("N0") + " руб.";
                    dataPoint.Color = System.Drawing.Color.SteelBlue;
                    dataPoint.LabelForeColor = System.Drawing.Color.Black;

                    series.Points.Add(dataPoint);
                }

                // Настройка для линейной диаграммы
                if (series.ChartType == SeriesChartType.Line)
                {
                    series.BorderWidth = 3;
                    series.MarkerStyle = MarkerStyle.Circle;
                    series.MarkerSize = 8;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка загрузки месячной диаграммы: {ex.Message}");
            }
        }

        private SeriesChartType GetChartType()
        {
            try
            {
                if (CbChartType.SelectedItem is ComboBoxItem selectedItem && selectedItem.Content != null)
                {
                    return selectedItem.Content.ToString() switch
                    {
                        "Круговая" => SeriesChartType.Pie,
                        "Линейная" => SeriesChartType.Line,
                        _ => SeriesChartType.Column
                    };
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка определения типа диаграммы: {ex.Message}");
            }

            // Значение по умолчанию
            return SeriesChartType.Column;
        }

        // Обработчики событий для вкладки диаграмм
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source is TabControl && TabCharts.IsSelected && _isInitialized)
            {
                // При переключении на вкладку диаграмм загружаем данные
                LoadChart();
            }
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized)
            {
                LoadPayments();
                LoadStatistics();
                if (TabCharts.IsSelected)
                {
                    LoadChart();
                }
            }
        }

        private void CbCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized)
            {
                LoadPayments();
                LoadStatistics();
                if (TabCharts.IsSelected)
                {
                    LoadChart();
                }
            }
        }

        private void CbChartType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized && TabCharts.IsSelected)
            {
                LoadChart();
            }
        }

        private void CbGroupBy_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized && TabCharts.IsSelected)
            {
                LoadChart();
            }
        }

        private void BtnApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            if (_isInitialized)
            {
                LoadPayments();
                LoadStatistics();
                if (TabCharts.IsSelected)
                {
                    LoadChart();
                }
            }
        }

        private void BtnClearFilter_Click(object sender, RoutedEventArgs e)
        {
            if (_isInitialized)
            {
                var startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                var endDate = DateTime.Now;

                DpStartDate.SelectedDate = startDate;
                DpEndDate.SelectedDate = endDate;
                CbCategory.SelectedIndex = -1;

                LoadPayments();
                LoadStatistics();
                if (TabCharts.IsSelected)
                {
                    LoadChart();
                }
            }
        }

        // Остальные методы остаются без изменений...
        private void BtnAddPayment_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddEditPaymentPage(null, _currentUser));
        }

        private void BtnDeletePayment_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button?.DataContext != null)
            {
                dynamic payment = button.DataContext;
                int paymentId = payment.ID;

                var result = MessageBox.Show("Вы уверены, что хотите удалить этот платеж?",
                    "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        using (var db = new Entities())
                        {
                            var paymentToDelete = db.Payments.Find(paymentId);
                            if (paymentToDelete != null)
                            {
                                db.Payments.Remove(paymentToDelete);
                                db.SaveChanges();
                                LoadPayments();
                                LoadStatistics();
                                if (TabCharts.IsSelected)
                                {
                                    LoadChart();
                                }
                                MessageBox.Show("Платеж успешно удален!");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка удаления платежа: {ex.Message}");
                    }
                }
            }
        }

        private void BtnChangePass_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                NavigationService.Navigate(new ChangePasswordPage(_currentUser));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка перехода на страницу смены пароля: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            // Реализация экспорта...
        }

        private void BtnLogout_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите выйти?", "Подтверждение выхода",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    NavigationService.Navigate(new AuthPage());
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка выхода: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DgStatistics_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Пустая реализация
        }
    }
}