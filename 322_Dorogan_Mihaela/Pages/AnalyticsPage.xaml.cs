using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class AnalyticsPage : Page
    {
        private User _currentAdmin;

        public AnalyticsPage(User admin)
        {
            InitializeComponent();
            _currentAdmin = admin;
            InitializeCharts();
            LoadData();
        }

        private void InitializeCharts()
        {
            // Установка дат по умолчанию
            var endDate = DateTime.Now;
            var startDate = new DateTime(endDate.Year, endDate.Month, 1);

            DpCatStartDate.SelectedDate = startDate;
            DpCatEndDate.SelectedDate = endDate;
            DpUserStartDate.SelectedDate = startDate;
            DpUserEndDate.SelectedDate = endDate;
            DpTrendStartDate.SelectedDate = startDate.AddMonths(-6);
            DpTrendEndDate.SelectedDate = endDate;
        }

        private void LoadData()
        {
            LoadCategoryChart();
            LoadUserChart();
            LoadTrendChart();
        }

        private void LoadCategoryChart()
        {
            try
            {
                var startDate = DpCatStartDate.SelectedDate;
                var endDate = DpCatEndDate.SelectedDate;

                using (var db = new Entities())
                {
                    var categoryData = db.Payments
                        .Include(p => p.Category)
                        .Where(p => p.Date >= startDate && p.Date <= endDate)
                        .GroupBy(p => p.Category.Name)
                        .Select(g => new
                        {
                            Category = g.Key,
                            Total = g.Sum(p => p.Num * p.Price),
                            Count = g.Count()
                        })
                        .OrderByDescending(x => x.Total)
                        .ToList();

                    // Временная заглушка для диаграммы
                    MessageBox.Show($"Загружено {categoryData.Count} категорий для анализа", "Информация");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных по категориям: {ex.Message}");
            }
        }

        private void LoadUserChart()
        {
            try
            {
                var startDate = DpUserStartDate.SelectedDate;
                var endDate = DpUserEndDate.SelectedDate;

                using (var db = new Entities())
                {
                    var userData = db.Payments
                        .Include(p => p.User)
                        .Where(p => p.Date >= startDate && p.Date <= endDate)
                        .GroupBy(p => p.User.FIO)
                        .Select(g => new
                        {
                            User = g.Key,
                            Total = g.Sum(p => p.Num * p.Price),
                            Count = g.Count()
                        })
                        .OrderByDescending(x => x.Total)
                        .Take(10)
                        .ToList();

                    // Временная заглушка для диаграммы
                    MessageBox.Show($"Загружено {userData.Count} пользователей для анализа", "Информация");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных по пользователям: {ex.Message}");
            }
        }

        private void LoadTrendChart()
        {
            try
            {
                var startDate = DpTrendStartDate.SelectedDate;
                var endDate = DpTrendEndDate.SelectedDate;

                using (var db = new Entities())
                {
                    var trendData = db.Payments
                        .Where(p => p.Date >= startDate && p.Date <= endDate)
                        .GroupBy(p => DbFunctions.TruncateTime(p.Date))
                        .Select(g => new
                        {
                            Date = g.Key,
                            Total = g.Sum(p => p.Num * p.Price),
                            Count = g.Count()
                        })
                        .OrderBy(x => x.Date)
                        .ToList();

                    // Временная заглушка для диаграммы
                    MessageBox.Show($"Загружено {trendData.Count} дней для анализа трендов", "Информация");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных динамики: {ex.Message}");
            }
        }

        private void DateRange_Changed(object sender, SelectionChangedEventArgs e)
        {
            LoadData();
        }

        private void ChartType_Changed(object sender, SelectionChangedEventArgs e)
        {
            LoadData();
        }

        private void GroupBy_Changed(object sender, SelectionChangedEventArgs e)
        {
            LoadTrendChart();
        }

        private void BtnExportCatChart_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Экспорт диаграммы будет реализован позже", "Информация");
        }

        private void BtnExportUserChart_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Экспорт диаграммы будет реализован позже", "Информация");
        }

        private void BtnExportTrend_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Экспорт диаграммы будет реализован позже", "Информация");
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}