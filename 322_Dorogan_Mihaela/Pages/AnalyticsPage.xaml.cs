using System;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Globalization;

namespace _322_Dorogan_Mihaela.Pages
{
    public partial class AnalyticsPage : Page
    {
        private User _currentAdmin;
        private bool _isInitialized = false;

        public AnalyticsPage(User admin)
        {
            InitializeComponent();
            _currentAdmin = admin;
            Loaded += AnalyticsPage_Loaded;
        }

        private void AnalyticsPage_Loaded(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized)
            {
                InitializeCharts();
                LoadData();
                _isInitialized = true;
            }
        }

        private void InitializeCharts()
        {
            try
            {
                // Инициализация диаграммы категорий
                InitializeChart(ChartCategories, "Распределение платежей по категориям");

                // Инициализация диаграммы пользователей
                InitializeChart(ChartUsers, "Топ пользователей по расходам");

                // Инициализация диаграммы трендов
                InitializeChart(ChartTrend, "Динамика платежей по времени");
            }
            catch
            {
                // Не показываем ошибку пользователю при инициализации
                System.Diagnostics.Debug.WriteLine("Ошибка инициализации диаграмм");
            }
        }

        private void InitializeChart(Chart chart, string title)
        {
            try
            {
                if (chart == null) return;

                chart.ChartAreas.Clear();
                chart.Legends.Clear();
                chart.Titles.Clear();
                chart.Series.Clear();

                chart.ChartAreas.Add(new ChartArea());
                chart.Legends.Add(new Legend());
                chart.Titles.Add(new Title(title));

                if (chart.Titles.Count > 0)
                {
                    chart.Titles[0].Font = new System.Drawing.Font("Arial", 14, System.Drawing.FontStyle.Bold);
                }

                // Настройка области диаграммы
                if (chart.ChartAreas.Count > 0)
                {
                    chart.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
                    chart.ChartAreas[0].AxisX.Interval = 1;
                    chart.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 8);
                }
            }
            catch
            {
                System.Diagnostics.Debug.WriteLine("Ошибка инициализации диаграммы");
            }
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
                using (var db = new DEntities())
                {
                    // Загружаем ВСЕ данные без фильтрации по дате
                    var categoryData = db.Payments
                        .Include(p => p.Category)
                        .GroupBy(p => p.Category.Name)
                        .Select(g => new CategoryData
                        {
                            Name = g.Key,
                            Total = g.Sum(p => p.Num * p.Price),
                            Count = g.Count()
                        })
                        .Where(x => x.Total > 0) // Только категории с платежами
                        .OrderByDescending(x => x.Total)
                        .ToList();

                    if (ChartCategories == null) return;

                    ChartCategories.Series.Clear();

                    // Если нет данных, показываем сообщение
                    if (!categoryData.Any())
                    {
                        ChartCategories.Titles[0].Text = "Нет данных по категориям";
                        return;
                    }

                    var series = new Series("Категории");
                    series.ChartType = GetChartType(CmbCatChartType);
                    series.IsValueShownAsLabel = true;
                    series.LabelFormat = "{0:N0} руб.";
                    series.Font = new System.Drawing.Font("Arial", 8);

                    // Цвета для диаграммы
                    System.Drawing.Color[] colors = {
                        System.Drawing.Color.SteelBlue, System.Drawing.Color.LightSeaGreen,
                        System.Drawing.Color.Tomato, System.Drawing.Color.Gold,
                        System.Drawing.Color.MediumOrchid, System.Drawing.Color.LightSkyBlue
                    };

                    int colorIndex = 0;
                    foreach (var item in categoryData)
                    {
                        var dataPoint = new DataPoint();
                        dataPoint.SetValueXY(item.Name, (double)item.Total);
                        dataPoint.AxisLabel = item.Name;
                        dataPoint.LegendText = item.Name;
                        dataPoint.Label = item.Total?.ToString("N0") + " руб." ?? "0 руб.";
                        dataPoint.Color = colors[colorIndex % colors.Length];
                        dataPoint.LabelForeColor = System.Drawing.Color.Black;

                        series.Points.Add(dataPoint);
                        colorIndex++;
                    }

                    ChartCategories.Series.Add(series);

                    // Форматирование для круговой диаграммы
                    if (series.ChartType == SeriesChartType.Pie)
                    {
                        series["PieLabelStyle"] = "Outside";
                        series["PieLineColor"] = "Black";
                        series.Label = "#VALX: #VALY руб.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных по категориям: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadUserChart()
        {
            try
            {
                var topCount = GetTopCount(); // Получаем количество пользователей для отображения

                using (var db = new DEntities())
                {
                    // Загружаем ВСЕ данные без фильтрации по дате
                    var userQuery = db.Payments
                        .Include(p => p.User)
                        .GroupBy(p => new { p.User.ID, p.User.FIO })
                        .Select(g => new UserData
                        {
                            Name = g.Key.FIO,
                            Total = g.Sum(p => p.Num * p.Price),
                            Count = g.Count()
                        })
                        .Where(x => x.Total > 0) // Только пользователи с платежами
                        .OrderByDescending(x => x.Total);

                    // Применяем ограничение по количеству, если не выбрано "Все"
                    var userData = (topCount > 0) ? userQuery.Take(topCount).ToList() : userQuery.ToList();

                    if (ChartUsers == null) return;

                    ChartUsers.Series.Clear();

                    // Если нет данных, показываем сообщение
                    if (!userData.Any())
                    {
                        ChartUsers.Titles[0].Text = "Нет данных по пользователям";
                        return;
                    }

                    var series = new Series("Пользователи");
                    series.ChartType = SeriesChartType.Column;
                    series.IsValueShownAsLabel = true;
                    series.LabelFormat = "{0:N0} руб.";
                    series.Font = new System.Drawing.Font("Arial", 8);
                    series.Color = System.Drawing.Color.SteelBlue;

                    foreach (var item in userData)
                    {
                        var dataPoint = new DataPoint();
                        dataPoint.SetValueXY(item.Name, (double)item.Total);
                        dataPoint.AxisLabel = item.Name;
                        dataPoint.LegendText = item.Name;
                        dataPoint.Label = item.Total?.ToString("N0") + " руб." ?? "0 руб.";
                        dataPoint.Color = System.Drawing.Color.SteelBlue;

                        series.Points.Add(dataPoint);
                    }

                    ChartUsers.Series.Add(series);

                    // Настройка внешнего вида
                    if (ChartUsers.ChartAreas.Count > 0)
                    {
                        ChartUsers.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
                        ChartUsers.ChartAreas[0].AxisX.Interval = 1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных по пользователям: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadTrendChart()
        {
            try
            {
                // Проверяем, что компоненты инициализированы
                if (ChartTrend == null || CmbTrendGroupBy == null)
                {
                    return;
                }

                using (var db = new DEntities())
                {
                    string groupBy = "По дням"; // значение по умолчанию

                    // Безопасно получаем тип группировки
                    if (CmbTrendGroupBy.SelectedItem is ComboBoxItem groupByItem && groupByItem.Content != null)
                    {
                        groupBy = groupByItem.Content.ToString();
                    }

                    List<TrendData> data = new List<TrendData>();

                    // Загружаем ВСЕ данные без фильтрации по дате
                    switch (groupBy)
                    {
                        case "По неделям":
                            data = db.Payments
                                .AsEnumerable()
                                .GroupBy(p => new {
                                    Year = p.Date.Year,
                                    Week = (p.Date.DayOfYear - 1) / 7
                                })
                                .Select(g => new TrendData
                                {
                                    Period = $"{g.Key.Year}-W{g.Key.Week + 1}",
                                    Total = g.Sum(p => p.Num * p.Price),
                                    Count = g.Count()
                                })
                                .Where(x => x.Total > 0)
                                .OrderBy(x => x.Period)
                                .ToList();
                            break;

                        case "По месяцам":
                            data = db.Payments
                                .AsEnumerable()
                                .GroupBy(p => new { p.Date.Year, p.Date.Month })
                                .Select(g => new TrendData
                                {
                                    Period = new DateTime(g.Key.Year, g.Key.Month, 1).ToString("MMM yyyy", new CultureInfo("ru-RU")),
                                    Total = g.Sum(p => p.Num * p.Price),
                                    Count = g.Count()
                                })
                                .Where(x => x.Total > 0)
                                .OrderBy(x => x.Period)
                                .ToList();
                            break;

                        default: // По дням
                            data = db.Payments
                                .AsEnumerable()
                                .GroupBy(p => p.Date.Date)
                                .Select(g => new TrendData
                                {
                                    Period = g.Key.ToString("dd.MM.yyyy"),
                                    Total = g.Sum(p => p.Num * p.Price),
                                    Count = g.Count()
                                })
                                .Where(x => x.Total > 0)
                                .OrderBy(x => x.Period)
                                .ToList();
                            break;
                    }

                    ChartTrend.Series.Clear();

                    // Если нет данных, показываем сообщение
                    if (!data.Any())
                    {
                        ChartTrend.Titles[0].Text = "Нет данных по динамике";
                        return;
                    }

                    var series = new Series("Динамика");
                    series.ChartType = SeriesChartType.Line;
                    series.IsValueShownAsLabel = true;
                    series.LabelFormat = "{0:N0} руб.";
                    series.Font = new System.Drawing.Font("Arial", 8);
                    series.Color = System.Drawing.Color.Green;
                    series.BorderWidth = 3;
                    series.MarkerStyle = MarkerStyle.Circle;
                    series.MarkerSize = 8;

                    foreach (var item in data)
                    {
                        var dataPoint = new DataPoint();
                        dataPoint.SetValueXY(item.Period, (double)item.Total);
                        dataPoint.AxisLabel = item.Period;
                        dataPoint.LegendText = item.Period;
                        dataPoint.Label = item.Total?.ToString("N0") + " руб." ?? "0 руб.";
                        dataPoint.Color = System.Drawing.Color.Green;

                        series.Points.Add(dataPoint);
                    }

                    ChartTrend.Series.Add(series);

                    // Настройка внешнего вида с проверками
                    if (ChartTrend.ChartAreas.Count > 0)
                    {
                        ChartTrend.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
                        ChartTrend.ChartAreas[0].AxisX.Interval = 1;
                    }
                }
            }
            catch (System.NullReferenceException)
            {
                System.Diagnostics.Debug.WriteLine("NullReference в LoadTrendChart");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных динамики: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int GetTopCount()
        {
            try
            {
                if (CmbUserTopCount.SelectedItem is ComboBoxItem selectedItem && selectedItem.Content != null)
                {
                    string content = selectedItem.Content.ToString();
                    if (content == "Все") return 0; // 0 означает "все"
                    if (int.TryParse(content, out int count)) return count;
                }
            }
            catch
            {
                System.Diagnostics.Debug.WriteLine("Ошибка получения топ-количества");
            }

            return 5; // значение по умолчанию
        }

        private SeriesChartType GetChartType(ComboBox comboBox)
        {
            try
            {
                if (comboBox?.SelectedItem is ComboBoxItem selectedItem && selectedItem.Content != null)
                {
                    return selectedItem.Content.ToString() switch
                    {
                        "Круговая" => SeriesChartType.Pie,
                        "Линейная" => SeriesChartType.Line,
                        _ => SeriesChartType.Column
                    };
                }
            }
            catch
            {
                System.Diagnostics.Debug.WriteLine("Ошибка определения типа диаграммы");
            }

            return SeriesChartType.Column;
        }

        private void ChartType_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized)
            {
                LoadCategoryChart();
            }
        }

        private void TopCount_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized)
            {
                LoadUserChart();
            }
        }

        private void GroupBy_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized)
            {
                LoadTrendChart();
            }
        }

        private void BtnExportCatChart_Click(object sender, RoutedEventArgs e)
        {
            ExportChartToExcel(ChartCategories, "Категории");
        }

        private void BtnExportUserChart_Click(object sender, RoutedEventArgs e)
        {
            ExportChartToExcel(ChartUsers, "Пользователи");
        }

        private void BtnExportTrend_Click(object sender, RoutedEventArgs e)
        {
            ExportChartToExcel(ChartTrend, "Динамика");
        }

        private void ExportChartToExcel(Chart chart, string reportType)
        {
            try
            {
                if (chart == null || chart.Series.Count == 0 || chart.Series[0].Points.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта", "Внимание",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FileName = $"Диаграмма_{reportType}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    Excel.Application excelApp = null;
                    Excel.Workbook workbook = null;
                    Excel.Worksheet worksheet = null;

                    try
                    {
                        excelApp = new Excel.Application();
                        workbook = excelApp.Workbooks.Add();
                        worksheet = workbook.ActiveSheet;

                        // Заголовок отчета
                        worksheet.Cells[1, 1] = $"Отчет: {chart.Titles[0]?.Text ?? reportType}";
                        worksheet.Cells[1, 1].Font.Bold = true;
                        worksheet.Cells[1, 1].Font.Size = 14;

                        worksheet.Cells[2, 1] = $"Сгенерирован: {DateTime.Now:dd.MM.yyyy HH:mm}";

                        // Заголовки таблицы
                        worksheet.Cells[4, 1] = "Название";
                        worksheet.Cells[4, 2] = "Сумма (руб.)";
                        worksheet.Cells[4, 1].Font.Bold = true;
                        worksheet.Cells[4, 2].Font.Bold = true;

                        // Данные из диаграммы
                        int row = 5;
                        var series = chart.Series[0];
                        foreach (var point in series.Points)
                        {
                            worksheet.Cells[row, 1] = point.AxisLabel;
                            worksheet.Cells[row, 2] = point.YValues[0];
                            row++;
                        }

                        // Итоговая строка
                        worksheet.Cells[row, 1] = "ВСЕГО:";
                        worksheet.Cells[row, 1].Font.Bold = true;
                        worksheet.Cells[row, 2] = series.Points.Sum(p => p.YValues[0]);
                        worksheet.Cells[row, 2].Font.Bold = true;

                        // Форматирование
                        Excel.Range moneyCol = worksheet.Range[worksheet.Cells[5, 2], worksheet.Cells[row, 2]];
                        moneyCol.NumberFormat = "#,##0.00 \"руб.\"";

                        // Автоподбор ширины колонок
                        worksheet.Columns[1].AutoFit();
                        worksheet.Columns[2].AutoFit();

                        // Добавление границ
                        Excel.Range dataRange = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[row, 2]];
                        dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        // Сохранение
                        workbook.SaveAs(saveDialog.FileName);

                        MessageBox.Show($"Диаграмма '{reportType}' успешно экспортирована в Excel!\nФайл: {Path.GetFileName(saveDialog.FileName)}",
                            "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                        // Открытие файла
                        System.Diagnostics.Process.Start(saveDialog.FileName);
                    }
                    finally
                    {
                        // Освобождение ресурсов
                        workbook?.Close(false);
                        excelApp?.Quit();

                        // Явное освобождение COM объектов
                        if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                        if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                        if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта диаграммы: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }

    // Вспомогательные классы для типизации данных
    public class CategoryData
    {
        public string Name { get; set; }
        public decimal? Total { get; set; }
        public int Count { get; set; }
    }

    public class UserData
    {
        public string Name { get; set; }
        public decimal? Total { get; set; }
        public int Count { get; set; }
    }

    public class TrendData
    {
        public string Period { get; set; }
        public decimal? Total { get; set; }
        public int Count { get; set; }
    }
}