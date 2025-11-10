using System;
using System.Linq;
using System.Windows;

namespace _322_Dorogan_Mihaela
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Инициализация базы данных при запуске
            //InitializeDatabase();

            // Глобальная обработка необработанных исключений
            AppDomain.CurrentDomain.UnhandledException += (s, args) =>
            {
                MessageBox.Show($"Критическая ошибка: {((Exception)args.ExceptionObject).Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            };

            DispatcherUnhandledException += (s, args) =>
            {
                MessageBox.Show($"Ошибка приложения: {args.Exception.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                args.Handled = true;
            };
        }

        private void InitializeDatabase()
        {
            try
            {
                MessageBox.Show("Начало инициализации базы данных...");

                using (var db = new DEntities())
                {
                    MessageBox.Show("Подключение к базе данных создано");

                    // Пробуем просто подключиться к базе
                    if (db.Database.Exists())
                    {
                        MessageBox.Show("База данных существует");
                        AddInitialData(db);
                    }
                    else
                    {
                        MessageBox.Show("База данных не существует, создаем...");
                        db.Database.Create();
                        MessageBox.Show("База данных создана");
                        AddInitialData(db);
                    }
                }

                MessageBox.Show("Инициализация завершена успешно");
            }
            catch (System.Data.SqlClient.SqlException sqlEx)
            {
                MessageBox.Show($"Ошибка SQL: {sqlEx.Message}\n\nПроверьте:\n1. Установлен ли SQL Server Express\n2. Запущена ли служба SQL Server",
                    "Ошибка базы данных", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Общая ошибка: {ex.Message}\n\n{ex.StackTrace}",
                    "Ошибка приложения", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddInitialData(DEntities db)
        {
            // Добавляем администратора по умолчанию
            if (!db.Users.Any(u => u.Login == "admin"))
            {
                var adminUser = new User
                {
                    Login = "admin",
                    Password = GetHash("admin123"), // Нужно добавить метод GetHash
                    FIO = "Администратор Системы",
                    Role = "Admin"
                };
                db.Users.Add(adminUser);
            }

            // Добавляем базовые категории
            if (!db.Categories.Any())
            {
                var categories = new[]
                {
                new Category { Name = "Продукты питания" },
                new Category { Name = "Коммунальные услуги" },
                new Category { Name = "Транспорт" },
                new Category { Name = "Развлечения" },
                new Category { Name = "Одежда" }
            };

                foreach (var category in categories)
                {
                    db.Categories.Add(category);
                }
            }

            db.SaveChanges();
        }

        private string GetHash(string input)
        {
            using (var sha1 = System.Security.Cryptography.SHA1.Create())
            {
                var hash = sha1.ComputeHash(System.Text.Encoding.UTF8.GetBytes(input));
                return string.Concat(hash.Select(b => b.ToString("X2")));
            }
        }
    }
}