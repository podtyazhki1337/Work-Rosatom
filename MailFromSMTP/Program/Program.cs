using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Linq;

namespace ChangeManagementSystem
{
    public class EmailRequest
    {
        public int ID { get; set; }
        public int CR_ID { get; set; }
        public string Group { get; set; }
        public string Email { get; set; }
    }

    public interface IEmailRepository
    {
        void Insert(EmailRequest request);
        void Update(int id, EmailRequest request);
        void Delete(int id);
        List<EmailRequest> GetByCRId(int crId);
    }

    public class EmailService
    {
        private readonly IEmailRepository _repository;
        private readonly Dictionary<string, string> _groupTemplates;
        private readonly Dictionary<string, string> _groupLinks;
        private readonly string _logFilePath;

        public EmailService(IEmailRepository repository, string logFilePath)
        {
            _repository = repository;
            _logFilePath = logFilePath;

            _groupTemplates = new Dictionary<string, string>
            {
                { "Филиал АСЭ в Венгрии", "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\nПрошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\nСсылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, …" },
                { "АЭП", "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\nПрошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\nСсылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, …" },
                { "Субподрядчик", "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\nПрошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\nСсылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, …" },
                { "Венгерский Заказчик", "Dear Sir,\n\nI am sending you Change Request No. {0} for your information.\nLink to materials: {1}\n\nBest regards, …" }
            };

            _groupLinks = new Dictionary<string, string>
            {
                { "Филиал АСЭ в Венгрии", "http://ase-hungary/change/{0}" },
                { "АЭП", "http://voshod/change/{0}" },
                { "Субподрядчик", "http://subcontractor/change/{0}" },
                { "Венгерский Заказчик", "http://ftp/change/{0}" }
            };
        }

        public void SendEmailsConsole(int crId, string deadline = "")
        {
            var emailRequests = _repository.GetByCRId(crId);
            List<string> failedEmails = new List<string>();
            StringBuilder resultMessage = new StringBuilder("Рассылка произведена");
            StringBuilder logBuilder = new StringBuilder();

            foreach (var request in emailRequests)
            {
                string template = _groupTemplates[request.Group];
                string link = string.Format(_groupLinks[request.Group], crId);
                string subject = request.Group == "Венгерский Заказчик" ? $"Change Request No. {crId} notification" : $"О рассмотрении и согласовании Запросов на изменения № {crId}";
                string body = request.Group == "Венгерский Заказчик" ? string.Format(template, crId, link) : string.Format(template, crId, link, deadline);

                logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало отправки письма:");
                logBuilder.AppendLine($"Кому: {request.Email}");
                logBuilder.AppendLine($"Тема: {subject}");
                logBuilder.AppendLine($"Тело письма:\n{body}");

                try
                {
                    SendEmail(request.Email, subject, body);
                    logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Письмо успешно отправлено на {request.Email} (уведомление о доставке запрошено)");
                }
                catch (Exception ex)
                {
                    failedEmails.Add(request.Email);
                    logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Ошибка отправки письма на {request.Email}: {ex.Message}");
                    Console.WriteLine($"Ошибка отправки письма на {request.Email}: {ex.Message}");
                }

                logBuilder.AppendLine("----------------------------------------");
            }

            if (failedEmails.Count > 0)
            {
                resultMessage.Append("\nПисьма для " + string.Join(", ", failedEmails.ToArray()) + " не доставлены, т.к. адрес не существует");
            }

            Console.WriteLine(resultMessage.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }

        private void SendEmail(string to, string subject, string body)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("no-reply@changemanagement.com");
            mail.To.Add(to);
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = false;
            mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess; // Запрос уведомления о доставке

            SmtpClient smtp = new SmtpClient("smtp.yourserver.com"); // Замените на реальный SMTP-сервер
            smtp.Send(mail);

            mail.Dispose();
        }

        public bool ValidateRequest(EmailRequest request, out string error)
        {
            error = string.Empty;
            if (string.IsNullOrEmpty(request.Group))
            {
                error = "Поле Group обязательно";
                return false;
            }
            if (string.IsNullOrEmpty(request.Email))
            {
                error = "Поле Email обязательно";
                return false;
            }
            if (request.Email.Length > 100 || request.Group.Length > 100)
            {
                error = "Превышена максимальная длина поля (100 символов)";
                return false;
            }
            if (!_groupTemplates.ContainsKey(request.Group))
            {
                error = "Указана недопустимая группа";
                return false;
            }
            return true;
        }
    }

    public class SQLiteEmailRepository : IEmailRepository
    {
        private readonly string _connectionString;
        private readonly string _logFilePath;

        public SQLiteEmailRepository(string dbPath, string logFilePath)
        {
            _connectionString = $"Data Source={dbPath};Version=3;";
            _logFilePath = logFilePath;
        }

        public void Insert(EmailRequest request)
        {
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало операции Insert");

            using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                string query = "INSERT INTO ChM_Send_emails (CR_ID, [Group], Email) VALUES (@CR_ID, @Group, @Email)";
                logBuilder.AppendLine($"Выполняется запрос: {query}");
                logBuilder.AppendLine($"Параметры: CR_ID={request.CR_ID}, Group={request.Group}, Email={request.Email}");

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@CR_ID", request.CR_ID);
                    cmd.Parameters.AddWithValue("@Group", request.Group);
                    cmd.Parameters.AddWithValue("@Email", request.Email);
                    int rowsAffected = cmd.ExecuteNonQuery();
                    logBuilder.AppendLine($"Результат: Вставлено строк - {rowsAffected}");
                }
            }

            logBuilder.AppendLine("----------------------------------------");
            Console.WriteLine(logBuilder.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }

        public void Update(int id, EmailRequest request)
        {
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало операции Update");

            using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                string query = "UPDATE ChM_Send_emails SET CR_ID = @CR_ID, [Group] = @Group, Email = @Email WHERE ID = @ID";
                logBuilder.AppendLine($"Выполняется запрос: {query}");
                logBuilder.AppendLine($"Параметры: ID={id}, CR_ID={request.CR_ID}, Group={request.Group}, Email={request.Email}");

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    cmd.Parameters.AddWithValue("@CR_ID", request.CR_ID);
                    cmd.Parameters.AddWithValue("@Group", request.Group);
                    cmd.Parameters.AddWithValue("@Email", request.Email);
                    int rowsAffected = cmd.ExecuteNonQuery();
                    logBuilder.AppendLine($"Результат: Обновлено строк - {rowsAffected}");
                }
            }

            logBuilder.AppendLine("----------------------------------------");
            Console.WriteLine(logBuilder.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }

        public void Delete(int id)
        {
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало операции Delete");

            using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                string query = "DELETE FROM ChM_Send_emails WHERE ID = @ID";
                logBuilder.AppendLine($"Выполняется запрос: {query}");
                logBuilder.AppendLine($"Параметры: ID={id}");

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);
                    int rowsAffected = cmd.ExecuteNonQuery();
                    logBuilder.AppendLine($"Результат: Удалено строк - {rowsAffected}");
                }
            }

            logBuilder.AppendLine("----------------------------------------");
            Console.WriteLine(logBuilder.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }

        public List<EmailRequest> GetByCRId(int crId)
        {
            List<EmailRequest> requests = new List<EmailRequest>();
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало операции GetByCRId");

            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
                {
                    conn.Open();
                    logBuilder.AppendLine($"Соединение с базой открыто: {conn.State}");

                    string query = "SELECT ID, CR_ID, [Group], Email FROM ChM_Send_emails WHERE CR_ID = @CR_ID";
                    logBuilder.AppendLine($"Выполняется запрос: {query}");
                    logBuilder.AppendLine($"Параметры: CR_ID={crId}");

                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@CR_ID", crId);
                        logBuilder.AppendLine("Команда подготовлена, выполняется чтение...");

                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            logBuilder.AppendLine("Результат:");
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    var request = new EmailRequest
                                    {
                                        ID = reader["ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["ID"]),
                                        CR_ID = reader["CR_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["CR_ID"]),
                                        Group = reader["Group"] == DBNull.Value ? string.Empty : reader["Group"].ToString(),
                                        Email = reader["Email"] == DBNull.Value ? string.Empty : reader["Email"].ToString()
                                    };
                                    requests.Add(request);
                                    logBuilder.AppendLine($"ID={request.ID}, CR_ID={request.CR_ID}, Group={request.Group}, Email={request.Email}");
                                }
                            }
                            else
                            {
                                logBuilder.AppendLine("Нет данных");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logBuilder.AppendLine($"Ошибка при обращении к базе: {ex.Message}");
                logBuilder.AppendLine($"Стек вызовов: {ex.StackTrace}");
            }

            logBuilder.AppendLine($"Найдено записей: {requests.Count}");
            logBuilder.AppendLine("----------------------------------------");
            Console.WriteLine(logBuilder.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());

            return requests;
        }
    }

    public class EmailServiceTests
    {
        private readonly IEmailRepository _repository;
        private readonly EmailService _service;

        public EmailServiceTests(IEmailRepository repository, string logFilePath)
        {
            _repository = repository;
            _service = new EmailService(repository, logFilePath);
        }

        public void RunTests()
        {
            Console.WriteLine("Запуск тестов...");

            TestValidationSuccess();
            TestValidationFailureEmptyGroup();
            TestValidationFailureEmptyEmail();
            TestValidationFailureLongFields();
            TestInsertAndRetrieve();
            TestUpdate();
            TestDelete();

            Console.WriteLine("Тесты завершены.");
        }

        private void TestValidationSuccess()
        {
            var request = new EmailRequest { CR_ID = 1, Group = "Филиал АСЭ в Венгрии", Email = "test@ase.com" };
            bool result = _service.ValidateRequest(request, out string error);
            Console.WriteLine(result && string.IsNullOrEmpty(error) ? "TestValidationSuccess: Успех" : $"TestValidationSuccess: Провал - {error}");
        }

        private void TestValidationFailureEmptyGroup()
        {
            var request = new EmailRequest { CR_ID = 1, Group = "", Email = "test@ase.com" };
            bool result = _service.ValidateRequest(request, out string error);
            Console.WriteLine(!result && error == "Поле Group обязательно" ? "TestValidationFailureEmptyGroup: Успех" : $"TestValidationFailureEmptyGroup: Провал - {error}");
        }

        private void TestValidationFailureEmptyEmail()
        {
            var request = new EmailRequest { CR_ID = 1, Group = "Филиал АСЭ в Венгрии", Email = "" };
            bool result = _service.ValidateRequest(request, out string error);
            Console.WriteLine(!result && error == "Поле Email обязательно" ? "TestValidationFailureEmptyEmail: Успех" : $"TestValidationFailureEmptyEmail: Провал - {error}");
        }

        private void TestValidationFailureLongFields()
        {
            var longString = new string('a', 101);
            var request = new EmailRequest { CR_ID = 1, Group = longString, Email = "test@ase.com" };
            bool result = _service.ValidateRequest(request, out string error);
            Console.WriteLine(!result && error == "Превышена максимальная длина поля (100 символов)" ? "TestValidationFailureLongFields: Успех" : $"TestValidationFailureLongFields: Провал - {error}");
        }

        private void TestInsertAndRetrieve()
        {
            var request = new EmailRequest { CR_ID = 2, Group = "АЭП", Email = "test@aep.com" };
            _repository.Insert(request);
            var retrieved = _repository.GetByCRId(2);
            bool success = retrieved.Count > 0 && retrieved.Exists(r => r.CR_ID == 2 && r.Group == "АЭП" && r.Email == "test@aep.com");
            Console.WriteLine(success ? "TestInsertAndRetrieve: Успех" : "TestInsertAndRetrieve: Провал");
        }

        private void TestUpdate()
        {
            var request = new EmailRequest { CR_ID = 3, Group = "Субподрядчик", Email = "test@sub.com" };
            _repository.Insert(request);
            var inserted = _repository.GetByCRId(3).Find(r => r.Email == "test@sub.com");
            if (inserted != null)
            {
                inserted.Email = "updated@sub.com";
                _repository.Update(inserted.ID, inserted);
                var updated = _repository.GetByCRId(3).Find(r => r.ID == inserted.ID);
                bool success = updated != null && updated.Email == "updated@sub.com";
                Console.WriteLine(success ? "TestUpdate: Успех" : "TestUpdate: Провал");
            }
            else
            {
                Console.WriteLine("TestUpdate: Провал - не удалось найти вставленную запись");
            }
        }

        private void TestDelete()
        {
            var request = new EmailRequest { CR_ID = 4, Group = "Венгерский Заказчик", Email = "test@hungary.com" };
            _repository.Insert(request);
            var inserted = _repository.GetByCRId(4).Find(r => r.Email == "test@hungary.com");
            if (inserted != null)
            {
                _repository.Delete(inserted.ID);
                var retrieved = _repository.GetByCRId(4);
                bool success = retrieved.All(r => r.Email != "test@hungary.com");
                Console.WriteLine(success ? "TestDelete: Успех" : "TestDelete: Провал");
            }
            else
            {
                Console.WriteLine("TestDelete: Провал - не удалось найти вставленную запись");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string dbPath = @"C:\Users\39800323\RiderProjects\ConsoleApplication1\Program\test1.db";
            string logFilePath = @"C:\Users\39800323\RiderProjects\ConsoleApplication1\Program\email_logs.txt";
            var repository = new SQLiteEmailRepository(dbPath, logFilePath);
            var service = new EmailService(repository, logFilePath);

            // Запуск тестов
            var tests = new EmailServiceTests(repository, logFilePath);
            tests.RunTests();

            // Пример использования с вводом email вручную
            Console.WriteLine("\nПример создания записи:");
            Console.WriteLine("Доступные группы: Филиал АСЭ в Венгрии, АЭП, Субподрядчик, Венгерский Заказчик");
            Console.Write("Введите номер изменения (CR_ID, только целое число): ");
            int crId;
            while (!int.TryParse(Console.ReadLine(), out crId))
            {
                Console.WriteLine("Ошибка: введите целое число для CR_ID.");
                Console.Write("Введите номер изменения (CR_ID, только целое число): ");
            }

            Console.Write("Введите группу: ");
            string group = Console.ReadLine();

            Console.Write("Введите email: ");
            string email = Console.ReadLine();

            Console.Write("Введите срок (например, 20.03.2025, оставьте пустым, если не требуется): ");
            string deadline = Console.ReadLine();

            var request = new EmailRequest { CR_ID = crId, Group = group, Email = email };

            if (service.ValidateRequest(request, out string error))
            {
                repository.Insert(request);
                service.SendEmailsConsole(crId, deadline);
            }
            else
            {
                Console.WriteLine($"Ошибка: {error}");
            }
        }
    }
}