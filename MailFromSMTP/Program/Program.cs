using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Linq;

namespace ChangeManagementSystem
{
    /// <summary>
    /// Сущность для хранения данных об отправляемом письме
    /// </summary>
    public class EmailRequest
    {
        public int ID { get; set; }
        public string CR_ID { get; set; }    
        public string Group { get; set; }
        public string Email { get; set; }
    }

    // -------------------------------------------------------------------------
    // Интерфейс + реализации репозитория (заглушка и SQLite)
    // -------------------------------------------------------------------------
    public interface IEmailRepository
    {
        void Insert(EmailRequest request);
        void Update(int id, EmailRequest request);
        void Delete(int id);
        List<EmailRequest> GetByCRId(string crId);

        int InsertLetter(string crId, string letterNumber, DateTime sentDate);
        void InsertApprovalHistory(string crId, int letterId, int statusId);
        void MarkLetterAsDelivered(int letterId);

        void InsertIncomingResponse(string fromEmail, string crId, string status, string comment);
    }

    public class StubEmailRepository : IEmailRepository
    {
        private readonly List<EmailRequest> _requests = new List<EmailRequest>();
        private int _autoIncrement = 1;

        public void Insert(EmailRequest request)
        {
            Console.WriteLine($"[StubRepository] Insert called for CR_ID={request.CR_ID}, Group={request.Group}, Email={request.Email}");
            request.ID = _autoIncrement++;
            _requests.Add(request);
        }

        public void Update(int id, EmailRequest request)
        {
            Console.WriteLine($"[StubRepository] Update called for ID={id}, CR_ID={request.CR_ID}, Group={request.Group}, Email={request.Email}");
            var existing = _requests.Find(r => r.ID == id);
            if (existing != null)
            {
                existing.CR_ID = request.CR_ID;
                existing.Group = request.Group;
                existing.Email = request.Email;
            }
        }

        public void Delete(int id)
        {
            Console.WriteLine($"[StubRepository] Delete called for ID={id}");
            _requests.RemoveAll(r => r.ID == id);
        }

        public List<EmailRequest> GetByCRId(string crId)
        {
            Console.WriteLine($"[StubRepository] GetByCRId called for CR_ID={crId}");
            return _requests.Where(r => r.CR_ID == crId).ToList();
        }

        public int InsertLetter(string crId, string letterNumber, DateTime sentDate)
        {
            Console.WriteLine($"[StubRepository] InsertLetter called: CR_ID={crId}, LetterNumber={letterNumber}, SentDate={sentDate}");
            // Возвращаем некий рандомный ID
            return new Random().Next(100, 999);
        }

        public void InsertApprovalHistory(string crId, int letterId, int statusId)
        {
            Console.WriteLine($"[StubRepository] InsertApprovalHistory called: CR_ID={crId}, Letter_ID={letterId}, Status_ID={statusId}");
        }

        public void MarkLetterAsDelivered(int letterId)
        {
            Console.WriteLine($"[StubRepository] MarkLetterAsDelivered called: Letter_ID={letterId}");
        }

        public void InsertIncomingResponse(string fromEmail, string crId, string status, string comment)
        {
            Console.WriteLine($"[StubRepository] InsertIncomingResponse: from={fromEmail}, CR_ID={crId}, Status={status}, Comment={comment}");
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

        public List<EmailRequest> GetByCRId(string crId)
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
                                        CR_ID = reader["CR_ID"] == DBNull.Value ? "" : reader["CR_ID"].ToString(),
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

        public int InsertLetter(string crId, string letterNumber, DateTime sentDate)
        {
            int newId = 0;

            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало операции InsertLetter");
            using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                string query = @"
                    INSERT INTO Letters (CR_ID, LetterNumber, SentDate)
                    VALUES (@CR_ID, @LetterNumber, @SentDate);
                    SELECT last_insert_rowid();";
                logBuilder.AppendLine($"Выполняется запрос: {query}");
                logBuilder.AppendLine($"Параметры: CR_ID={crId}, LetterNumber={letterNumber}, SentDate={sentDate}");

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@CR_ID", crId);
                    cmd.Parameters.AddWithValue("@LetterNumber", letterNumber);
                    cmd.Parameters.AddWithValue("@SentDate", sentDate.ToString("yyyy-MM-dd HH:mm:ss"));

                    object result = cmd.ExecuteScalar();
                    if (result != null)
                    {
                        newId = Convert.ToInt32(result);
                    }
                    logBuilder.AppendLine($"Результат: LetterID={newId}");
                }
            }
            logBuilder.AppendLine("----------------------------------------");
            Console.WriteLine(logBuilder.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());

            return newId;
        }

        public void InsertApprovalHistory(string crId, int letterId, int statusId)
        {
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало операции InsertApprovalHistory");

            using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                string query = @"
                    INSERT INTO ChM_Approval_history (CR_ID, Letter_ID, Status_ID, [Date])
                    VALUES (@CR_ID, @Letter_ID, @Status_ID, @Date)";

                logBuilder.AppendLine($"Выполняется запрос: {query}");
                logBuilder.AppendLine($"Параметры: CR_ID={crId}, Letter_ID={letterId}, Status_ID={statusId}, Date={DateTime.Now}");

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@CR_ID", crId);
                    cmd.Parameters.AddWithValue("@Letter_ID", letterId);
                    cmd.Parameters.AddWithValue("@Status_ID", statusId);
                    cmd.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    int rowsAffected = cmd.ExecuteNonQuery();
                    logBuilder.AppendLine($"Результат: Вставлено строк - {rowsAffected}");
                }
            }
            logBuilder.AppendLine("----------------------------------------");
            Console.WriteLine(logBuilder.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }

        public void MarkLetterAsDelivered(int letterId)
        {
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало операции MarkLetterAsDelivered");

            using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                string query = @"
                    UPDATE ChM_Approval_history
                    SET Status_ID = 2, [Date] = @Date
                    WHERE Letter_ID = @Letter_ID";

                logBuilder.AppendLine($"Выполняется запрос: {query}");
                logBuilder.AppendLine($"Параметры: Letter_ID={letterId}, Status_ID=2");

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Letter_ID", letterId);
                    cmd.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    int rowsAffected = cmd.ExecuteNonQuery();
                    logBuilder.AppendLine($"Результат: Обновлено строк - {rowsAffected}");
                }
            }
            logBuilder.AppendLine("----------------------------------------");
            Console.WriteLine(logBuilder.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }

        public void InsertIncomingResponse(string fromEmail, string crId, string status, string comment)
        {
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало операции InsertIncomingResponse");

            using (SQLiteConnection conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();
                string query = @"
                    INSERT INTO ChM_Incoming_Responses (FromEmail, CR_ID, Status, Comment, ReceivedDate)
                    VALUES (@FromEmail, @CR_ID, @Status, @Comment, @ReceivedDate)";

                logBuilder.AppendLine($"Выполняется запрос: {query}");
                logBuilder.AppendLine($"Параметры: fromEmail={fromEmail}, CR_ID={crId}, Status={status}, Comment={comment}");

                using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@FromEmail", fromEmail);
                    cmd.Parameters.AddWithValue("@CR_ID", crId);
                    cmd.Parameters.AddWithValue("@Status", status);
                    cmd.Parameters.AddWithValue("@Comment", comment ?? "");
                    cmd.Parameters.AddWithValue("@ReceivedDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    int rowsAffected = cmd.ExecuteNonQuery();
                    logBuilder.AppendLine($"Результат: Вставлено строк - {rowsAffected}");
                }
            }
            logBuilder.AppendLine("----------------------------------------");
            Console.WriteLine(logBuilder.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }
    }

    // -------------------------------------------------------------------------
    // Интерфейс + реализации для отправки писем (заглушка и реальная)
    // -------------------------------------------------------------------------
    public interface IEmailSender
    {
        void SendEmail(string to, string subject, string body);
    }

    public class StubEmailSender : IEmailSender
    {
        public void SendEmail(string to, string subject, string body)
        {
            Console.WriteLine("[StubEmailSender] Письмо не отправляем, только логируем:");
            Console.WriteLine("  Кому: {0}", to);
            Console.WriteLine("  Тема: {0}", subject);
            Console.WriteLine("  Тело:\n{0}", body);
        }
    }

    public class SmtpEmailSender : IEmailSender
    {
        private readonly string _host;
        private readonly int _port;
        private readonly string _user;
        private readonly string _password;
        private readonly bool _enableSsl;

        public SmtpEmailSender(string host, int port, string user, string password, bool enableSsl)
        {
            _host = host;
            _port = port;
            _user = user;
            _password = password;
            _enableSsl = enableSsl;
        }

        public void SendEmail(string to, string subject, string body)
        {
            // В .NET 3.5 SmtpClient не IDisposable
            var smtp = new SmtpClient(_host, _port);
            smtp.EnableSsl = _enableSsl;
            smtp.Credentials = new NetworkCredential(_user, _password);

            using (var mail = new MailMessage())
            {
                mail.From = new MailAddress(_user);
                mail.To.Add(to);
                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = false;
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;

                smtp.Send(mail);
            }
        }
    }

    // -------------------------------------------------------------------------
    // EmailService: "Approved:" / "Rejected:" и двуязычная инструкция
    // -------------------------------------------------------------------------
    public class EmailService
    {
        private readonly IEmailRepository _repository;
        private readonly IEmailSender _emailSender;
        private readonly Dictionary<string, string> _groupTemplates;
        private readonly Dictionary<string, string> _groupLinks;
        private readonly string _logFilePath;

        public EmailService(IEmailRepository repository, IEmailSender emailSender, string logFilePath)
        {
            _repository = repository;
            _emailSender = emailSender;
            _logFilePath = logFilePath;

            _groupTemplates = new Dictionary<string, string>
            {
                // Русские группы
                { "Филиал АСЭ в Венгрии",
                    "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\n" +
                    "Прошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\n" +
                    "Ссылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, УКИТ"
                },
                { "АЭП",
                    "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\n" +
                    "Прошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\n" +
                    "Ссылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, УКИТ"
                },
                { "Субподрядчик",
                    "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\n" +
                    "Прошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\n" +
                    "Ссылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, УКИТ"
                },
                // Англ группа
                { "Венгерский Заказчик",
                    "Dear Sir,\n\nI am sending you Change Request No. {0} for your information.\n" +
                    "Link to materials: {1}\n\nBest regards, CCRMD"
                }
            };

            _groupLinks = new Dictionary<string, string>
            {
                { "Филиал АСЭ в Венгрии", "http://ase-hungary/change/{0}" },
                { "АЭП", "http://voshod/change/{0}" },
                { "Субподрядчик", "http://subcontractor/change/{0}" },
                { "Венгерский Заказчик", "http://ftp/change/{0}" }
            };
        }

        private static bool IsNullOrWhiteSpace(string s)
        {
            return s == null || s.Trim().Length == 0;
        }

        public bool ValidateRequest(EmailRequest request, out string error)
        {
            error = string.Empty;
            if (IsNullOrWhiteSpace(request.Group))
            {
                error = "Поле Group обязательно";
                return false;
            }
            if (IsNullOrWhiteSpace(request.Email))
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

        /// <summary>
        /// Отправка писем со строками Approved/Rejected и инструкцией (рус/англ).
        /// </summary>
        public void SendEmailsConsole(string crId, string deadline = "")
        {
            var emailRequests = _repository.GetByCRId(crId);
            List<string> failedEmails = new List<string>();
            StringBuilder resultMessage = new StringBuilder("Рассылка произведена");
            StringBuilder logBuilder = new StringBuilder();

            foreach (var request in emailRequests)
            {
                string template = _groupTemplates[request.Group];
                string link = string.Format(_groupLinks[request.Group], crId);

                string subject = (request.Group == "Венгерский Заказчик")
                    ? string.Format("Change Request No. {0} notification", crId)
                    : string.Format("О рассмотрении и согласовании Запросов на изменения № {0}", crId);

                // Формируем тело (англ. шаблон – 2 аргумента, русский – 3)
                string body;
                if (request.Group == "Венгерский Заказчик")
                {
                    body = string.Format(template, crId, link);
                }
                else
                {
                    body = string.Format(template, crId, link, deadline);
                }

                // Добавляем «таблицу» Approved/Rejected и инструкцию (RU/EN)
                body +=
                    "\n\n" +
                    "=====================================\n" +
                    "| Approved:                         |\n" +
                    "| Rejected:                         |\n" +
                    "=====================================\n" +
                    "\n"+
                    "Инструкция для ответа на письмо:\n"+
                    "При согласовании, напишите в строке \"Approved\" после двоеточия любой комментарий\n" +
                    "\t или \n" +
                    "при отлонении, напишите в строке \"Rejected\" после двоеточия причину отклонения (обязательно).\n" +
                    "----\n" +
                    "\n"+
                    "Instruction to reply to this e-mail:\n"+
                    "If approved, write any comment after the colon on the line \"Approved\"\n" +
                    "\t or \n"+
                    "If rejected, write the reason for rejection on the line \"Rejected\" after the colon (mandatory).\n" +
                    "----\n";

                logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало отправки письма:");
                logBuilder.AppendLine($"Кому: {request.Email}");
                logBuilder.AppendLine($"Тема: {subject}");
                logBuilder.AppendLine($"Тело письма:\n{body}");

                try
                {
                    _emailSender.SendEmail(request.Email, subject, body);

                    // Сохраняем письмо (letters) + статус 8
                    var letterNumber = "CR" + request.CR_ID + "_" + DateTime.Now.ToString("yyyyMMdd");
                    var sentDate = DateTime.Now;
                    int letterId = _repository.InsertLetter(request.CR_ID, letterNumber, sentDate);
                    _repository.InsertApprovalHistory(request.CR_ID, letterId, 8);

                    logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Письмо успешно отправлено на {request.Email}");
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
                resultMessage.Append(
                    "\nПисьма для " + string.Join(", ", failedEmails.ToArray()) +
                    " не доставлены (SMTP или адрес не существует)."
                );
            }

            Console.WriteLine(resultMessage.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }

        /// <summary>
        /// Обработка входящего письма (Approved/Rejected).
        /// Approved может быть просто "+", Rejected требует обязательного текста.
        /// Если обе строки или ни одна – ошибка.
        /// </summary>
        public void ProcessIncomingEmail(string fromAddress, string subject, string body)
        {
            string crId = ExtractCrIdFromSubject(subject);

            // Читаем содержимое строк:
            string approvedLine = ParseFieldValue(body, "Approved:");
            string rejectedLine = ParseFieldValue(body, "Rejected:");

            bool hasApproved = !IsNullOrWhiteSpace(approvedLine);
            bool hasRejected = !IsNullOrWhiteSpace(rejectedLine);

            if (hasApproved && hasRejected)
            {
                // Заполнены обе
                _emailSender.SendEmail(fromAddress, "Ошибка / Error with CR " + crId,
                    "Заполнены и Approved, и Rejected. Укажите только одну строку.\n" +
                    "Both Approved and Rejected are filled. Please fill only one line.");
                return;
            }
            if (!hasApproved && !hasRejected)
            {
                // Пусты обе
                _emailSender.SendEmail(fromAddress, "Ошибка / Error with CR " + crId,
                    "Не заполнены ни Approved, ни Rejected.\n" +
                    "Both Approved and Rejected are empty. Please fill exactly one.");
                return;
            }

            // Если есть Approved, считаем статус Approved. Любой текст (или просто +) - ок
            if (hasApproved)
            {
                string finalStatus = "Approved";
                string comment = approvedLine.Trim(); // Может быть "+" или текст

                _repository.InsertIncomingResponse(fromAddress, crId, finalStatus, comment);
                _emailSender.SendEmail(fromAddress, "Ответ по CR " + crId + " принят / Response accepted",
                    "Статус: Approved\nКомментарий: " + comment + "\nThank you!");
                return;
            }

            // Иначе Rejected
            // Но для Rejected требуется причина (не просто "+")
            // Если пользователь поставил только "+", это недостаточно
            string reason = rejectedLine.Trim();
            if (reason == "+")
            {
                // Ошибка, нужно больше текста
                _emailSender.SendEmail(fromAddress, "Ошибка / Error with CR " + crId,
                    "Для отклонения требуется указать причину.\n" +
                    "Rejected line must contain a reason, not just '+'.");
                return;
            }

            // Всё ок, считаем Rejected
            _repository.InsertIncomingResponse(fromAddress, crId, "Rejected", reason);
            _emailSender.SendEmail(fromAddress, "Ответ по CR " + crId + " принят / Response accepted",
                "Статус: Rejected\nПричина: " + reason + "\nThank you!");
        }

        private string ExtractCrIdFromSubject(string subject)
        {
            // Упрощённо ищем последнее число
            string crId = "";
            var parts = subject.Split(' ');
            foreach (var part in parts)
            {
                int dummy;
                if (int.TryParse(part, out dummy))
                {
                    crId = part;
                }
            }
            return crId;
        }

        private string ParseFieldValue(string text, string fieldName)
        {
            int idx = text.IndexOf(fieldName, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return "";

            idx += fieldName.Length;
            int endIdx = text.IndexOf('\n', idx);
            string line;
            if (endIdx < 0)
                line = text.Substring(idx);
            else
                line = text.Substring(idx, endIdx - idx);

            return line.Trim();
        }
    }

    // -------------------------------------------------------------------------
    // Тесты для EmailService, показывают работу логики
    // -------------------------------------------------------------------------
    public class EmailServiceTests
    {
        private readonly IEmailRepository _repository;
        private readonly EmailService _service;

        public EmailServiceTests(IEmailRepository repository, IEmailSender emailSender, string logFilePath)
        {
            _repository = repository;
            _service = new EmailService(repository, emailSender, logFilePath);
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
            TestInsertLettersAndApprovalHistory();
            TestMarkLetterAsDelivered();
            TestProcessIncomingEmail();

            Console.WriteLine("Тесты завершены.");
        }

        // --- Обычные тесты на репозиторий и валидацию ---

        private void TestValidationSuccess()
        {
            var request = new EmailRequest { CR_ID = "1", Group = "Филиал АСЭ в Венгрии", Email = "test@ase.com" };
            bool result = _service.ValidateRequest(request, out string error);
            Console.WriteLine(result && string.IsNullOrEmpty(error)
                ? "TestValidationSuccess: Успех"
                : "TestValidationSuccess: Провал - " + error);
        }

        private void TestValidationFailureEmptyGroup()
        {
            var request = new EmailRequest { CR_ID = "1", Group = "", Email = "test@ase.com" };
            bool result = _service.ValidateRequest(request, out string error);
            Console.WriteLine(!result && error == "Поле Group обязательно"
                ? "TestValidationFailureEmptyGroup: Успех"
                : "TestValidationFailureEmptyGroup: Провал - " + error);
        }

        private void TestValidationFailureEmptyEmail()
        {
            var request = new EmailRequest { CR_ID = "1", Group = "Филиал АСЭ в Венгрии", Email = "" };
            bool result = _service.ValidateRequest(request, out string error);
            Console.WriteLine(!result && error == "Поле Email обязательно"
                ? "TestValidationFailureEmptyEmail: Успех"
                : "TestValidationFailureEmptyEmail: Провал - " + error);
        }

        private void TestValidationFailureLongFields()
        {
            var longString = new string('a', 101);
            var request = new EmailRequest { CR_ID = "1", Group = longString, Email = "test@ase.com" };
            bool result = _service.ValidateRequest(request, out string error);
            Console.WriteLine(!result && error == "Превышена максимальная длина поля (100 символов)"
                ? "TestValidationFailureLongFields: Успех"
                : "TestValidationFailureLongFields: Провал - " + error);
        }

        private void TestInsertAndRetrieve()
        {
            var request = new EmailRequest { CR_ID = "2", Group = "АЭП", Email = "test@aep.com" };
            _repository.Insert(request);
            var retrieved = _repository.GetByCRId("2");
            bool success = (retrieved.Count > 0 && retrieved.Exists(r => r.CR_ID == "2" && r.Group == "АЭП" && r.Email == "test@aep.com"));
            Console.WriteLine(success ? "TestInsertAndRetrieve: Успех" : "TestInsertAndRetrieve: Провал");
        }

        private void TestUpdate()
        {
            var request = new EmailRequest { CR_ID = "3", Group = "Субподрядчик", Email = "test@sub.com" };
            _repository.Insert(request);

            var inserted = _repository.GetByCRId("3").Find(r => r.Email == "test@sub.com");
            if (inserted != null)
            {
                inserted.Email = "updated@sub.com";
                _repository.Update(inserted.ID, inserted);
                var updated = _repository.GetByCRId("3").Find(r => r.ID == inserted.ID);
                bool success = (updated != null && updated.Email == "updated@sub.com");
                Console.WriteLine(success ? "TestUpdate: Успех" : "TestUpdate: Провал");
            }
            else
            {
                Console.WriteLine("TestUpdate: Провал - не удалось найти вставленную запись");
            }
        }

        private void TestDelete()
        {
            var request = new EmailRequest { CR_ID = "4", Group = "Венгерский Заказчик", Email = "test@hungary.com" };
            _repository.Insert(request);
            var inserted = _repository.GetByCRId("4").Find(r => r.Email == "test@hungary.com");
            if (inserted != null)
            {
                _repository.Delete(inserted.ID);
                var retrieved = _repository.GetByCRId("4");
                bool success = retrieved.All(r => r.Email != "test@hungary.com");
                Console.WriteLine(success ? "TestDelete: Успех" : "TestDelete: Провал");
            }
            else
            {
                Console.WriteLine("TestDelete: Провал - не удалось найти вставленную запись");
            }
        }

        private void TestInsertLettersAndApprovalHistory()
        {
            Console.WriteLine("TestInsertLettersAndApprovalHistory: старт...");
            string crId = "10";
            string letterNumber = "CR" + crId + "_" + DateTime.Now.ToString("yyyyMMdd");
            DateTime sentDate = DateTime.Now;

            int letterId = _repository.InsertLetter(crId, letterNumber, sentDate);
            _repository.InsertApprovalHistory(crId, letterId, 8);

            Console.WriteLine("TestInsertLettersAndApprovalHistory: завершён.");
        }

        private void TestMarkLetterAsDelivered()
        {
            Console.WriteLine("TestMarkLetterAsDelivered: старт...");
            int letterId = 123;
            _repository.MarkLetterAsDelivered(letterId);
            Console.WriteLine("TestMarkLetterAsDelivered: письмо (ID=" + letterId + ") помечено как доставленное (Status_ID=2).");
        }

        private void TestProcessIncomingEmail()
        {
            Console.WriteLine("TestProcessIncomingEmail: старт...");
            // Сымитируем входящее письмо, где пользователь заполнил Rejected и указал причину
            string from = "some.user@domain.com";
            string subject = "RE: О рассмотрении и согласовании Запросов на изменения № 123";
            string body =
                "Approved:\n" +
                "Rejected:  Я против\n";

            _service.ProcessIncomingEmail(from, subject, body);
            Console.WriteLine("TestProcessIncomingEmail: завершён.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string dbPath = @"C:\Users\39800323\RiderProjects\MailFromSMTP\Program\test1.db";
            string logFilePath = @"C:\Users\39800323\RiderProjects\MailFromSMTP\Program\email_logs.txt";

            // Выбираем репозиторий:
            IEmailRepository repository = new StubEmailRepository();
            // Или: IEmailRepository repository = new SQLiteEmailRepository(dbPath, logFilePath);

            // Выбираем отправку писем:
            IEmailSender emailSender = new StubEmailSender();
            // Или: IEmailSender emailSender = new SmtpEmailSender("smtp.mycompany.com", 25, "user@company.com", "password", false);

            // Запускаем тесты
            var tests = new EmailServiceTests(repository, emailSender, logFilePath);
            tests.RunTests();

            // Демонстрация ручного ввода
            Console.WriteLine("\nПример создания записи:");
            Console.WriteLine("Доступные группы: Филиал АСЭ в Венгрии, АЭП, Субподрядчик, Венгерский Заказчик");

            Console.Write("Введите номер изменения (CR_ID, можно строку): ");
            string crId = Console.ReadLine();

            Console.Write("Введите группу: ");
            string group = Console.ReadLine();

            Console.Write("Введите email: ");
            string email = Console.ReadLine();

            Console.Write("Введите срок (например, 20.03.2025, оставьте пустым, если не требуется): ");
            string deadline = Console.ReadLine();

            var request = new EmailRequest { CR_ID = crId, Group = group, Email = email };

            var service = new EmailService(repository, emailSender, logFilePath);

            if (service.ValidateRequest(request, out string error))
            {
                repository.Insert(request);
                service.SendEmailsConsole(crId, deadline);
            }
            else
            {
                Console.WriteLine("Ошибка: " + error);
            }

            Console.WriteLine("Нажмите Enter для выхода...");
            Console.ReadLine();
        }
    }
}
