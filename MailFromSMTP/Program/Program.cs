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
    /// Сущность для отправляемого/хранимого email-запроса
    /// </summary>
    public class EmailRequest
    {
        public int ID { get; set; }
        public string CR_ID { get; set; }    
        public string Group { get; set; }
        public string Email { get; set; }
    }

    /// <summary>
    /// Интерфейс репозитория, расширенный для входящих ответов
    /// </summary>
    public interface IEmailRepository
    {
        void Insert(EmailRequest request);
        void Update(int id, EmailRequest request);
        void Delete(int id);
        List<EmailRequest> GetByCRId(string crId);

        // Методы для писем (Letters) и историй (ChM_Approval_history)
        int InsertLetter(string crId, string letterNumber, DateTime sentDate);
        void InsertApprovalHistory(string crId, int letterId, int statusId);
        void MarkLetterAsDelivered(int letterId);

        // Запись результата входящего письма (статус, комментарий и пр.)
        void InsertIncomingResponse(string fromEmail, string crId, string status, string comment);
    }

    /// <summary>
    /// Заглушечный репозиторий (Stub) – не требует настоящей БД
    /// </summary>
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

    /// <summary>
    /// Реальный репозиторий для работы с SQLite (нужны таблицы соответствующей структуры).
    /// </summary>
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

    //======================================================
    // Новая часть: IEmailSender, StubEmailSender, SmtpEmailSender
    //======================================================
    public interface IEmailSender
    {
        void SendEmail(string to, string subject, string body);
    }

    /// <summary>
    /// Заглушка для отправки писем – чтобы не подключаться к реальному SMTP.
    /// </summary>
    public class StubEmailSender : IEmailSender
    {
        public void SendEmail(string to, string subject, string body)
        {
            Console.WriteLine("[StubEmailSender] Письмо не отправляется по сети, только логируем...");
            Console.WriteLine("  Кому: {0}", to);
            Console.WriteLine("  Тема: {0}", subject);
            Console.WriteLine("  Тело:\n{0}", body);
        }
    }

    /// <summary>
    /// Реальная отправка писем через SmtpClient
    /// </summary>
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
            using (var mail = new MailMessage())
            {
                mail.From = new MailAddress(_user);
                mail.To.Add(to);
                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = false;
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;

                var smtp = new SmtpClient(_host, _port);
                smtp.EnableSsl = _enableSsl;
                smtp.Credentials = new NetworkCredential(_user, _password);
                smtp.Send(mail);
                
            }
        }
    }


    /// <summary>
    /// Сервис для работы с EmailRequest + обработка входящих.
    /// Теперь EmailService вызывает не напрямую SmtpClient, а _emailSender.
    /// </summary>
    public class EmailService
    {
        private readonly IEmailRepository _repository;
        private readonly IEmailSender _emailSender; // <-- добавлено
        private readonly Dictionary<string, string> _groupTemplates;
        private readonly Dictionary<string, string> _groupLinks;
        private readonly string _logFilePath;

        public EmailService(IEmailRepository repository, IEmailSender emailSender, string logFilePath)
        {
            _repository = repository;
            _emailSender = emailSender;  // <-- сохраняем реализацию
            _logFilePath = logFilePath;

            _groupTemplates = new Dictionary<string, string>
            {
                { "Филиал АСЭ в Венгрии",
                    "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\n" +
                    "Прошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\n" +
                    "Ссылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, …"
                },
                { "АЭП",
                    "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\n" +
                    "Прошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\n" +
                    "Ссылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, …"
                },
                { "Субподрядчик",
                    "Добрый день!\n\nНаправляю Вам на рассмотрение Запрос на изменение № {0}.\n" +
                    "Прошу Вас организовать оперативное рассмотрение и проработку указанных материалов.\n" +
                    "Ссылка на материалы: {1}\nПрошу рассмотреть и направить ОС в срок до: {2}\n\nС уважением, …"
                },
                { "Венгерский Заказчик",
                    "Dear Sir,\n\nI am sending you Change Request No. {0} for your information.\n" +
                    "Link to materials: {1}\n\nBest regards, …"
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

        public void SendEmailsConsole(string crId, string deadline = "")
        {
            var emailRequests = _repository.GetByCRId(crId);
            List<string> failedEmails = new List<string>();
            StringBuilder resultMessage = new StringBuilder("Рассылка произведена");
            StringBuilder logBuilder = new StringBuilder();

            foreach (var request in emailRequests)
            {
                string template = _groupTemplates.ContainsKey(request.Group)
                    ? _groupTemplates[request.Group]
                    : "Нет шаблона для группы " + request.Group;

                string link = _groupLinks.ContainsKey(request.Group)
                    ? string.Format(_groupLinks[request.Group], crId)
                    : "[Нет ссылки]";

                string subject = (request.Group == "Венгерский Заказчик")
                    ? string.Format("Change Request No. {0} notification", crId)
                    : string.Format("О рассмотрении и согласовании Запросов на изменения № {0}", crId);

                string body = (request.Group == "Венгерский Заказчик")
                    ? string.Format(template, crId, link)
                    : string.Format(template, crId, link, deadline);

                // Добавляем блок с ячейками Approved / Rejected / Comments
                body +=
                    "\n\n" +
                    "====================\n" +
                    "Инструкция по ответу:\n" +
                    "Пожалуйста, заполните ТОЛЬКО одно из следующих полей (после двоеточия):\n" +
                    "Approved:\n" +
                    "Rejected:\n" +
                    "Comments:\n" +
                    "\n" +
                    "Если заполнены два и более полей, это будет считаться ошибкой.\n" +
                    "====================\n";

                logBuilder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Начало отправки письма:");
                logBuilder.AppendLine($"Кому: {request.Email}");
                logBuilder.AppendLine($"Тема: {subject}");
                logBuilder.AppendLine($"Тело письма:\n{body}");

                try
                {
                    // Вызов именно через _emailSender, а не напрямую SmtpClient
                    _emailSender.SendEmail(request.Email, subject, body);

                    // Сохраняем письмо (Letters) + ApprovalHistory (Status=8)
                    var letterNumber = string.Format("CR{0}_{1:yyyyMMdd}", request.CR_ID, DateTime.Now);
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
                    " не доставлены, т.к. адрес не существует или SMTP недоступен."
                );
            }

            Console.WriteLine(resultMessage.ToString());
            File.AppendAllText(_logFilePath, logBuilder.ToString());
        }

        // ----- Методы для входящей почты -----

        public void ProcessIncomingEmail(string fromAddress, string subject, string body)
        {
            string crId = ExtractCrIdFromSubject(subject);

            string approvedVal = ParseFieldValue(body, "Approved:");
            string rejectedVal = ParseFieldValue(body, "Rejected:");
            string commentsVal = ParseFieldValue(body, "Comments:");

            int nonEmptyCount = 0;
            if (!IsNullOrWhiteSpace(approvedVal)) nonEmptyCount++;
            if (!IsNullOrWhiteSpace(rejectedVal)) nonEmptyCount++;
            if (!IsNullOrWhiteSpace(commentsVal)) nonEmptyCount++;

            if (nonEmptyCount == 0)
            {
                // Отправляем письмо-ошибку
                _emailSender.SendEmail(fromAddress, "Ошибка в ответе по CR " + crId,
                    "Не найдено заполненных полей Approved:/Rejected:/Comments:.\n" +
                    "Пожалуйста, заполните хотя бы одно поле и пришлите заново.");
                return;
            }
            if (nonEmptyCount > 1)
            {
                // Тоже письмо-ошибка
                _emailSender.SendEmail(fromAddress, "Ошибка в ответе по CR " + crId,
                    "Вы заполнили более одного поля (Approved, Rejected, Comments).\n" +
                    "Пожалуйста, оставьте только ОДНО значение и перешлите письмо снова.");
                return;
            }

            // Определяем статус
            string finalStatus;
            string comment = "";
            if (!IsNullOrWhiteSpace(approvedVal))
            {
                finalStatus = "Approved";
            }
            else if (!IsNullOrWhiteSpace(rejectedVal))
            {
                finalStatus = "Rejected";
            }
            else
            {
                finalStatus = "Comments";
                comment = commentsVal.Trim();
            }

            // Записываем в базу
            _repository.InsertIncomingResponse(fromAddress, crId, finalStatus, comment);

            // Отправляем подтверждение
            _emailSender.SendEmail(fromAddress, "Ответ по CR " + crId + " принят",
                "Статус: " + finalStatus + "\nКомментарий: " + comment + "\nСпасибо!");
        }

        private string ExtractCrIdFromSubject(string subject)
        {
            // Упрощенно – ищем последнее число
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
                line = text.Substring(idx).Trim();
            else
                line = text.Substring(idx, endIdx - idx).Trim();

            return line;
        }
    }

    /// <summary>
    /// Тесты для EmailService и репозитория
    /// </summary>
    public class EmailServiceTests
    {
        private readonly IEmailRepository _repository;
        private readonly EmailService _service;

        public EmailServiceTests(IEmailRepository repository, IEmailSender emailSender, string logFilePath)
        {
            _repository = repository;
            // Передаём emailSender в конструктор
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

        // -- те же тесты, что и раньше --

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
            // Сымитируем входящее письмо, где пользователь заполнил Rejected:
            string from = "some.user@domain.com";
            string subject = "RE: О рассмотрении и согласовании Запросов на изменения № 123";
            string body =
                "Approved:\n" +
                "Rejected:   Да\n" +
                "Comments:\n";

            // Проверяем обработку
            _service.ProcessIncomingEmail(from, subject, body);

            Console.WriteLine("TestProcessIncomingEmail: завершён.");
        }
    }

    /// <summary>
    /// Точка входа
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            string dbPath = @"C:\Users\39800323\RiderProjects\MailFromSMTP\Program\test1.db";
            string logFilePath = @"C:\Users\39800323\RiderProjects\MailFromSMTP\Program\email_logs.txt";

            // 1) Репозиторий (Stub или SQLite)
            IEmailRepository repository = new StubEmailRepository();
            // IEmailRepository repository = new SQLiteEmailRepository(dbPath, logFilePath);

            // 2) Отправка писем (Stub или реальный SMTP)
            IEmailSender emailSender = new StubEmailSender();
            // IEmailSender emailSender = new SmtpEmailSender("smtp.mycompany.com", 25, "no-reply@mycompany.com", "password", false);

            // Создаём сервис через EmailServiceTests, где мы передаём и repository, и emailSender
            var tests = new EmailServiceTests(repository, emailSender, logFilePath);
            tests.RunTests();

            // Пример использования
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

            // Теперь, чтобы вручную вызвать сервис, нам нужно тоже создать экземпляр EmailService
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
