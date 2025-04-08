using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Text;
using System.Linq; // Added for LINQ's Select

namespace ConsoleApplication2
{
    public class ChangeData
    {
        public int ChangeId { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Initiator { get; set; }
        public DateTime CreatedDate { get; set; }
    }

    public class EmailService
    {
        public List<string> SendNotifications(int changeId, List<int> selectedRecipientIds, ChangeData changeData)
        {
            return new List<string>(); // Dummy implementation
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Запуск демонстрации сервиса отправки email");
                
                EmailService emailService = new EmailService();
                DemoSendNotifications(emailService);
                
                Console.WriteLine("Демонстрация завершена. Нажмите любую клавишу для выхода...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                Console.ReadKey();
            }
        }

        /// <summary>
        /// Демонстрация отправки уведомлений
        /// </summary>
        static void DemoSendNotifications(EmailService emailService)
        {
            Console.WriteLine("=== Демонстрация отправки уведомлений ===");
            
            try
            {
                int changeId = 123;
                List<int> selectedRecipientIds = new List<int> { 1, 2, 3 };
                
                ChangeData changeData = new ChangeData
                {
                    ChangeId = changeId,
                    Title = "Изменение конструкции насосной станции",
                    Description = "Изменение конфигурации трубопроводов согласно новым требованиям",
                    Initiator = "Иванов И.И.",
                    CreatedDate = DateTime.Now
                };
                
                // Fix: Convert List<int> to string[] using Select
                string recipientsList = string.Join(", ", selectedRecipientIds.Select(id => id.ToString()).ToArray());
                Console.WriteLine($"Отправка уведомлений по изменению #{changeId} для {selectedRecipientIds.Count} получателей: {recipientsList}");
                
                List<string> failedEmails = emailService.SendNotifications(changeId, selectedRecipientIds, changeData);
                
                if (failedEmails.Count == 0)
                {
                    Console.WriteLine("Все уведомления отправлены успешно");
                }
                else
                {
                    string failedEmailsList = string.Join(", ", failedEmails.ToArray());
                    Console.WriteLine($"Не удалось отправить уведомления на следующие адреса: {failedEmailsList}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при отправке уведомлений: {ex.Message}");
            }
            
            Console.WriteLine();
        }
    }
}