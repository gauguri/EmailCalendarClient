using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Microsoft.Data.Sqlite;

namespace GraphEmailClient.Contacts
{
    public class ContactRepository
    {
        private readonly string _databasePath;
        private readonly string _connectionString;

        public ContactRepository()
        {
            var appDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GraphEmailClient");
            if (!Directory.Exists(appDirectory))
            {
                Directory.CreateDirectory(appDirectory);
            }

            _databasePath = Path.Combine(appDirectory, "contacts.db");
            _connectionString = $"Data Source={_databasePath}";

            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS Contacts (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Company TEXT,
                    PointPerson TEXT,
                    Email TEXT NOT NULL UNIQUE,
                    LastEmailSent TEXT
                );";
            command.ExecuteNonQuery();
        }

        public IReadOnlyList<ContactRecord> GetContacts()
        {
            var contacts = new List<ContactRecord>();

            using var connection = new SqliteConnection(_connectionString);
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText = "SELECT Id, Company, PointPerson, Email, LastEmailSent FROM Contacts ORDER BY Company COLLATE NOCASE";

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                var lastEmailSentRaw = reader.IsDBNull(4) ? null : reader.GetString(4);
                DateTime? lastEmailSent = null;
                if (!string.IsNullOrWhiteSpace(lastEmailSentRaw) &&
                    DateTime.TryParse(lastEmailSentRaw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed))
                {
                    lastEmailSent = parsed;
                }

                contacts.Add(new ContactRecord
                {
                    Id = reader.GetInt64(0),
                    Company = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                    PointPerson = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                    Email = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                    LastEmailSentUtc = lastEmailSent
                });
            }

            return contacts;
        }

        public int UpsertContacts(IEnumerable<ContactRecord> contacts)
        {
            if (contacts == null)
            {
                return 0;
            }

            using var connection = new SqliteConnection(_connectionString);
            connection.Open();
            using var transaction = connection.BeginTransaction();

            var upsertCommand = connection.CreateCommand();
            upsertCommand.Transaction = transaction;
            upsertCommand.CommandText = @"
                INSERT INTO Contacts (Company, PointPerson, Email, LastEmailSent)
                VALUES ($company, $pointPerson, $email, $lastEmail)
                ON CONFLICT(Email) DO UPDATE SET
                    Company = excluded.Company,
                    PointPerson = excluded.PointPerson;";

            var companyParam = upsertCommand.CreateParameter();
            companyParam.ParameterName = "$company";
            upsertCommand.Parameters.Add(companyParam);

            var pointPersonParam = upsertCommand.CreateParameter();
            pointPersonParam.ParameterName = "$pointPerson";
            upsertCommand.Parameters.Add(pointPersonParam);

            var emailParam = upsertCommand.CreateParameter();
            emailParam.ParameterName = "$email";
            upsertCommand.Parameters.Add(emailParam);

            var lastEmailParam = upsertCommand.CreateParameter();
            lastEmailParam.ParameterName = "$lastEmail";
            upsertCommand.Parameters.Add(lastEmailParam);

            var processed = 0;
            foreach (var contact in contacts)
            {
                processed++;
                companyParam.Value = string.IsNullOrWhiteSpace(contact.Company) ? (object)DBNull.Value : contact.Company;
                pointPersonParam.Value = string.IsNullOrWhiteSpace(contact.PointPerson) ? (object)DBNull.Value : contact.PointPerson;
                emailParam.Value = string.IsNullOrWhiteSpace(contact.Email) ? throw new InvalidOperationException("Email is required for a contact record.") : contact.Email.Trim();
                lastEmailParam.Value = contact.LastEmailSentUtc.HasValue ? contact.LastEmailSentUtc.Value.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture) : (object)DBNull.Value;

                upsertCommand.ExecuteNonQuery();
            }

            transaction.Commit();
            return processed;
        }

        public void UpdateLastEmailSent(string email, DateTime timestampUtc)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return;
            }

            var normalizedEmail = email.Trim();
            var timestamp = timestampUtc.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);

            using var connection = new SqliteConnection(_connectionString);
            connection.Open();

            using var updateCommand = connection.CreateCommand();
            updateCommand.CommandText = "UPDATE Contacts SET LastEmailSent = $lastEmail WHERE Email = $email";
            updateCommand.Parameters.AddWithValue("$lastEmail", timestamp);
            updateCommand.Parameters.AddWithValue("$email", normalizedEmail);

            var affected = updateCommand.ExecuteNonQuery();

            if (affected == 0)
            {
                using var insertCommand = connection.CreateCommand();
                insertCommand.CommandText = @"
                    INSERT INTO Contacts (Email, LastEmailSent)
                    VALUES ($email, $lastEmail);";
                insertCommand.Parameters.AddWithValue("$email", normalizedEmail);
                insertCommand.Parameters.AddWithValue("$lastEmail", timestamp);
                insertCommand.ExecuteNonQuery();
            }
        }
    }
}
