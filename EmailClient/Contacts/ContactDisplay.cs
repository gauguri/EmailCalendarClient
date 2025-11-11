using System;

namespace GraphEmailClient.Contacts
{
    public class ContactDisplay
    {
        private static readonly TimeSpan FollowUpThreshold = TimeSpan.FromDays(7);

        public string Company { get; set; } = string.Empty;

        public string PointPerson { get; set; } = string.Empty;

        public string Email { get; set; } = string.Empty;

        public DateTime? LastEmailSentUtc { get; set; }

        public string LastEmailDisplay => LastEmailSentUtc.HasValue
            ? LastEmailSentUtc.Value.ToLocalTime().ToString("g")
            : "Never";

        public string FollowUpStatus
        {
            get
            {
                if (!LastEmailSentUtc.HasValue)
                {
                    return "Never contacted";
                }

                var delta = DateTime.UtcNow - LastEmailSentUtc.Value;
                if (delta >= FollowUpThreshold)
                {
                    var days = Math.Max(1, (int)Math.Round(delta.TotalDays));
                    return $"Follow-up overdue ({days} days)";
                }

                return "Recently contacted";
            }
        }

        public bool FollowUpDue => !LastEmailSentUtc.HasValue || (DateTime.UtcNow - LastEmailSentUtc.Value) >= FollowUpThreshold;

        public static ContactDisplay FromRecord(ContactRecord record)
        {
            return new ContactDisplay
            {
                Company = record.Company,
                PointPerson = record.PointPerson,
                Email = record.Email,
                LastEmailSentUtc = record.LastEmailSentUtc
            };
        }
    }
}
