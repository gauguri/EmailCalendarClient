using System;

namespace GraphEmailClient.Contacts
{
    public class ContactRecord
    {
        public long Id { get; set; }

        public string Company { get; set; } = string.Empty;

        public string PointPerson { get; set; } = string.Empty;

        public string Email { get; set; } = string.Empty;

        public DateTime? LastEmailSentUtc { get; set; }
    }
}
