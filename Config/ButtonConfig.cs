using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace TypicalReply.Config
{
    public enum ForwardType
    {
        Unknown = 0,
        Attachment = 1,
        Inline = 2,
    }

    public enum RecipientsType
    {
        Unknown = 0,
        Blank = 1,
        Sender = 2,
        All = 3,
        UserSpecification = 4
    }

    public enum AllowedDomainsType
    {
        Unknown = 0,
        All = 1,
        UserSpecification = 2
    }

    /// <summary>
    /// Config for drop down item
    /// </summary>
    public class ButtonConfig
    {
        public string Id { get; set; }
        public string Label { get; set; }

        /// <summary>
        /// Shortcut key for this item
        /// This is not work for now.
        /// </summary>
        //public string AccessKey { get; set; }
        public string SubjectPrefix { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public List<string> Recipients { get; set; }

        [JsonIgnore]
        public RecipientsType RecipientsType
        {
            get
            {
                if (Recipients == null || !Recipients.Any())
                {
                    return RecipientsType.Blank;
                }
                RecipientsType receipients;
                if (Enum.TryParse(this.Recipients[0], true, out receipients))
                {
                    // If "UserSpecification" is specified, this returns Recipients.UserSpecification.
                    // In that case, users may actually specify "UserSpecification", so this works fine.
                    return receipients;
                }
                return RecipientsType.UserSpecification;
            }
        }

        public bool QuoteType { get; set; }
        public List<string> AllowedDomains { get; set; }

        [JsonIgnore]
        public List<string> LoweredAllowedDomains => AllowedDomains.Select(_ => _.ToLowerInvariant()).ToList();

        [JsonIgnore]
        public AllowedDomainsType AllowedDomainsType
        {
            get
            {
                if (AllowedDomains == null || 
                    !AllowedDomains.Any() || 
                    AllowedDomains.First() == "*")
                {
                    return AllowedDomainsType.All;
                }
                return AllowedDomainsType.UserSpecification;
            }
        }
        public ForwardType ForwardType { get; set; }
    }
}
