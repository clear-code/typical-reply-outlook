using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TypicalReply.Config
{
    public class TemplateConfig
    {
        public string Id { get; set; }
        public string Label { get; set; }
        public string AccessKey { get; set; }
        public string SubjectPrefix { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string BodyImage { get; set; }
        public string Recipients { get; set; }
        public string QuoteType { get; set; }
        public string AllowedDomains { get; set; }
        public string Icon { get; set; }
        public string Locale { get; set; }
    }
}
