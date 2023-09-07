using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TypicalReply.Config
{
    internal class TemplateConfig
    {
        internal string Id { get; set; }
        internal string Label { get; set; }
        internal string AccessKey { get; set; }
        internal string SubjectPrefix { get; set; }
        internal string Subject { get; set; }
        internal string Body { get; set; }
        internal string BodyImage { get; set; }
        internal string Recipients { get; set; }
        internal string QuoteType { get; set; }
        internal string AllowedDomains { get; set; }
        internal string Icon { get; set; }
        internal string Locale { get; set; }
    }
}
