using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TypicalReply.Config
{
    public class TypicalReplyConfig
    {
        public string RibbonLabel { get; set; } = "Typical Reply";
        public List<TemplateConfig> TemplateConfigList { get; set; }
    }
}
