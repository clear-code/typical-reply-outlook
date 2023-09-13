using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TypicalReply.Config
{
    public class Config
    {
        public string Culture { get; set; }
        public string GalleryLabel { get; set; } = "Typical Reply";
        public List<ButtonConfig> ButtonConfigList { get; set; }
    }
}
