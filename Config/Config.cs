﻿using System;
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
        public string GroupLabel { get; set; } = "Typical Reply";
        public string TabMailInsertAfterMso { get; set; } = "GroupMailRespond";
        public string TabReadInsertAfterMso { get; set; } = "GroupRespond";
        public string ContextMenuInsertAfterMso { get; set; } = "Forward";

        public List<ButtonConfig> ButtonConfigList { get; set; }

        /// <summary>
        /// Merge with other.
        /// Note that this method doesn't check Culture.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        internal Config Merge(Config other)
        {
            if (other is null)
            {
                return this;
            }
            if (!string.IsNullOrEmpty(other.GroupLabel))
            {
                this.GroupLabel = other.GroupLabel;
            }
            if (this.ButtonConfigList is null)
            {
                this.ButtonConfigList = other.ButtonConfigList;
            }
            else if (other.ButtonConfigList != null)
            {                
                foreach (ButtonConfig otherButtonConfig in other.ButtonConfigList)
                {
                    int index = this.ButtonConfigList.FindIndex(_ => _.Id == otherButtonConfig.Id);
                    if (index >= 0)
                    {
                        this.ButtonConfigList[index] = otherButtonConfig;
                    }
                    else
                    {
                        this.ButtonConfigList.Add(otherButtonConfig);
                    }
                }
            }
            return this;
        }
    }
}
