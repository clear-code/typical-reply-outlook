using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TypicalReply.Config;

namespace TypicalReply
{
    internal static class Const
    {
        internal static class RegistryPath
        {
            internal static readonly string DefaultPolicy = @"SOFTWARE\Policies\TypicalReply\Default";
            internal static readonly string Policy = @"SOFTWARE\Policies\TypicalReply";
        }

        internal static class Config
        {
            internal static readonly string FileName = @"TypicalReplyConfig.json";
        }

        internal static class Button
        {
            internal static readonly string TabMailGroupId = "TypicalReplyTabMailGroup";
            internal static readonly string TabReadMessageGroupId  = "TypicalReplyTabReadMessageGroup";
            internal static readonly string ContextMenuGalleryId  = "TypicalReplyContextMenuGallery";
        }
    }
}
