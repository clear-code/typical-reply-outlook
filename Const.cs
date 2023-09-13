﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TypicalReply
{
    internal static class Const
    {
        internal static class RegistryPath
        {
            internal static readonly string DefaultPolicy = @"SOFTWARE\Policies\TypicalReply\Default";
            internal static readonly string Policy = @"SOFTWARE\Policies\TypicalReply";
        }

        internal static class Button
        {
            internal static string TabMailGroupGalleryId { get; } = "TypicalReplyTabMailGroupGallery";
            internal static string TabReadMessageGroupGalleryId { get; } = "TypicalReplyTabReadMessageGroupGallery";
            internal static string ContextMenuGalleryId { get; } = "TypicalReplyContextMenuGallery";
        }
    }
}
