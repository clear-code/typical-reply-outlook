﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Xml;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using TypicalReply.Config;
using Microsoft.Office.Core;
using System.Xml.Linq;

// TODO:  リボン (XML) アイテムを有効にするには、次の手順に従います:

// 1: 次のコード ブロックを ThisAddin、ThisWorkbook、ThisDocument のいずれかのクラスにコピーします。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. ボタンのクリックなど、ユーザーの操作を処理するためのコールバック メソッドを、このクラスの
//    "リボンのコールバック" 領域に作成します。メモ: このリボンがリボン デザイナーからエクスポートされたものである場合は、
//    イベント ハンドラー内のコードをコールバック メソッドに移動し、リボン拡張機能 (RibbonX) のプログラミング モデルで
//    動作するように、コードを変更します。

// 3. リボン XML ファイルのコントロール タグに、コードで適切なコールバック メソッドを識別するための属性を割り当てます。  

// 詳細については、Visual Studio Tools for Office ヘルプにあるリボン XML のドキュメントを参照してください。


namespace TypicalReply
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                Logger.Log("Start to setup custom UI");
                string ribbonTemplate = GetResourceText("TypicalReply.Ribbon.xml");
                var xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(ribbonTemplate);
                string namespaceURI = xmlDocument.ChildNodes[1].NamespaceURI;
                Global global = Global.GetInstance();
                XmlNode galleryInTabMailElem = xmlDocument.SelectSingleNode($"//*[@id='{Global.TabMailGroupGalleryId}']");
                XmlNode galleryInTabReadMessageElem = xmlDocument.SelectSingleNode($"//*[@id='{Global.TabReadMessageGroupGalleryId}']");
                XmlNode contextDropDownElem = xmlDocument.SelectSingleNode($"//*[@id='{Global.MenuInContextMenuId}']");

                var targetParams = new List<(XmlNode, string)>
                {
                    (galleryInTabMailElem, Global.TabMailGroupGalleryId),
                    (galleryInTabReadMessageElem, Global.TabReadMessageGroupGalleryId),
                    (contextDropDownElem, Global.MenuInContextMenuId),
                };

                foreach (var templateConfig in global.Config.TemplateConfigList)
                {
                    if (string.IsNullOrEmpty(templateConfig.Id))
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(templateConfig.Label))
                    {
                        continue;
                    }
                    foreach (var (node, postfix) in targetParams)
                    {
                        XmlElement button = xmlDocument.CreateElement("button", namespaceURI);
                        button.SetAttribute("id", $"{templateConfig.Id}{postfix}");
                        button.SetAttribute("label", templateConfig.Label);
                        button.SetAttribute("onAction", nameof(OnClickButton));
                        if (!string.IsNullOrEmpty(templateConfig.AccessKey))
                        {
                            button.SetAttribute("keytip", templateConfig.AccessKey);
                        }
                        node.AppendChild(button);
                    }
                }
                Logger.Log("Finish to setup custom UI");
                return xmlDocument.InnerXml;
            }
            catch (System.Exception ex)
            {
                Logger.Log(ex);
                throw;
            }
        }

        #endregion

        #region リボンのコールバック
        //ここでコールバック メソッドを作成します。コールバック メソッドの追加について詳しくは https://go.microsoft.com/fwlink/?LinkID=271226 をご覧ください

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        private MailItem GetActiveExplorerMailItem()
        {
            Explorer activeExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            //TODO: Accept multiple selection
            if (activeExplorer.Selection.Count > 0 &&
                activeExplorer.Selection[1] is MailItem selObject)
            {
                return selObject;
            }
            return null;
        }

        private MailItem GetActiveImspectorMailItem()
        {
            return Globals.ThisAddIn.Application.ActiveInspector()?.CurrentItem as MailItem;
        }

        private Outlook.MailItem GetMailItem()
        {
            return GetActiveExplorerMailItem() ?? GetActiveImspectorMailItem();
        }

        private MailItem CreateNewMail(TemplateConfig config, MailItem selectedMailItem)
        {
            MailItem itemToReply;

            switch (config.RecipientsType)
            {
                case RecipientsType.All:
                    itemToReply = selectedMailItem.ReplyAll();
                    break;
                case RecipientsType.Sender:
                    itemToReply = selectedMailItem.Reply();
                    break;
                case RecipientsType.UserSpecification:
                    itemToReply = selectedMailItem.Reply();
                    while (itemToReply.Recipients.Count > 0)
                    {
                        itemToReply.Recipients.Remove(1);
                    }
                    foreach (var recipient in config.Recipients)
                    {
                        itemToReply.Recipients.Add(recipient);
                    }
                    break;
                default:
                    itemToReply = selectedMailItem.Reply();
                    while (itemToReply.Recipients.Count > 0)
                    {
                        itemToReply.Recipients.Remove(1);
                    }
                    break;
            }

            foreach(Recipient recipient in itemToReply.Recipients)
            {
                var address = recipient.Address;
                var name = recipient.Name;
            }

            if (!string.IsNullOrEmpty(config.Subject))
            {
                itemToReply.Subject = config.Subject;
            }

            if (!string.IsNullOrEmpty(config.SubjectPrefix))
            {
                itemToReply.Subject = $"{config.SubjectPrefix} {itemToReply.Subject}";
            }

            string replyMessage = "";

            if (config.QuoteType && !string.IsNullOrEmpty(selectedMailItem.Body))
            {
                switch (itemToReply.BodyFormat)
                {
                    case OlBodyFormat.olFormatHTML:
                    case OlBodyFormat.olFormatRichText:
                        itemToReply.BodyFormat = OlBodyFormat.olFormatPlain;
                        replyMessage = "\n\n> -----Original Message-----\n";
                        replyMessage += string.Join("\n", itemToReply.Body.Split('\n').Select(_ => $"> {_}"));
                        break;
                    default:
                        replyMessage = itemToReply.Body;
                        break;
                }
            }

            itemToReply.Body = config.Body ?? "";
            if (!string.IsNullOrEmpty(replyMessage))
            {
                itemToReply.Body += replyMessage;
            }

            switch (config.ForwardType)
            {
                case ForwardType.Attachment:
                    itemToReply.Attachments.Add(selectedMailItem, OlAttachmentType.olEmbeddeditem);
                    break;
                case ForwardType.Inline:
                    //TODO: Support Inline
                    itemToReply.Attachments.Add(selectedMailItem, OlAttachmentType.olEmbeddeditem);
                    break;
            }

            return itemToReply;
        }

        public void OnClickButton(IRibbonControl control)
        {
            TypicalReplyConfig typicalReplyConfig = Global.GetInstance().Config;
            var config = typicalReplyConfig
                .TemplateConfigList
                .FirstOrDefault(_ =>
                    control.Id == $"{_.Id}{Global.TabMailGroupGalleryId}" || 
                    control.Id == $"{_.Id}{Global.MenuInContextMenuId}"   ||
                    control.Id == $"{_.Id}{Global.TabReadMessageGroupGalleryId}");

            if (config == null)
            {
                //TODO: Logging error;
                return;
            }
            MailItem selectedMailItem = GetMailItem();
            MailItem newMailItem = CreateNewMail(config, selectedMailItem);
            newMailItem.Display();
        }

        public string GetLabel(IRibbonControl control)
        {
            TypicalReplyConfig typicalReplyConfig = Global.GetInstance().Config;
            return typicalReplyConfig.RibbonLabel;
        }

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
