using System;
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
using System.Globalization;
using System.Threading;

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
                if (ribbonID != "Microsoft.Outlook.Mail.Read" && ribbonID != "Microsoft.Outlook.Explorer")
                {
                    return "";
                }
                Logger.Log("Start to setup custom UI");
                string ribbonTemplate = GetResourceText("TypicalReply.Ribbon.xml");
                var xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(ribbonTemplate);
                string namespaceURI = xmlDocument.ChildNodes[1].NamespaceURI;
                RuntimeParams global = RuntimeParams.GetInstance();

                XmlNode removeTargetNode = null;
                if (ribbonID == "Microsoft.Outlook.Mail.Read")
                {
                    // An error "TabMail does not found" can be raised if TabMail is specified 
                    // when target ribbon is not "Microsoft.Outlook.Mail.Read"
                    removeTargetNode = xmlDocument.SelectSingleNode("//*[@idMso='TabMail']");
                }
                else if (ribbonID == "Microsoft.Outlook.Explorer")
                {
                    // An error "TabReadMessage does not found" can be raised if TabReadMessage is specified
                    // when target ribbon is not "Microsoft.Outlook.Mail.Read"
                    removeTargetNode = xmlDocument.SelectSingleNode("//*[@idMso='TabReadMessage']");
                }
                removeTargetNode?.ParentNode.RemoveChild(removeTargetNode);
                XmlNode groupInTabMailElem = xmlDocument.SelectSingleNode($"//*[@id='{Const.Button.TabMailGroupId}']");
                XmlNode groupInTabReadMessageElem = xmlDocument.SelectSingleNode($"//*[@id='{Const.Button.TabReadMessageGroupId}']");
                XmlNode contextGalleryElem = xmlDocument.SelectSingleNode($"//*[@id='{Const.Button.ContextMenuGalleryId}']");

                var insertMsoTargetParams = new List<(XmlNode, string)>
                {
                    (groupInTabMailElem, global.Config.TabMailInsertAfterMso),
                    (groupInTabReadMessageElem, global.Config.TabReadInsertAfterMso),
                    (contextGalleryElem, global.Config.ContextMenuInsertAfterMso),
                };

                foreach (var (node, insertAfterMso) in insertMsoTargetParams)
                {
                    var nodeAsElement = node as XmlElement;
                    if (node == null)
                    {
                        continue;
                    }
                    if (!string.IsNullOrEmpty(insertAfterMso))
                    {
                        nodeAsElement.SetAttribute("insertAfterMso", insertAfterMso);
                    }
                }

                var targetParams = new List<(XmlNode, string)>
                {
                    (groupInTabMailElem, Const.Button.TabMailGroupId),
                    (groupInTabReadMessageElem, Const.Button.TabReadMessageGroupId),
                    (contextGalleryElem, Const.Button.ContextMenuGalleryId),
                };

                foreach (var buttonConfig in global.Config.ButtonConfigList)
                {
                    if (string.IsNullOrEmpty(buttonConfig.Id))
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(buttonConfig.Label))
                    {
                        continue;
                    }
                    foreach (var (node, suffix) in targetParams)
                    {
                        if (node == null)
                        {
                            continue;
                        }
                        XmlElement button = xmlDocument.CreateElement("button", namespaceURI);
                        button.SetAttribute("id", $"{buttonConfig.Id}{suffix}");
                        button.SetAttribute("label", buttonConfig.Label);
                        button.SetAttribute("onAction", nameof(OnClickButton));
                        if (suffix != Const.Button.ContextMenuGalleryId)
                        {
                            //We cannot specify button size in the context menu.
                            switch (buttonConfig.Size)
                            {
                                case ButtonSize.Large:
                                    button.SetAttribute("size", "large");
                                    break;
                                case ButtonSize.Normal:
                                    button.SetAttribute("size", "normal");
                                    break;
                            }
                        }
                        if (!string.IsNullOrEmpty(buttonConfig.Image))
                        {
                            button.SetAttribute("image", buttonConfig.Image);
                        }
                        //if (!string.IsNullOrEmpty(templateConfig.AccessKey))
                        //{
                        //    button.SetAttribute("keytip", templateConfig.AccessKey);
                        //}
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

        public virtual object Ribbon_LoadImage(string imageName)
        {
            switch (imageName)
            {
                case "logo.png":
                    return Properties.Resources.logo;
            }
            return null;
        }

        #endregion

        private MailItem GetActiveExplorerMailItem()
        {
            Explorer activeExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (activeExplorer.Selection.Count > 0 &&
                activeExplorer.Selection[1] is MailItem selObject)
            {
                return selObject;
            }
            return null;
        }

        private MailItem GetActiveInspectorMailItem()
        {
            return Globals.ThisAddIn.Application.ActiveInspector()?.CurrentItem as MailItem;
        }

        private MailItem CreateReplyMail(ButtonConfig config, MailItem selectedMailItem)
        {
            MailItem mailItemToReply = null;

            switch (config.RecipientsType)
            {
                case RecipientsType.All:
                    mailItemToReply = selectedMailItem.ReplyAll();
                    break;
                case RecipientsType.Sender:
                    mailItemToReply = selectedMailItem.Reply();
                    break;
                case RecipientsType.SpecifiedByUser:
                    mailItemToReply = selectedMailItem.Reply();
                    while (mailItemToReply.Recipients.Count > 0)
                    {
                        mailItemToReply.Recipients.Remove(1);
                    }
                    foreach (var recipient in config.Recipients)
                    {
                        mailItemToReply.Recipients.Add(recipient);
                    }
                    break;
                default:
                    mailItemToReply = selectedMailItem.Reply();
                    while (mailItemToReply.Recipients.Count > 0)
                    {
                        mailItemToReply.Recipients.Remove(1);
                    }
                    break;
            }

            if (config.AllowedDomainsType == AllowedDomainsType.SpecifiedByUser)
            {
                var loweredAllowedDomains = config.LoweredAllowedDomains;
                foreach (Recipient recipient in mailItemToReply.Recipients)
                {
                    RecipientInfo recipientInfo = new RecipientInfo(recipient);
                    string targetDomain = recipientInfo.Domain.ToLowerInvariant();
                    if (loweredAllowedDomains.Any(_ => _ == targetDomain))
                    {
                        continue;
                    }
                    Logger.Log($"Prohibited domain: {targetDomain}");
                    Marshal.ReleaseComObject(mailItemToReply);
                    return null;
                }
            }

            if (!string.IsNullOrEmpty(config.Subject))
            {
                mailItemToReply.Subject = config.Subject;
            }

            if (!string.IsNullOrEmpty(config.SubjectPrefix))
            {
                mailItemToReply.Subject = $"{config.SubjectPrefix} {mailItemToReply.Subject}";
            }

            string replyMessage = "";

            if (config.QuoteType && !string.IsNullOrEmpty(selectedMailItem.Body))
            {
                switch (mailItemToReply.BodyFormat)
                {
                    case OlBodyFormat.olFormatHTML:
                    case OlBodyFormat.olFormatRichText:
                        mailItemToReply.BodyFormat = OlBodyFormat.olFormatPlain;
                        replyMessage = "\n\n> -----Original Message-----\n";
                        replyMessage += string.Join("\n", mailItemToReply.Body.TrimStart().Split('\n').Select(_ => $"> {_}"));
                        break;
                    default:
                        replyMessage = mailItemToReply.Body;
                        break;
                }
            }

            mailItemToReply.Body = config.Body ?? "";
            if (!string.IsNullOrEmpty(replyMessage))
            {
                mailItemToReply.Body += replyMessage;
            }

            switch (config.ForwardType)
            {
                case ForwardType.Attachment:
                    mailItemToReply.Attachments.Add(selectedMailItem, OlAttachmentType.olEmbeddeditem);
                    break;
                case ForwardType.Inline:
                    //TODO: Support Inline
                    mailItemToReply.Attachments.Add(selectedMailItem, OlAttachmentType.olEmbeddeditem);
                    break;
            }

            return mailItemToReply;
        }

        public void OnClickButton(IRibbonControl control)
        {
            string id = control.Id;
            ButtonConfig config = RuntimeParams.
                GetInstance().
                Config.
                ButtonConfigList?.
                FirstOrDefault(_ =>
                    id == $"{_.Id}{Const.Button.TabMailGroupId}" ||
                    id == $"{_.Id}{Const.Button.ContextMenuGalleryId}" ||
                    id == $"{_.Id}{Const.Button.TabReadMessageGroupId}");
            if (config is null)
            {
                Logger.Log("Failed to get the target button config");
                return;
            }

            MailItem selectedMailItem;
            if (control.Id == $"{config.Id}{Const.Button.TabReadMessageGroupId}")
            {
                selectedMailItem = GetActiveInspectorMailItem();
            }
            else
            {
                selectedMailItem = GetActiveExplorerMailItem();
            }

            if (selectedMailItem is null)
            {
                Logger.Log("No MailItem found");
                return;
            }
            MailItem replyMail = CreateReplyMail(config, selectedMailItem);
            replyMail?.Display();
        }

        public string GetLabel(IRibbonControl control)
        {
            Config.Config typicalReplyConfig = RuntimeParams.GetInstance().Config;
            return typicalReplyConfig.GroupLabel;
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
