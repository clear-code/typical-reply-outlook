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
                Logger.Log("Start to setup custom UI");
                string ribbonTemplate = GetResourceText("TypicalReply.Ribbon.xml");
                var xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(ribbonTemplate);
                string namespaceURI = xmlDocument.ChildNodes[1].NamespaceURI;
                Global global = Global.GetInstance();
                XmlNode galleryInTabMailElem = xmlDocument.SelectSingleNode($"//*[@id='{Const.Button.TabMailGroupGalleryId}']");
                XmlNode galleryInTabReadMessageElem = xmlDocument.SelectSingleNode($"//*[@id='{Const.Button.TabReadMessageGroupGalleryId}']");
                XmlNode contextDropDownElem = xmlDocument.SelectSingleNode($"//*[@id='{Const.Button.ContextMenuGalleryId}']");

                var targetParams = new List<(XmlNode, string)>
                {
                    (galleryInTabMailElem, Const.Button.TabMailGroupGalleryId),
                    (galleryInTabReadMessageElem, Const.Button.TabReadMessageGroupGalleryId),
                    (contextDropDownElem, Const.Button.ContextMenuGalleryId),
                };

                (string cultureName, string lang) = GetCurrentUICultureInfo();
                foreach (var templateConfig in global.Config.ButtonConfigList)
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

        private MailItem CreateNewMail(ButtonConfig config, MailItem selectedMailItem)
        {
            MailItem itemToReply = null;

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

            if (config.AllowedDomainsType == AllowedDomainsType.UserSpecification)
            {
                var loweredAllowedDomains = config.LoweredAllowedDomains;
                foreach (Recipient recipient in itemToReply.Recipients)
                {
                    RecipientInfo recipientInfo = new RecipientInfo(recipient);
                    string targetDomain = recipientInfo.Domain.ToLowerInvariant();
                    if (loweredAllowedDomains.Any(_ => _ == targetDomain))
                    {
                        continue;
                    }
                    Marshal.ReleaseComObject(itemToReply);
                    return null;
                }
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
                        replyMessage += string.Join("\n", itemToReply.Body.TrimStart().Split('\n').Select(_ => $"> {_}"));
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

        private (string cultureName, string lang) GetCurrentUICultureInfo()
        {
            string currentUICultureName = CultureInfo.CurrentUICulture.Name;
            return (currentUICultureName, currentUICultureName.Split('-')[0]);
        }

        public void OnClickButton(IRibbonControl control)
        {
            string id = control.Id;
            ButtonConfig config = Global.
                GetInstance().
                Config.
                ButtonConfigList?.
                FirstOrDefault(_ =>
                    id == $"{_.Id}{Const.Button.TabMailGroupGalleryId}" ||
                    id == $"{_.Id}{Const.Button.ContextMenuGalleryId}" ||
                    id == $"{_.Id}{Const.Button.TabReadMessageGroupGalleryId}");
            if (config is null)
            {
                //TODO: Logging error;
                return;
            }

            MailItem selectedMailItem;

            if (control.Id == $"{config.Id}{Const.Button.TabReadMessageGroupGalleryId}")
            {
                selectedMailItem = GetActiveInspectorMailItem();
            }
            else
            {
                selectedMailItem = GetActiveExplorerMailItem();
            }

            MailItem newMailItem = CreateNewMail(config, selectedMailItem);

            newMailItem?.Display();
        }

        public string GetLabel(IRibbonControl control)
        {
            Config.Config typicalReplyConfig = Global.GetInstance().Config;
            return typicalReplyConfig.GalleryLabel;
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
