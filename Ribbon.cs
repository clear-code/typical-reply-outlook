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
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            string ribbonTemplate = GetResourceText("TypicalReply.Ribbon.xml");
            var xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(ribbonTemplate);
            string namespaceURI = xmlDocument.ChildNodes[1].NamespaceURI;
            var global = Global.GetInstance();
            var ribbonDropDownElem = xmlDocument.SelectSingleNode("//*[@id='RibbonDropDownTypicalReply']");
            var contextDropDownElem = xmlDocument.SelectSingleNode("//*[@id='ContextMenuTypicalReply']");

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
                var ribbonDropDownButton = xmlDocument.CreateElement("button", namespaceURI);
                ribbonDropDownButton.SetAttribute("id", $"{templateConfig.Id}RibbonDropdown");
                ribbonDropDownButton.SetAttribute("label", templateConfig.Label);
                ribbonDropDownButton.SetAttribute("onAction", "OnCreateTemplate");
                if (!string.IsNullOrEmpty(templateConfig.AccessKey))
                {
                    ribbonDropDownButton.SetAttribute("keytip", templateConfig.AccessKey);
                }
                ribbonDropDownElem.AppendChild(ribbonDropDownButton);

                var contextDropDownButton = xmlDocument.CreateElement("button", namespaceURI);
                contextDropDownButton.SetAttribute("id", $"{templateConfig.Id}ContextMenu");
                contextDropDownButton.SetAttribute("label", templateConfig.Label);
                contextDropDownButton.SetAttribute("onAction", "OnCreateTemplate");
                if (!string.IsNullOrEmpty(templateConfig.AccessKey))
                {
                    contextDropDownButton.SetAttribute("keytip", templateConfig.AccessKey);
                }
                contextDropDownElem.AppendChild(contextDropDownButton);
            }
            return xmlDocument.InnerXml;
        }

        #endregion

        #region リボンのコールバック
        //ここでコールバック メソッドを作成します。コールバック メソッドの追加について詳しくは https://go.microsoft.com/fwlink/?LinkID=271226 をご覧ください

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        private Outlook.MailItem GetActiveExplorerMailItem()
        {
            Outlook.Explorer activeExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
            //TODO: Accept multiple selection
            if (activeExplorer.Selection.Count > 0 &&
                activeExplorer.Selection[1] is Outlook.MailItem selObject)
            {
                return selObject;
            }
            return null;
        }

        private Outlook.MailItem GetActiveImspectorMailItem()
        {
            return Globals.ThisAddIn.Application.ActiveInspector()?.CurrentItem as Outlook.MailItem;
        }

        private Outlook.MailItem GetMailItem()
        {
            return GetActiveExplorerMailItem() ?? GetActiveImspectorMailItem();
        }

        private Outlook.MailItem CreateNewMail(TemplateConfig config, Outlook.MailItem selectedMailItem)
        {
            Outlook.MailItem newMailItem = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
            var subjectPrefix = config.SubjectPrefix ?? "";
            var subject = config.Subject ?? "";
            newMailItem.Subject = $"{subjectPrefix} {subject}";
            newMailItem.Body = config.Body ?? "";

            switch (config.ForwardType)
            {
                case ForwardType.Attachment:
                    newMailItem.Attachments.Add(selectedMailItem, Outlook.OlAttachmentType.olEmbeddeditem);
                    break;
                default:
                    newMailItem.Body += "\n";
                    newMailItem.Body += string.Join("\n", selectedMailItem.Body.Split('\n').Select(_ => " > ${_}"));
                    break;
            }

            switch (config.RecipientsType)
            {
                case RecipientsType.All:
                    newMailItem.Recipients.Add(selectedMailItem.Sender.Address);
                    newMailItem.CC = selectedMailItem.CC;
                    foreach (Recipient originalRecipient in selectedMailItem.Recipients)
                    {
                        if (newMailItem.Sender.Address != originalRecipient.Address)
                        {
                            newMailItem.CC += $" {selectedMailItem.Sender.Address}";

                        }
                    }
                    break;
                case RecipientsType.Sender:
                    newMailItem.Recipients.Add(selectedMailItem.Sender.Address);
                    break;
                case RecipientsType.UserSpecification:
                    foreach (var recipient in config.Recipients.Split())
                    {
                        newMailItem.Recipients.Add(recipient);
                    }
                    break;
                default:
                    break;

            }

            return newMailItem;
        }

        public void OnCreateTemplate(Office.IRibbonControl control)
        {
            TypicalReplyConfig typicalReplyConfig = Global.GetInstance().Config;
            var config = typicalReplyConfig.TemplateConfigList.FirstOrDefault(_ => control.Id.Contains(_.Id));
            if (config == null)
            {
                //TODO: Logging error;
                return;
            }
            Outlook.MailItem selectedMailItem = GetMailItem();
            Outlook.MailItem newMailItem = CreateNewMail(config, selectedMailItem);
            newMailItem.Display();
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
