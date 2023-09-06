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
            var groupElement = xmlDocument.SelectSingleNode("//*[@id='DropDownTypicalReply']");
            var button = xmlDocument.CreateElement("button", namespaceURI);
            //button.SetAttribute("onAction", "OnCreateTemplate");
            button.SetAttribute("id", "ButtonTypicalReply");
            button.SetAttribute("label", "CreateTemplate");
            button.SetAttribute("onAction", "OnCreateTemplate");
            groupElement.AppendChild(button);
            button = xmlDocument.CreateElement("button", namespaceURI);
            button.SetAttribute("id", "ButtonTypicalReply2");
            button.SetAttribute("label", "CreateTemplate2");
            button.SetAttribute("onAction", "OnCreateTemplate");
            groupElement.AppendChild(button);
            return xmlDocument.InnerXml;
            //return GetResourceText("TypicalReply.Ribbon.xml");
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

        private Outlook.MailItem CreateNewMail(List<Outlook.MailItem> mailsForAttacihment)
        {
            Outlook.MailItem newMailItem = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
            foreach (Outlook.MailItem mailItem in mailsForAttacihment)
            {
                newMailItem.Attachments.Add(mailItem, Outlook.OlAttachmentType.olEmbeddeditem);
            }
            newMailItem.Subject = "new message!";
            return newMailItem;
        }

        public void OnCreateTemplate(Office.IRibbonControl control)
        {
            Outlook.MailItem selectedMailItem = GetMailItem();
            Outlook.MailItem newMailItem = CreateNewMail(new List<Outlook.MailItem>() { selectedMailItem });
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
