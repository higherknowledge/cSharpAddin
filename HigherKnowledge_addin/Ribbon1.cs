using HigherKnowledge_addin.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace HigherKnowledge_addin
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        string response = null;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("HigherKnowledge_addin.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void onViewButton(Office.IRibbonControl control)
        {
            
            if (response == null)
                fetch();
            string[] val = response.Split('|');
            string dis = "Subject : \n\n" + val[0] + "\n\nCC :\n\n" + val[1] + "\n\nBody :\n\n" + val[2];
            MessageBox.Show(dis);
        }

        public void OnReplyButton(Office.IRibbonControl control)
        {
            var context = control.Context;
            if(context is Outlook.Explorer)
            {
                var explorer = context as Outlook.Explorer;
                var selections = explorer.Selection;
                foreach(var child in selections)
                {
                    if(child is Outlook.MailItem)
                    {
                        var mail = child as Outlook.MailItem;
                        showDialog(mail);
                        break;
                    }
                }
            }

            else if(context is Outlook.Inspector)
            {
                var ins = context as Outlook.Inspector;
                if (ins.CurrentItem is Outlook.MailItem)
                {
                    var mail = ins.CurrentItem as Outlook.MailItem;
                    showDialog(mail);
                    ins.Close(Outlook.OlInspectorClose.olSave);
                }

                else
                    MessageBox.Show("Cannot perform the action in the current context");
            }
            else
            {
                MessageBox.Show("Cannot perform the action in the current context");
            }
        }

        private void showDialog(Outlook.MailItem mail)
        {
            string name = mail.Sender.Address;
            DialogResult result = MessageBox.Show("Send HK response to " + name,"Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                MessageBox.Show(ThisAddIn.User);
                Outlook.MailItem reply = mail.Reply();

                 if (response == null)
                {
                    fetch();
                }

                string[] val = response.Split('|');
                reply.Subject = val[0];
                reply.CC = val[1];
                reply.Body = val[2];
                reply.Send(); 
            }
        }
        
        #endregion

        #region Helpers

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

        public Bitmap getDone(Office.IRibbonControl control)
        {
            return Resources.done;
        }

        public Bitmap getView(Office.IRibbonControl control)
        {
            return Resources.View;
        }

        private void fetch()
        {
            string raw = "https://raw.githubusercontent.com/";
            string path = "abhivijay96/Templates/master/";

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(raw + path + ThisAddIn.User);
                //please comment the above line and uncomment the below line to test this
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(raw + path + "hkaddin@gmail.com");
                var res = (HttpWebResponse)request.GetResponse();
                var stream = res.GetResponseStream();
                StreamReader reader = new StreamReader(stream);
                response = reader.ReadToEnd();
                reader.Close();
                stream.Close();
            }

            catch (Exception e)
            {
                MessageBox.Show("Unable to fetch the template");
            }
        } 
        #endregion
    }
}
