using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace EmailSendWarning
{
    public partial class EmailSendWarning
    {
        private void EmailSendWarning_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        private void EmailSendWarning_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
            try
            {

                Outlook.MailItem mail = Item as Outlook.MailItem;

                List<String> strRecipients = new List<String>();
                StringBuilder strStringBuilder = new StringBuilder();

                Microsoft.Office.Interop.Outlook.Recipients recipients = mail.Recipients;

                Boolean resolved;

                foreach (Microsoft.Office.Interop.Outlook.Recipient rec in recipients)
                {
                    resolved = rec.Resolve();
                    if (resolved)
                    {
                        strRecipients.Add("[" + rec.Address + "]");
                    }
                }

                strStringBuilder.Append("Your are about to send an email. Please confirm the below details of the email and press 'OK' to send the email and press cancel to stop sending." + Environment.NewLine);
                strStringBuilder.Append("--------------------------------------------------------------------" + Environment.NewLine);
                strStringBuilder.Append("--------------------------------------------------------------------" + Environment.NewLine);
                strStringBuilder.Append("Subject : " + mail.Subject + " " + Environment.NewLine);
                strStringBuilder.Append("--------------------------------------------------------------------" + Environment.NewLine);
                strStringBuilder.Append("Sender Email: " + mail.SendUsingAccount.DisplayName + " " + Environment.NewLine);
                strStringBuilder.Append("Sender Name: " + mail.SendUsingAccount.UserName + " " + Environment.NewLine);
                strStringBuilder.Append("--------------------------------------------------------------------" + Environment.NewLine);
                strStringBuilder.Append("Recipients : " + String.Join(", ", strRecipients.ToArray()).ToString() + " " + Environment.NewLine);
                strStringBuilder.Append("--------------------------------------------------------------------" + Environment.NewLine);


                System.Windows.Forms.DialogResult objDialogResult = System.Windows.Forms.MessageBox.Show(strStringBuilder.ToString(),
                    "Warning Message",
                    System.Windows.Forms.MessageBoxButtons.OKCancel, System.Windows.Forms.MessageBoxIcon.Warning
                    , System.Windows.Forms.MessageBoxDefaultButton.Button2, System.Windows.Forms.MessageBoxOptions.ServiceNotification);


                if (objDialogResult == System.Windows.Forms.DialogResult.Cancel)
                {
                    Cancel = true;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message,
                    "Warning Message",
                    System.Windows.Forms.MessageBoxButtons.OKCancel, System.Windows.Forms.MessageBoxIcon.Warning
                    , System.Windows.Forms.MessageBoxDefaultButton.Button1, System.Windows.Forms.MessageBoxOptions.RightAlign);

                throw ex;
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(EmailSendWarning_Startup);
            this.Shutdown += new System.EventHandler(EmailSendWarning_Shutdown);
        }

        #endregion
    }
}
