using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace Srplug
{
    public partial class ThisAddIn
    {
        public static List<String> srd_email;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            srd_email = new List<String>();
            ListSharedMailboxes();
           

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }
        private void ListSharedMailboxes()
        {
            try
            {
                Outlook.Application outlookApp = this.Application;
                Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
                
                // Get the default account
                Outlook.Account defaultAccount = GetDefaultAccount(outlookNamespace);
                if (defaultAccount != null)
                {
                    string sharedMailboxesInfo = "Shared Mailboxes:\n";
                    foreach (Outlook.Folder folder in outlookNamespace.Folders)
                    {
                        // Check if the folder is a shared mailbox
                        if (folder.Store.DisplayName != defaultAccount.DisplayName && folder.Store.DisplayName.Contains("@"))
                        {
                           // sharedMailboxesInfo += $"{folder.Name} - {folder.Store.DisplayName}\n";
                           foreach(Account or in outlookNamespace.Accounts)
                            {
                                if(folder.Store.DisplayName.ToLower().Contains(or.DisplayName.ToLower()))
                                {
                                    srd_email.Add(or.DisplayName);
                                }
                            }
                           // string st_te = folder.Store.DisplayName.ToString();
                           // srd_email.Add(folder.Store.DisplayName);
                        }
                    }
                 //   MessageBox.Show(sharedMailboxesInfo, "Shared Mailboxes");
                }
                else
                {
                    MessageBox.Show("No default account found.", "Error");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private Outlook.Account GetDefaultAccount(Outlook.NameSpace outlookNamespace)
        {
            foreach (Outlook.Account account in outlookNamespace.Accounts)
            {
                if (account.AccountType == Outlook.OlAccountType.olExchange)
                {
                    return account; // Return the first Exchange account found
                }
            }
            return null; // No default account found
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        /// 
        public static Account GetAccountByEmail(string emailAddress)
        {
            Application outlookApp = new Application();
            foreach (Account account in outlookApp.Session.Accounts)
            {
                if (account.SmtpAddress.Equals(emailAddress, StringComparison.OrdinalIgnoreCase))
                {
                    return account;
                }
            }
            return null; // Return null if no account matches
        }
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
