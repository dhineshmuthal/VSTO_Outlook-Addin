using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace Srplug
{
    public partial class Ribbon1
    {
        Outlook.Inspectors inspectors;
        // public static List<RibbonDropDownItem> comboBoxItems;
        //Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {           
            foreach(string temps in Srplug.ThisAddIn.srd_email.Distinct())
            {
                RibbonDropDownItem ribbonDropDownItemImpl2 = Factory.CreateRibbonDropDownItem();
                ribbonDropDownItemImpl2.Label = temps;
                comboBox1.Items.Add(ribbonDropDownItemImpl2);
                
            }
           // Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
           // this.comboBox1.Items.Add(comboBoxItems);
        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            String cmbsel_sharedmail = comboBox1.Text;
          //  MessageBox.Show(comboBox1.Text);
            var item = e.Control.Context as Inspector;
            var mailItem = item.CurrentItem as MailItem;
           

           // Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "Ticket from Shared Mail box";
                   // mailItem.Body = "";
                    mailItem.SendUsingAccount = ThisAddIn.GetAccountByEmail(cmbsel_sharedmail);
                }
             //   mailItem.SendUsingAccount ="dhinesh@muthal365.com";

            }

        }
        
    }
   

    
}
