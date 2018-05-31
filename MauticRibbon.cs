using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MauticOutlookPlugin {
    public partial class MauticRibbon
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {
            

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e) {
            Globals.ThisAddIn.MarkMessageTrackable(Globals.ThisAddIn.Application.ActiveInspector().CurrentItem as Outlook.MailItem, toggleButton1.Checked);
        }
    }
}
