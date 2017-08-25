using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Threading.Tasks;
using System.Threading;

namespace Duplikate_Entferner
{
    public partial class Duplikate
    {
        private void Duplikate_Load(object sender, RibbonUIEventArgs e)
        {
            toggleButton1.Checked = Properties.Settings.Default.auto_delete;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            button1.Enabled = false;
            this.label1.Label = "Verarbeite ...";

            Task t = Task.Factory.StartNew(() => {
                return Globals.ThisAddIn.removeDuplikates();
            }).ContinueWith(r => {
                this.label1.Label = "Erfolgreich " + r.Result +" Elemente gelöscht.";
                button1.Enabled = true;
            });
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.auto_delete = toggleButton1.Checked;
        }
    }
}
