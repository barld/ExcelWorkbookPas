using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorkbook1
{
    public partial class Blad1
    {
        List<leerlingen> leerlingMeldingen = new List<leerlingen>();
        private ListObject listObject;
        private void Blad1_Startup(object sender, System.EventArgs e)
        {
            this.listObject = this.Controls.AddListObject(this.Range["B2"], "Products");

            leerlingMeldingen = new List<leerlingen>();
            leerlingMeldingen.Add(new leerlingen{code="042040E1BC1D80♪"});
            leerlingMeldingen.Add(new leerlingen { code = "042040E1BC1D80♪" });
            
            
            this.listObject.DataSource = leerlingMeldingen;
            this.listObject.AutoSetDataBoundColumnHeaders = true;
            //list1.

        }

        private void Blad1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.Startup += new System.EventHandler(this.Blad1_Startup);
            this.Shutdown += new System.EventHandler(this.Blad1_Shutdown);

        }

        #endregion

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string t = textBox1.Text.ToString();
            if(t.Length != 0 && textBox1.Text.Last() == '♪')
            {
                this.leerlingMeldingen.Add(new leerlingen { code = textBox1.Text.ToString() });
                this.listObject.DataSource = null;
                this.listObject.DataSource = this.leerlingMeldingen;
                textBox1.Clear();
            }
        }

    }
}
