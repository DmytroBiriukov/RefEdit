using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Globalization;
using Microsoft.Office.Tools.Ribbon;
//using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;

namespace RefEditor
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

           // maskedTextBox1.Mask="W{3}";

            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Range rng = Globals.ThisAddIn.Application.Selection.Range;
            //rng=doc.Range(0,0);

            Microsoft.Office.Interop.Word.Hyperlinks myLinks = doc.Hyperlinks;
            string test_file_Path = "#";
            object linkAddr = test_file_Path;
            string test_bookmark;
            test_bookmark=maskedTextBox1.Text.ToString();
            
            object linkSubAddr = test_bookmark;
            string screenTip;

            if((screenTip=textBox1.Text.ToString()).Length == 0)
            {      screenTip = "Author Title // Journal Volume Pages";
            }
            object linkScreenTip = screenTip;
            string test_todisplay = test_file_Path+test_bookmark;
            object linkToDisplay = test_todisplay;
            Microsoft.Office.Interop.Word.Hyperlink myLink = myLinks.Add(rng, ref linkAddr, ref linkSubAddr, ref linkScreenTip, ref linkToDisplay);

            this.Close();
        }
    }
}
