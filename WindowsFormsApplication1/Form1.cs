using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void button1_Click_1(object sender, EventArgs e)
        {
            richTextBox1.Text = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                var validator = new OpenXmlValidator();
                int count = 0;
                
                foreach (ValidationErrorInfo error in validator.Validate(WordprocessingDocument.Open(openFileDialog1.FileName, true)))
                {
                    if (count == 0)
                    {
                        richTextBox1.Text = "";
                    }
                    
                    richTextBox1.Text += "\r\n";
                    count++;
                    richTextBox1.Text += ("Error Count : " + count) + "\r\n";
                    richTextBox1.Text += ("Description : " + error.Description) + "\r\n";
                    richTextBox1.Text += ("Path: " + error.Path.XPath) + "\r\n";
                    richTextBox1.Text += ("Part: " + error.Part.Uri) + "\r\n";
                    richTextBox1.Text += ("Node: " + error.Node) + "\r\n";
                    richTextBox1.Text += ("Error Type: " + error.ErrorType) + "\r\n";
                }
                lblError.Text += "end of file";
            }
            catch (Exception ex)
            {
                richTextBox1.Text += (ex.Message);
            }
        }

       

        
    }
}
