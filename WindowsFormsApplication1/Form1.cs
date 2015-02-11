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
            lblError.Text = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                var validator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in validator.Validate(WordprocessingDocument.Open(openFileDialog1.FileName, true)))
                {
                    lblError.Text += "\r\n";
                    count++;
                    lblError.Text += ("Error Count : " + count) + "\r\n";
                    lblError.Text += ("Description : " + error.Description) + "\r\n";
                    lblError.Text += ("Path: " + error.Path.XPath) + "\r\n";
                    lblError.Text += ("Part: " + error.Part.Uri) + "\r\n";
                }
                lblError.Text += "end of file";
            }
            catch (Exception ex)
            {
                lblError.Text += (ex.Message);
            }
        }

       

        
    }
}
