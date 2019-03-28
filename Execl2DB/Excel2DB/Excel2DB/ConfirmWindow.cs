using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel2DB
{
    public partial class ConfirmWindow : Form
    {

        private string dispStr;
        private List<string> sourceItems;
        private List<string> DataBaseItems;


        public ConfirmWindow(string dispStr, List<string> sourceItems, List<string> DataBaseItems)
        {
            InitializeComponent();
            this.dispStr = dispStr;
            this.sourceItems = sourceItems;
            this.DataBaseItems = DataBaseItems;
        }

        

        private void ConfirmWindow_Load(object sender, EventArgs e)
        {
            textBox1.Text = dispStr;

            textBox2.Text = "数据对应关系如下：\r\n\r\n";
            for (int i = 0; i < sourceItems.Count; i++)
            {
                textBox2.Text += sourceItems[i] + "  ————>  " + DataBaseItems[i] + "\r\n";
            }
            textBox2.Text += "\r\n你可以在界面中的数据表格中修改对应关系";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Yes;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.No;
        }

       
    }
}
