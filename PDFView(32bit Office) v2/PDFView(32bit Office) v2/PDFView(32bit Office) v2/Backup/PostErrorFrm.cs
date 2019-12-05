using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddIn4
{
    public partial class PostErrorFrm : Form
    {
        public PostErrorFrm()
        {
            InitializeComponent();
        }
        public PostErrorFrm(string str)
        {
            InitializeComponent();
            richTextBox1.Text = str;
        }
    }
}
