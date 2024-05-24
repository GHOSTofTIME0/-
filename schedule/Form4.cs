using System;
using System.Windows.Forms;

namespace schedule
{
    public partial class Form4 : Form
    {

        public string teacherFullName
        {
            get { return textBox1.Text; }
        }

        public string subjectName
        {
            get { return textBox2.Text; }
        }

        public Form4()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;

            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;

            }
        }
    }
}
