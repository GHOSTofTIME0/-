using System;
using System.Windows.Forms;

namespace schedule
{
    public partial class CreateFile : Form
    {
        public DateTime DateBegin { get; private set; }
        public DateTime DateEnd { get; private set; }
        public string EnteredText
        {
            get { return textBox1.Text; }
        }
        public CreateFile()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) //TODO если просто закрыть форму, то TabControl все равно отрисуется. ПОФИКСИТЬ
        {
            if (!string.IsNullOrWhiteSpace(textBox1.Text))
            {
                DateBegin = dateTimePicker1.Value.Date;
                DateEnd = dateTimePicker2.Value.Date;
                this.Close();
            }
            else MessageBox.Show("Вы не ввели название", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
