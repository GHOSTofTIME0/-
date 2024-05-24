using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace schedule
{
    public partial class Form3 : Form
    {

        private ListBox listBox;
        private Button btnDelete;
        private Button btnClose;


        public Form3()
        {
            InitializeComponent();
            InitializeComponents();
            LoadSubjects();
        }

        private void InitializeComponents()
        {
            listBox = new ListBox();
            btnDelete = new Button();
            btnClose = new Button();

            listBox.Dock = DockStyle.Fill;
            btnDelete.Text = "Удалить";
            btnClose.Text = "Закрыть";

            btnDelete.Enabled = false;

            listBox.SelectedIndexChanged += (sender, e) =>
            {
                btnDelete.Enabled = (listBox.SelectedIndex != -1);
            };

            btnDelete.Click += BtnDelete_Click;
            btnClose.Click += (sender, e) =>
            {
                Close();
            };

            Controls.Add(listBox);
            Controls.Add(btnDelete); Controls.Add(btnClose);

            listBox.Height = 150;
            btnDelete.Dock = DockStyle.Bottom;
            btnClose.Dock = DockStyle.Bottom;
        }

        private void LoadSubjects()
        {
            string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
            string templatePath = Path.Combine(projectPath, "source", "предметы.txt");
            listBox.Items.Clear();

            if (File.Exists(templatePath))
            {
                string[] lines = File.ReadAllLines(templatePath);
                listBox.Items.AddRange(lines);
            }
        }

        private void RemoveSubject(string subject)
        {
            string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
            string templatePath = Path.Combine(projectPath, "source", "предметы.txt");
            if (File.Exists(templatePath))
            {
                List<string> lines = new List<string>(File.ReadAllLines(templatePath));
                lines.Remove(subject);
                File.WriteAllLines(templatePath, lines);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (listBox.SelectedIndex != -1)
            {
                string selectedSubject = listBox.SelectedItem.ToString();
                RemoveSubject(selectedSubject);
                LoadSubjects();
            }
        }
    }
}
