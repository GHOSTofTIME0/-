using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace schedule
{
    public partial class infoSubjectsForm : Form
    {

        Form1 form1Instance = new Form1();
        public infoSubjectsForm()
        {
            InitializeComponent();
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            TabControl tabControl = new TabControl();
            this.SuspendLayout();

            // Initialize TabControl
            tabControl.Dock = DockStyle.Fill;
            this.Controls.Add(tabControl);
            string tableName;
            // Add 11 pages to the TabControl
            for (int i = 1; i <= 11; i++)
            {
                TabPage tabPage = new TabPage($"Класс {i}");
                tabControl.TabPages.Add(tabPage);

                if(i > 4 && i <10)
                {
                    tableName = "Disciplines5_9";
                }
                else if (i > 9 && i < 12)
                {
                    tableName = "Disciplines10_11";
                }
                else
                {
                    tableName = "Disciplines1_4";
                }
               
                Dictionary<int, Dictionary<string, int>> headers = form1Instance.GetDisciplinesHeaders(form1Instance.connectString, tableName, i);

                // Add Labels to the tabPage
                AddSubjectLabels(tabPage, headers[i]);
            }

            // Form2
            this.AutoScaleDimensions = new SizeF(8F, 16F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(941, 785);
            this.Name = "Form2";
            this.Text = "Subjects and Difficulties";
            this.ResumeLayout(false);
        }

        private void AddSubjectLabels(TabPage tabPage, Dictionary<string, int> subjects)
        {
            int labelX = 10;
            int labelY = 10;
            int spaceBetweenLabels = 30;

            foreach (var subject in subjects)
            {
                Label subjectLabel = new Label();
                subjectLabel.Text = subject.Key;
                subjectLabel.Location = new Point(labelX, labelY);
                subjectLabel.Font = new Font("TimesNewRoman", 10);
                tabPage.Controls.Add(subjectLabel);

                Label difficultyLabel = new Label();
                difficultyLabel.Text = subject.Value.ToString();
                difficultyLabel.Location = new Point(labelX + 200, labelY); // Adjust X position for difficulty label
                tabPage.Controls.Add(difficultyLabel);

                labelY += spaceBetweenLabels; // Move to the next line for the next pair of labels
            }
        }
    }
}
