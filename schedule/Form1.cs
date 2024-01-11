using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Humanizer;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace schedule
{
    public partial class Form1 : Form
    {
        private TabControl tab;
        private DataGridView dataGridView1;
        private string[] daysOfWeek = { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота" };
        private string outputPath;
        public Form1()
        {
            InitializeComponent();
        }




        //СОЗДАНИЕ 
        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void создатьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            CreateFile createFileForm = new CreateFile();
            createFileForm.FormClosed += CreateFileForm_FormClosed;

            createFileForm.Show();
        }

        private void CreateFileForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            CreateFile createFileForm = (CreateFile)sender;
            DateTime dateTimePickerValue1 = createFileForm.DateBegin;
            DateTime dateTimePickerValue2 = createFileForm.DateEnd;
            string enteredText = createFileForm.EnteredText;
            this.Text = enteredText;

            ClearTabs();
            InitializeTab();
        }



        private void InitializeTab()
        {
            tab = new TabControl();

            tab.Size = new Size((System.Drawing.Point)this.Size);


            for (int i = 0; i < 11; i++)
            {
                TabPage tp = new TabPage(Convert.ToString(i + 1));
                DataGridView Grid = CreateDataGridView();
                tp.Controls.Add(Grid);
                tab.TabPages.Add(tp);

            }

            this.Controls.Add(tab);

            // Убедитесь, что Controls.Add(tab) вызывается после добавления вкладок
            Controls.Add(tab);
            tab.Location = new System.Drawing.Point(0, menuStrip1.Bottom);
        }

        private void ClearTabs()
        {
            if (tab != null)
            {
                tab.TabPages.Clear();

                this.Controls.Remove(tab);

                tab.Dispose();
            }
        }

        private DataGridView CreateDataGridView()
        {
            DataGridView dataGridView = new DataGridView();
            dataGridView.Dock = DockStyle.Fill; // Используйте DockStyle.Fill

            DataGridViewTextBoxColumn rowNumerColumn = new DataGridViewTextBoxColumn();
            rowNumerColumn.HeaderText = "#";
            dataGridView.Columns.Add(rowNumerColumn);

            string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
            string templatePath = Path.Combine(projectPath, "source", "предметы.txt");

            for (int i = 0; i < 12; i++)
            {

                if (i % 2 == 0)
                {
                    DataGridViewComboBoxColumn subjectColumn = new DataGridViewComboBoxColumn();
                    subjectColumn.HeaderText = daysOfWeek[i / 2];
                    //subjectColumn.Items.AddRange(GetSubjectList().ToArray());
                    dataGridView.Columns.Add(subjectColumn);

                    if(File.Exists(templatePath))
                    {
                        string[] lines = File.ReadAllLines(templatePath);

                        foreach(string line in lines)
                        {
                            subjectColumn.Items.Add(line);
                        }
                    }
                }
                else
                {
                    DataGridViewTextBoxColumn roomColumn = new DataGridViewTextBoxColumn();
                    roomColumn.HeaderText = "кабинет";
                    dataGridView.Columns.Add(roomColumn);
                }
            }

            dataGridView.CellClick += dataGridView_CellClick;

            dataGridView.RowsAdded += dataGridView_RowsAdded;

            dataGridView.RowsRemoved += dataGridView_RowsRemoved;

            dataGridView.EditingControlShowing += dataGridView_EditingControlShowing;

            dataGridView.CellEndEdit += dataGridView_CellEndEdit;

            dataGridView1 = dataGridView;



            return dataGridView;
        }

        //////////////////////////////////////добавление и удаление предметов//////////////////////////////////////////////////////////////////

        private void удалитьПредметToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using(var subjectsListForm = new Form3())
            {
                subjectsListForm.ShowDialog();
            }
            UpdateComboBoxItems();
        }

        private void добавитьПредметToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
            string templatePath = Path.Combine(projectPath, "source", "предметы.txt");
            using (var inputForm = new Form2())
            {
                if(inputForm.ShowDialog() == DialogResult.OK)
                {
                    string newText = inputForm.nameText;

                    File.AppendAllText(templatePath, newText + "\r\n");
                }
            }
            UpdateComboBoxItems();
        }
        




        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////////////////////////////////////////////////////////////////////////////////////

        private void dataGridView_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridView dataGridView = (DataGridView)sender;

            for (int i = 0; i < e.RowCount; i++)
            {
                dataGridView.Rows[0].Cells[0].Value = "1";
                dataGridView.Rows[e.RowIndex + i].Cells[0].Value = (e.RowIndex + i + 1).ToString();
            }
        }

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 0 && e.RowIndex >= 0)
            {
                DataGridView dataGridView = (DataGridView)sender;

                // Проверяем, что текущая ячейка является ячейкой типа DataGridViewComboBoxCell
                if (dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewComboBoxCell comboBoxCell)
                {
                    // Проверяем, что Items в комбо-боксе пусты, и, если да, добавляем элементы
                    if (comboBoxCell.Items.Count == 0)
                    {
                        comboBoxCell.Items.AddRange(GetSubjectList());
                    }

                    // Начинаем редактирование ячейки
                    dataGridView.BeginEdit(true);

                    // Устанавливаем стиль отображения для комбо-бокса
                    comboBoxCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
                }
            }
        }

        //Проверка на ввод DataGridView

        private void dataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is TextBox textbox)
            {
                textbox.KeyPress += Textbox_KeyPress;
            }
        }

        private void Textbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex > 0)
            {
                // Получаем активный элемент редактирования
                var activeControl = this.ActiveControl;

                if (activeControl is TextBox editingTextBox)
                {
                    // При завершении редактирования удаляем обработчик события KeyPress
                    editingTextBox.KeyPress -= Textbox_KeyPress;
                }
            }
        }

        /////////////////////////////////////////////////////////


        private List<string> GetSubjectList()
        {
            return new List<string> { "Математика", "Химия", "Физика", "Литература", "История", "География" };
        }

        private string[] GetSubjectArray()
        {
            string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
            string templatePath = Path.Combine(projectPath, "source", "предметы.txt");

            if (File.Exists(templatePath))
            {
                return File.ReadAllLines(templatePath);
            }
            return new string[0];
        }

        private void dataGridView_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            DataGridView dataGridView = (DataGridView)sender;

            for (int i = e.RowIndex; i < dataGridView.Rows.Count; i++)
            {
                dataGridView.Rows[i].Cells[0].Value = (i + 1).ToString();
            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {

        }


        //СОХРАНЕНИЕ

        private void GenerateWordDocument()
        {
            try
            {
                string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
                string templatePath = Path.Combine(projectPath, "source", "расписание 2023-2024.docx");
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Документ Word (*.doc)|*.doc";
                saveFileDialog.Title = "Сохранить документ Word";
                saveFileDialog.FileName = $"{this.Text}.doc";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string outputPath = @saveFileDialog.FileName;
                    MessageBox.Show(outputPath);
                    File.Copy(templatePath, outputPath, true);

                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
                    {
                        for (int tabIndex = 0; tabIndex < tab.TabPages.Count; tabIndex++)
                        {
                            var tabPage = tab.TabPages[tabIndex];

                            int classNumber = int.Parse(tabPage.Text);
                            string wordClassNumber = classNumber.ToWords();

                            foreach (System.Windows.Forms.Control control in tabPage.Controls)
                            {
                                if (control is DataGridView dataGridView)
                                {
                                    for (int rowIndex = 0; rowIndex < dataGridView.Rows.Count - 1; rowIndex++)
                                    {
                                        for (int columnIndex = 1; columnIndex < dataGridView.Columns.Count; columnIndex++)
                                        {
                                            string cellValue = dataGridView[columnIndex, rowIndex].Value?.ToString();
                                            string bookmarkName = GetBookmarkName(wordClassNumber, columnIndex, rowIndex + 1);

                                            ReplaceContentControlText(wordDoc, bookmarkName, cellValue);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    MessageBox.Show(outputPath);
                }
                else MessageBox.Show("Отменено пользователем", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string CleanPath(string path)
        {
            string invalidFileNameChars = new string(Path.GetInvalidFileNameChars());
            string invalidPathChars = new string(Path.GetInvalidPathChars());

            string cleanedFileName = new string(path.Where(c => !invalidFileNameChars.Contains(c)).ToArray());
            string cleanedPath = new string(cleanedFileName.Where(c => !invalidPathChars.Contains(c)).ToArray());

            return cleanedPath;
        }


        private string GetBookmarkName(string wordClassNumber, int columnIndex, int rowIndex)
        {
            if (columnIndex % 2 == 0)
            {
                return $"{wordClassNumber}{columnIndex - 1}{rowIndex}{rowIndex}";
            }
            else
            {
                return $"{wordClassNumber}{columnIndex}{rowIndex}";
            }
        }

        private void ReplaceContentControlText(WordprocessingDocument wordDoc, string controlName, string text)
        {
            // Поиск текстового поля по имени в документе Word
            SdtElement contentControl = wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>()
                .FirstOrDefault(c => c.SdtProperties.Elements<SdtAlias>().FirstOrDefault()?.Val == controlName);

            if (contentControl != null)
            {
                // Вставка текста в текстовое поле
                Run run = contentControl.Descendants<Run>().FirstOrDefault();
                if (run != null)
                {
                    Text textElement = run.Descendants<Text>().FirstOrDefault();
                    if (textElement != null)
                    {
                        textElement.Text = text;
                    }
                }
            }
        }



        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void сохранитьToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (tab != null && tab.TabCount > 0)
            {
                GenerateWordDocument();

            }
            else MessageBox.Show("Чтобы сохранить файл его надобно создать", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///

        //Печать документа

        private void PrintWordDocument()
        {
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();

                var doc = wordApp.Documents.Open(outputPath);

                object missing = Missing.Value;

                doc.PrintOut(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                // Закрывает Word
                doc.Close(ref missing, ref missing, ref missing);
                wordApp.Quit(ref missing, ref missing, ref missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                PrintWordDocument();
            }
            else MessageBox.Show("Нельзя отправлять на печать несуществующий документ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void UpdateComboBoxItems()
        {
            foreach(TabPage tabpage in tab.TabPages)
            {
                foreach(System.Windows.Forms.Control control in tabpage.Controls)
                {
                    if(control is DataGridView dataGridView)
                    {
                        foreach(DataGridViewRow row in dataGridView.Rows)
                        {
                            for(int i = 1; i < row.Cells.Count; i += 2)
                            {
                                DataGridViewComboBoxCell comboBoxSell = (DataGridViewComboBoxCell)row.Cells[i];
                                comboBoxSell.Items.Clear();
                                comboBoxSell.Items.AddRange(GetSubjectArray());
                            }
                        }
                    }
                }
            }
        }

        ///////////////////////////////////////////////////////////////////////
    }

}

