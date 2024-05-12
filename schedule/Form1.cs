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
using System.Xml.Linq;
using System.Data.OleDb;
using System.Data;
using DocumentFormat.OpenXml.Bibliography;

namespace schedule
{
    public partial class Form1 : Form
    {
        private TabControl tab;
        private readonly string[] daysOfWeek = { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота" };
        private string outputPath;

        private static string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
        private readonly string SubjectFilePath = Path.Combine(projectPath, "source", "предметы.txt");

        private readonly string teachersFilePath = Path.Combine(projectPath, "source", "учителя.txt");

        private readonly string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\ilya2\\source\\repos\\schedule\\schedule\\bin\\Debug\\shedule.mdb;";
        public Form1()
        {
            InitializeComponent();
            disabledMenuItems();
        }




        //СОЗДАНИЕ 


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
                DataGridView Grid = CreateDataGridView(i+1);
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

        private DataGridView CreateDataGridView(int classNumber)
        {
            DataGridView dataGridView = new DataGridView();
            dataGridView.Dock = DockStyle.Fill;

            DataGridViewTextBoxColumn rowNumerColumn = new DataGridViewTextBoxColumn();
            rowNumerColumn.HeaderText = "#";
            dataGridView.Columns.Add(rowNumerColumn);
            dataGridView.AllowUserToAddRows = false;
            int maxRowsCount = 0;
            if(classNumber <=4)
            {
                maxRowsCount = 5;
            }
            else maxRowsCount= 7;

            for (int rowIndex = 1; rowIndex <= maxRowsCount; rowIndex++)
            {
                dataGridView.Rows.Add(rowIndex.ToString());
            }


            for (int i = 0; i < 12; i++)
            {
                int classNumberI = i + 1;
                if (i % 2 == 0)
                {
                    DataGridViewComboBoxColumn subjectColumn = new DataGridViewComboBoxColumn();
                    subjectColumn.HeaderText = daysOfWeek[i / 2];
                    dataGridView.Columns.Add(subjectColumn);

                    // Проверяем номер класса и выбираем соответствующую таблицу
                    string tableName;
                    if (classNumber >= 1 && classNumber <= 4)
                        tableName = "Disciplines1_4";
                    else if (classNumber >= 5 && classNumber <= 9)
                        tableName = "Disciplines5_9";
                    else
                        tableName = "Disciplines10_11";

                    // Получаем названия полей из таблицы и заполняем ComboBox
                    Dictionary<int, Dictionary<string, int>> headers = GetDisciplinesHeaders(connectString, tableName, classNumberI);
                    if (headers.ContainsKey(i + 1))
                    {
                        Dictionary<string, int> classHeaders = headers[classNumberI];
                        foreach (var header in classHeaders)
                        {
                            if (header.Value != 0)
                            {
                                subjectColumn.Items.Add(header.Key);
                            }
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

            // Остальной код обработки событий и т.д.

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
            
            using (var inputForm = new Form2())
            {
                if(inputForm.ShowDialog() == DialogResult.OK)
                {
                    string newText = inputForm.nameText;
                    int difficulty = Convert.ToInt32(inputForm.difficulty);
                    AddSubjectWithDifficultyToFile(newText, difficulty);
                }
            }
            UpdateComboBoxItems();
        }

        private void AddSubjectWithDifficultyToFile(string subject, int difficulty)
        {
            string lineToAdd = $"{subject}:{difficulty}";
            File.AppendAllText(SubjectFilePath, lineToAdd+ "\r\n");
            UpdateComboBoxItems();
        }

        private Dictionary<string, int> ReadSubjectsAndDifficultiesFromFile()
        {
            Dictionary<string, int> subjectsAndDifficulty = new Dictionary<string, int>();

            if(File.Exists(SubjectFilePath))
            {
                string[] lines = File.ReadAllLines(SubjectFilePath);

                foreach (string line in lines)
                {
                    string[] parts = line.Split(':');

                    if(parts.Length== 2)
                    {
                        string subject = parts[0];
                        int difficulty;

                        if (int.TryParse(parts[1], out difficulty))
                        {
                            subjectsAndDifficulty.Add(subject, difficulty);
                        }
                        else MessageBox.Show("Неверно введена сложность", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else MessageBox.Show("Неверный формат строки", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return subjectsAndDifficulty;
        }
        

        private void disabledMenuItems()
        {
            добавитьПредметToolStripMenuItem.Enabled = false;
            удалитьПредметToolStripMenuItem.Enabled = false;
        }

        private void enabledMenuItems()
        {
            добавитьПредметToolStripMenuItem.Enabled = true;
            удалитьПредметToolStripMenuItem.Enabled = true;
        }

        ////////////////////////////////////добавление/удаление учителей////////////////////////////////////////////////////////////////////

        private void предметыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using(var inputForm = new Form4())
            {
                if(inputForm.ShowDialog() == DialogResult.OK)
                {
                    string lineToEnter =$"{inputForm.teacherFullName}:{inputForm.subjectName}";

                    if(File.Exists(teachersFilePath))
                    {
                        File.AppendAllText(teachersFilePath, lineToEnter + "\r\n");
                    }
                    UpdateComboBoxItems();
                }
            }
        }

        private void удалитьУчителяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using(var inputForm = new Form5())
            {
                inputForm.ShowDialog();
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

        private void ClearWordDocument(WordprocessingDocument wordDoc)
        {
            foreach (SdtElement contentControl in wordDoc.MainDocumentPart.Document.Body.Descendants<SdtElement>())
            {
                Run run = contentControl.Descendants<Run>().FirstOrDefault();
                if (run != null)
                {
                    Text textElement = run.Descendants<Text>().FirstOrDefault();
                    if (textElement != null)
                    {
                        textElement.Text = "";
                    }
                }
            }
        }

        private void GenerateWordDocuments()
        {
            try
            {
                string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
                string templatePath1_4 = Path.Combine(projectPath, "source", "началка.docx");
                string templatePath5_11 = Path.Combine(projectPath, "source", "расписание 2023-2024.docx");

                GenerateWordDocument(templatePath1_4, 0, 4); // Заполняем "началка.docx" на основе первых четырех TabControl
                GenerateWordDocument(templatePath5_11, 4, tab.TabPages.Count); // Заполняем "расписание 2023-2024.docx" на основе остальных TabControl
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GenerateWordDocument(string templatePath, int startTabIndex, int endTabIndex)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Документ Word (*.docx)|*.docx";
                saveFileDialog.Title = "Сохранить документ Word";
                saveFileDialog.FileName = $"{Path.GetFileNameWithoutExtension(templatePath)}_{startTabIndex + 1}-{endTabIndex}.docx"; // Имя файла на основе индексов вкладок

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string outputPath = saveFileDialog.FileName;
                    File.Copy(templatePath, outputPath, true);

                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
                    {
                        ClearWordDocument(wordDoc);

                        for (int tabIndex = startTabIndex; tabIndex < endTabIndex; tabIndex++)
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
                    MessageBox.Show($"Документ успешно создан и сохранен по пути: {outputPath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else MessageBox.Show("Отменено пользователем", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        /*private void GenerateWordDocument()
        {
            try
            {
                string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
                string templatePath5_11 = Path.Combine(projectPath, "source", "расписание 2023-2024.docx");
                string templatePath1_4 = Path.Combine(projectPath, "source", "началка.docx");
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Документ Word (*.doc)|*.doc";
                saveFileDialog.Title = "Сохранить документ Word";
                saveFileDialog.FileName = $"{this.Text}.doc";


                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string outputPath = @saveFileDialog.FileName;
                    MessageBox.Show(outputPath);
                    File.Copy(templatePath5_11, outputPath, true);

                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
                    {
                        ClearWordDocument(wordDoc);
                        for (int tabIndex = 4; tabIndex < tab.TabPages.Count; tabIndex++)
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
        }*/

        /* if(saveFileDialog.ShowDialog() == DialogResult.OK)
         {
             string outputPath = saveFileDialog.FileName;
             File.Copy(templatePath1_4, outputPath, true );

             using(WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
             {
                 ClearWordDocument(wordDoc);
                 for(int tabIndex = 0; tabIndex < tab.TabPages.Count; tabIndex++)
                 {
                     var tabPage = tab.TabPages[tabIndex];
                     int classNumber = int.Parse(tabPage.Text);
                     string wordClassNumber = classNumber.ToWords();

                     foreach(System.Windows.Forms.Control control in tabPage.Controls)
                     {
                         if(control is DataGridView dataGridView)
                         {
                             for(int rowIndex = 0; rowIndex < dataGridView.Rows.Count - 1; rowIndex++)
                             {
                                 for(int columnIndex = 1; columnIndex < dataGridView.Columns.Count; columnIndex++)
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
         }
         else MessageBox.Show("Отменено пользователем", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);*/
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
                GenerateWordDocuments();

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

       

        private Dictionary<int, Dictionary<string, int>> GetDisciplinesHeaders(string connectionString, string tableName, int classNumberI) // Доделать, баги ебобаные
        {
            Dictionary<int, Dictionary<string, int>> headers = new Dictionary<int, Dictionary<string, int>>();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Создаем SQL-запрос для получения всех данных из таблицы
                    string query = $"SELECT * FROM {tableName}";

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {

                                int classNumber = Convert.ToInt32(reader["Class"]);
/*
                                if (tableName == "Disciplines5_9") classNumber += 4;
                                if (tableName == "Disciplines10_11") classNumber += 9;*/

                                // Создаем словарь для текущего номера класса, если его еще нет
                                if (!headers.ContainsKey(classNumberI))
                                    headers[classNumberI] = new Dictionary<string, int>();

                                // Получаем имена столбцов кроме "Class"
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    string columnName = reader.GetName(i);
                                    if (columnName != "Class")
                                    {
                                        // Добавляем заголовок и балл в словарь для текущего номера класса
                                        int fieldValue = Convert.ToInt32(reader[columnName]);
                                        if (fieldValue != 0)
                                        {
                                            headers[classNumberI][columnName] = fieldValue;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (OleDbException ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
            }

            return headers;
        }

        
        }

}

