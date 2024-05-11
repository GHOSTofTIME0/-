﻿using System;
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

namespace schedule
{
    public partial class Form1 : Form
    {
        private TabControl tab;
        private DataGridView dataGridView1;
        private string[] daysOfWeek = { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота" };
        private string outputPath;

        private static string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
        private string SubjectFilePath = Path.Combine(projectPath, "source", "предметы.txt");

        private string teachersFilePath = Path.Combine(projectPath, "source", "учителя.txt");

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
                    if (classNumber>=1 && classNumber <=4)
                        tableName = "Disciplines1_4";
                    else if (classNumber >= 5 && classNumber <= 9)
                        tableName = "Disciplines5_9";
                    else
                        tableName = "Disciplines10_11";

                    // Получаем названия полей из таблицы и заполняем ComboBox
                    Dictionary<int, List<string>> headers = GetDisciplinesHeaders(connectString, tableName, classNumberI);
                    if (headers.ContainsKey(i + 1))
                    {
                        foreach (string header in headers[i + 1])
                        {
                            subjectColumn.Items.Add(header);
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

        /* private DataGridView CreateDataGridView()
         {
             DataGridView dataGridView = new DataGridView();
             dataGridView.Dock = DockStyle.Fill; // Используйте DockStyle.Fill

             DataGridViewTextBoxColumn rowNumerColumn = new DataGridViewTextBoxColumn();
             rowNumerColumn.HeaderText = "#";
             dataGridView.Columns.Add(rowNumerColumn);



             for (int i = 0; i < 12; i++)
             {

                 if (i % 2 == 0)
                 {
                     DataGridViewComboBoxColumn subjectColumn = new DataGridViewComboBoxColumn();
                     subjectColumn.HeaderText = daysOfWeek[i / 2];
                     //subjectColumn.Items.AddRange(GetSubjectList().ToArray());
                     dataGridView.Columns.Add(subjectColumn);

                     if(File.Exists(SubjectFilePath))
                     {
                         string[] lines = File.ReadAllLines(SubjectFilePath);

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

             enabledMenuItems();

             return dataGridView;
         }*/

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
        ///
        /*public Dictionary<int, List<string>> GetDisciplinesHeaders(string connectionString, string tableName)
        {
            Dictionary<int, List<string>> headers = new Dictionary<int, List<string>>();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Получаем метаданные о столбцах таблицы
                    System.Data.DataTable schemaTable = connection.GetSchema("Columns", new string[] { null, null, tableName });

                    // Обрабатываем полученные данные
                    foreach (DataRow row in schemaTable.Rows)
                    {
                        string columnName = row["COLUMN_NAME"].ToString();

                        // Определяем номер класса из столбца "Key"
                        if (columnName == "Class")
                        {
                            int classNumber = Convert.ToInt32(row["TABLE_NAME"].ToString().Substring(row["TABLE_NAME"].ToString().Length - 1));

                            // Для таблицы Disciplines5_9 и Disciplines10_11 номер класса следует увеличить на 4 и 9 соответственно
                            if (tableName == "Disciplines5_9")
                                classNumber +=4;
                            else if (tableName == "Disciplines10_11")
                                classNumber += 9;

                            // Добавляем номер класса в словарь
                            if(!headers.ContainsKey(classNumber)) headers[classNumber] = new List<string>();
                        }
                        else
                        {
                            // Добавляем имя столбца в список заголовков для текущего номера класса
                            headers[int.Parse(row["TABLE_NAME"].ToString().Substring(row["TABLE_NAME"].ToString().Length - 1))].Add(columnName);

                        }
                    }
                }
                catch (OleDbException ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
            }

            return headers;
        }*/
        /*public Dictionary<int, List<string>> GetDisciplinesHeaders(string connectionString, string tableName)
        {
            Dictionary<int, List<string>> headers = new Dictionary<int, List<string>>();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Получаем метаданные о столбцах таблицы
                    System.Data.DataTable schemaTable = connection.GetSchema("Columns", new string[] { null, null, tableName });

                    // Обрабатываем полученные данные
                    foreach (DataRow row in schemaTable.Rows)
                    {
                        string columnName = row["COLUMN_NAME"].ToString();

                        // Определяем номер класса из индекса строки в таблице
                        int classNumber = schemaTable.Rows.IndexOf(row) + 1;

                        // Добавляем имя столбца в список заголовков для текущего номера класса
                        if (!headers.ContainsKey(classNumber))
                            headers[classNumber] = new List<string>();

                        headers[classNumber].Add(columnName);
                    }
                }
                catch (OleDbException ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
            }

            return headers;
        }*/
        /*public Dictionary<int, List<string>> GetDisciplinesHeaders(string connectionString, string tableName)
        {
            Dictionary<int, List<string>> headers = new Dictionary<int, List<string>>();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Получаем метаданные о столбцах таблицы
                    System.Data.DataTable schemaTable = connection.GetSchema("Columns", new string[] { null, null, tableName });

                    // Счетчик для определения номера строки
                    int rowNumber = 1;
                    if (tableName == "Disciplines5_9") rowNumber += 4;
                    if (tableName == "Disciplines10_11") rowNumber += 9;

                    // Обрабатываем полученные данные
                    foreach (DataRow row in schemaTable.Rows)
                    {
                        string columnName = row["COLUMN_NAME"].ToString();

                        // Добавляем имя столбца в список заголовков для текущего номера строки
                        if (!headers.ContainsKey(rowNumber))
                        {
                            headers[rowNumber] = new List<string>();
                            if (columnName != "Class")
                                headers[rowNumber].Add(columnName);
                        }

                        // Добавляем в список заголовков все поля, кроме "Class"

                        // Увеличиваем счетчик номера строки
                        rowNumber++;
                    }
                }
                catch (OleDbException ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
            }

            return headers;
        }*/
        public Dictionary<int, List<string>> GetDisciplinesHeaders(string connectionString, string tableName, int classNumber)
        {
            Dictionary<int, List<string>> headers = new Dictionary<int, List<string>>();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Получаем метаданные о столбцах таблицы
                    System.Data.DataTable schemaTable = connection.GetSchema("Columns", new string[] { null, null, tableName });

                    // Переменная для отслеживания номера строки (класса)

                    // Обрабатываем полученные данные
                    foreach (DataRow row in schemaTable.Rows)
                    {
                        string columnName = row["COLUMN_NAME"].ToString();

                        // Пропускаем столбец "Class"
                        if (columnName == "Class")
                            continue;

                        // Добавляем имя столбца в список заголовков для текущего номера класса
                        if (!headers.ContainsKey(classNumber))
                            headers[classNumber] = new List<string>();

                        // Добавляем в список заголовков все поля, кроме столбца "Class"
                        if (columnName != "Class")
                            headers[classNumber].Add(columnName);

                        // Увеличиваем номер строки (класса)
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

