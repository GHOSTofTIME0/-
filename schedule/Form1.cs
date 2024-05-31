using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Humanizer;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace schedule
{
    public partial class Form1 : Form
    {
        private TabControl tab;
        private DataGridView dataGridView1;
        private readonly string[] daysOfWeek = { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота" };
        private string outputPath;
        private static string projectPath = Directory.GetParent(System.Windows.Forms.Application.StartupPath).Parent.Parent.FullName;
        public  string connectString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Path.Combine(projectPath, "schedule", "source", "shedule.mdb")}";

        private Dictionary<int, Dictionary<string, int>> difficultyLimits = new Dictionary<int, Dictionary<string, int>>
{
    { 1, new Dictionary<string, int> { { "Понедельник", 21 }, { "Вторник", 23 }, { "Среда", 24 }, { "Четверг", 26 }, { "Пятница", 16 } } },
    { 2, new Dictionary<string, int> { { "Понедельник", 21 }, { "Вторник", 23 }, { "Среда", 24 }, { "Четверг", 26 }, { "Пятница", 16 } } },
    { 3, new Dictionary<string, int> { { "Понедельник", 22 }, { "Вторник", 21 }, { "Среда", 23 }, { "Четверг", 27 }, { "Пятница", 19 } } },
    { 4, new Dictionary<string, int> { { "Понедельник", 22 }, { "Вторник", 21 }, { "Среда", 26 }, { "Четверг", 23 }, { "Пятница", 19 } } },
    { 5, new Dictionary<string, int> { { "Понедельник", 30 }, { "Вторник", 32 }, { "Среда", 29 }, { "Четверг", 31 }, { "Пятница", 38 } } },
    { 6, new Dictionary<string, int> { { "Понедельник", 30 }, { "Вторник", 32 }, { "Среда", 29 }, { "Четверг", 31 }, { "Пятница", 38 } } },
    { 7, new Dictionary<string, int> { { "Понедельник", 43 }, { "Вторник", 42 }, { "Среда", 39 }, { "Четверг", 24 }, { "Пятница", 36 } } },
    { 8, new Dictionary<string, int> { { "Понедельник", 39 }, { "Вторник", 43 }, { "Среда", 47 }, { "Четверг", 21 }, { "Пятница", 36 } } },
    { 9, new Dictionary<string, int> { { "Понедельник", 43 }, { "Вторник", 49 }, { "Среда", 48 }, { "Четверг", 47 }, { "Пятница", 36 } } },
    { 10, new Dictionary<string, int> { { "Понедельник", 53 }, { "Вторник", 52 }, { "Среда", 34 }, { "Четверг", 34 }, { "Пятница", 43 } } },
    { 11, new Dictionary<string, int> { { "Понедельник", 53 }, { "Вторник", 52 }, { "Среда", 34 }, { "Четверг", 34 }, { "Пятница", 43 } } }
};


        public Dictionary<int, Dictionary<string, int>> headersI = new Dictionary<int, Dictionary<string, int>>();

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


        /////////////////////////////////////////// TABCONTROL ///////////////////////////////////////////////////////
        private void InitializeTab()
        {
            tab = new TabControl();

            tab.Size = new Size((System.Drawing.Point)this.Size);
            tab.BackColor = SystemColors.ActiveCaption;
            tab.DrawMode = TabDrawMode.OwnerDrawFixed;
            tab.DrawItem += new DrawItemEventHandler(tab_DrawItem);
            for (int i = 1; i < 12; i++)
            {
                TabPage tp = new TabPage("Класс" + Convert.ToString(i));
                DataGridView Grid = CreateDataGridView(i);
                tp.Controls.Add(Grid);
                tab.TabPages.Add(tp);

            }
            this.Controls.Add(tab);

            // Убедитесь, что Controls.Add(tab) вызывается после добавления вкладок
            Controls.Add(tab);
            tab.Location = new System.Drawing.Point(0, menuStrip1.Bottom);
        }

        private void tab_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabPage tp = tab.TabPages[e.Index];
            Rectangle tabBounds = tab.GetTabRect(e.Index);
            using(SolidBrush brush = new SolidBrush(SystemColors.GradientActiveCaption))
            {
                e.Graphics.FillRectangle(brush, tabBounds);
            }

            using(SolidBrush textBrush = new SolidBrush(System.Drawing.Color.Black))
            {
                StringFormat stringFlags = new StringFormat();
                stringFlags.Alignment = StringAlignment.Center;
                stringFlags.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(tp.Text, e.Font, textBrush, tabBounds, stringFlags);
            }
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

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



        ///////////////////////////////////////////    DATAGRIDVIEW   ///////////////////////////////////////////////////////

        private DataGridView CreateDataGridView(int classNumber)
        {
            DataGridView dataGridView = new DataGridView();
            dataGridView.Dock = DockStyle.Fill;

            DataGridViewTextBoxColumn rowNumerColumn = new DataGridViewTextBoxColumn();
            rowNumerColumn.HeaderText = "#";
            dataGridView.Columns.Add(rowNumerColumn);
            dataGridView.AllowUserToAddRows = false;
            dataGridView.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView.BackgroundColor = SystemColors.ActiveCaption;
            dataGridView.DefaultCellStyle.BackColor = SystemColors.GradientActiveCaption;
            int maxRowsCount = 0;
            if (classNumber <= 4)
            {
                maxRowsCount = 5;
            }
            else maxRowsCount = 7;

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
                    subjectColumn.DefaultCellStyle.BackColor = SystemColors.GradientActiveCaption;
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
                        Dictionary<int, Dictionary<string, int>> headers = GetDisciplinesHeaders(connectString, tableName, classNumber);
                        headersI[classNumber] = headers[classNumber];
                        if (headers.ContainsKey(classNumber))
                        {
                            Dictionary<string, int> classHeaders = headers[classNumber];
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


            dataGridView.EditingControlShowing += dataGridView_EditingControlShowing;
            dataGridView.CellEndEdit += dataGridView_CellEndEdit;
            dataGridView.CurrentCellDirtyStateChanged += DataGridView_CurrentCellDirtyStateChanged;
            dataGridView.CellValueChanged += DataGridView_CellValueChanged;
            dataGridView1 = dataGridView;
            enabledMenuItems();

            return dataGridView;
        }

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





        private void DataGridView_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.IsCurrentCellDirty)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }



        private void DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridView currentDataGridView = (DataGridView)sender;
                string classNumberText = currentDataGridView.Parent.Text;
                string classNumberString = System.Text.RegularExpressions.Regex.Match(classNumberText, @"\d+").Value;

                if (!int.TryParse(classNumberString, out int classNumber))
                {
                    MessageBox.Show($"Невозможно определить номер класса из текста '{classNumberText}'", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string selectedDay = currentDataGridView.Columns[e.ColumnIndex % 2 == 0 ? e.ColumnIndex - 1 : e.ColumnIndex].HeaderText;

                if (e.ColumnIndex % 2 != 0) // Проверяем только ячейки с предметами (нечетные столбцы)
                {
                    DataGridViewComboBoxCell comboBoxCell = currentDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewComboBoxCell;
                    string selectedSubject = comboBoxCell?.Value?.ToString();
                    int roomColumnIndex = e.ColumnIndex + 1;
                    string currentRoomNumber = currentDataGridView.Rows[e.RowIndex].Cells[roomColumnIndex]?.Value?.ToString();

                    if (selectedSubject != null)
                    {
                        if (!headersI.ContainsKey(classNumber))
                        {
                            MessageBox.Show($"Отсутствует информация о дисциплинах для класса {classNumber}.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (!headersI[classNumber].ContainsKey(selectedSubject))
                        {
                            MessageBox.Show($"Отсутствует информация о предмете '{selectedSubject}' для класса {classNumber}.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        int totalDifficulty = 0;
                        foreach (DataGridViewRow row in currentDataGridView.Rows)
                        {
                            if (row.Cells[e.ColumnIndex] is DataGridViewComboBoxCell cell && cell.Value != null)
                            {
                                string subject = cell.Value.ToString();
                                if (headersI[classNumber].ContainsKey(subject))
                                {
                                    totalDifficulty += headersI[classNumber][subject];
                                }
                            }
                        }

                        if (!difficultyLimits.ContainsKey(classNumber))
                        {
                            MessageBox.Show($"Отсутствует информация о лимитах сложности для класса {classNumber}.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (!difficultyLimits[classNumber].ContainsKey(selectedDay))
                        {
                            MessageBox.Show($"Отсутствует информация о лимитах сложности для дня '{selectedDay}' в классе {classNumber}.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (totalDifficulty > difficultyLimits[classNumber][selectedDay])
                        {
                            MessageBox.Show($"Сумма баллов сложности предметов в день '{selectedDay}' ({totalDifficulty}) превышает допустимое значение для {classNumber} класса ({difficultyLimits[classNumber][selectedDay]})", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            comboBoxCell.Value = null; // Сбрасываем значение текущей ячейки
                            return; // Прерываем выполнение метода
                        }
                    }

                    if (selectedSubject != null && currentRoomNumber != null)
                    {
                        // Проверка конфликтов предметов и кабинетов
                        foreach (TabPage tabPage in tab.TabPages)
                        {
                            DataGridView dataGridView = tabPage.Controls[0] as DataGridView;

                            if (dataGridView != null)
                            {
                                for (int rowIndex = 0; rowIndex < dataGridView.Rows.Count; rowIndex++)
                                {
                                    if (dataGridView != currentDataGridView || rowIndex != e.RowIndex)
                                    {
                                        if (dataGridView.Rows[rowIndex].Cells[e.ColumnIndex] is DataGridViewComboBoxCell otherComboBoxCell)
                                        {
                                            string otherSubject = otherComboBoxCell?.Value?.ToString();
                                            string otherRoomNumber = dataGridView.Rows[rowIndex].Cells[roomColumnIndex]?.Value?.ToString();

                                            if (otherSubject == selectedSubject && otherRoomNumber == currentRoomNumber)
                                            {
                                                MessageBox.Show($"Предмет '{selectedSubject}' уже назначен для {tabPage.Text} класса в день '{selectedDay}' в кабинете '{currentRoomNumber}'", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                                comboBoxCell.Value = null; // Сбрасываем значение текущей ячейки
                                                currentDataGridView.Rows[e.RowIndex].Cells[roomColumnIndex].Value = null; // Сбрасываем значение кабинета
                                                return; // Прерываем выполнение метода
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else // Проверяем только ячейки с кабинетами (четные столбцы)
                {
                    DataGridViewTextBoxCell roomCell = currentDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewTextBoxCell;
                    string currentRoomNumber = roomCell?.Value?.ToString();
                    int subjectColumnIndex = e.ColumnIndex - 1;
                    string selectedSubject = currentDataGridView.Rows[e.RowIndex].Cells[subjectColumnIndex]?.Value?.ToString();

                    if (selectedSubject != null && currentRoomNumber != null)
                    {
                        // Проверка конфликтов предметов и кабинетов
                        foreach (TabPage tabPage in tab.TabPages)
                        {
                            DataGridView dataGridView = tabPage.Controls[0] as DataGridView;

                            if (dataGridView != null)
                            {
                                for (int rowIndex = 0; rowIndex < dataGridView.Rows.Count; rowIndex++)
                                {
                                    if (dataGridView != currentDataGridView || rowIndex != e.RowIndex)
                                    {
                                        if (dataGridView.Rows[rowIndex].Cells[subjectColumnIndex] is DataGridViewComboBoxCell otherComboBoxCell)
                                        {
                                            string otherSubject = otherComboBoxCell?.Value?.ToString();
                                            string otherRoomNumber = dataGridView.Rows[rowIndex].Cells[e.ColumnIndex]?.Value?.ToString();

                                            if (otherSubject == selectedSubject && otherRoomNumber == currentRoomNumber)
                                            {
                                                MessageBox.Show($"Кабинет '{currentRoomNumber}' уже назначен для {tabPage.Text} класса в день '{selectedDay}' с предметом '{selectedSubject}'", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                                roomCell.Value = null; // Сбрасываем значение текущей ячейки кабинета
                                                currentDataGridView.Rows[e.RowIndex].Cells[subjectColumnIndex].Value = null; // Сбрасываем значение предмета
                                                return; // Прерываем выполнение метода
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }



        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////














        private void disabledMenuItems()
        {
            предпросмотрРасписанияToolStripMenuItem.Enabled = false;
        }

        private void enabledMenuItems()
        {
            предпросмотрРасписанияToolStripMenuItem.Enabled = true;
        }



        //СОХРАНЕНИЕ И ПЕЧАТЬ

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

        private void PrintWordDocument() // не робит
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

       

        // МЕТОД ДЛЯ БДШКИ

        public Dictionary<int, Dictionary<string, int>> GetDisciplinesHeaders(string connectionString, string tableName, int classNumberI) 
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
                                if (tableName == "Disciplines5_9")
                                {
                                    classNumber += 4;
                                }
                                else if (tableName == "Disciplines10_11")
                                {
                                    classNumber += 9;
                                }


                                // Создаем словарь для текущего номера класса, если его еще нет
                                if (!headers.ContainsKey(classNumber))
                                    headers[classNumber] = new Dictionary<string, int>();

                                // Получаем имена столбцов кроме "Class"
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    string columnName = reader.GetName(i);
                                    if (columnName != "Class")
                                    {
                                        
                                        int fieldValue = Convert.ToInt32(reader[columnName]);
                                        
                                        headers[classNumber][columnName] = fieldValue;

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


        ///////////////////////////////// ПРЕДПРОСМОТР ///////////////////////////////////////////////////////////////////////

        private void предпросмотрРасписанияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            previewForm previewForm = new previewForm();
            previewForm.LoadSchedule(tab, daysOfWeek);
            previewForm.ShowDialog();
        }

        private void баллыПредметовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            infoSubjectsForm infoSubjectsForm = new infoSubjectsForm();
            infoSubjectsForm.Show();
        }
    }

}

