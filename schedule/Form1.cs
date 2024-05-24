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
            tab.BackColor = SystemColors.ActiveCaption;
            tab.DrawMode = TabDrawMode.OwnerDrawFixed;
            tab.DrawItem += new DrawItemEventHandler(tab_DrawItem);
            for (int i = 1; i < 12; i++)
            {
                TabPage tp = new TabPage(Convert.ToString(i));
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
            /*dataGridView.CellValidating += DataGridView_CellValidating;*/
            dataGridView.CurrentCellDirtyStateChanged += DataGridView_CurrentCellDirtyStateChanged;
            dataGridView.CellValueChanged += DataGridView_CellValueChanged;
            dataGridView1 = dataGridView;
            enabledMenuItems();
            // Остальной код обработки событий и т.д.

            return dataGridView;
        }





        //////////////////////////////////////добавление и удаление предметов//////////////////////////////////////////////////////////////////

        private void удалитьПредметToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var subjectsListForm = new Form3())
            {
                subjectsListForm.ShowDialog();
            }
            /*UpdateComboBoxItems();*/
        }

        private void добавитьПредметToolStripMenuItem_Click(object sender, EventArgs e)
        {

            using (var inputForm = new Form2())
            {
                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    string newText = inputForm.nameText;
                    int difficulty = Convert.ToInt32(inputForm.difficulty);

                }
            }
            /*UpdateComboBoxItems();*/
        }






        private void disabledMenuItems()
        {
            добавитьПредметToolStripMenuItem.Enabled = false;
            удалитьПредметToolStripMenuItem.Enabled = false;
            предпросмотрРасписанияToolStripMenuItem.Enabled = false;
        }

        private void enabledMenuItems()
        {
            добавитьПредметToolStripMenuItem.Enabled = true;
            удалитьПредметToolStripMenuItem.Enabled = true;
            предпросмотрРасписанияToolStripMenuItem.Enabled = true;
        }

        ////////////////////////////////////добавление/удаление учителей////////////////////////////////////////////////////////////////////

        private void предметыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var inputForm = new Form4())
            {
                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    string lineToEnter = $"{inputForm.teacherFullName}:{inputForm.subjectName}";

                    if (File.Exists(teachersFilePath))
                    {
                        File.AppendAllText(teachersFilePath, lineToEnter + "\r\n");
                    }
                    /*UpdateComboBoxItems();*/
                }
            }
        }

        private void удалитьУчителяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var inputForm = new Form5())
            {
                inputForm.ShowDialog();
            }
            /*UpdateComboBoxItems();*/
        }


        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////////////////////////////////////////////////////////////////////////////////////



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





        private void DataGridView_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.IsCurrentCellDirty)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }


        private void DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex % 2 != 0) // Проверяем только ячейки с предметами (четные столбцы)
            {
                DataGridView currentDataGridView = (DataGridView)sender;
                DataGridViewComboBoxCell comboBoxCell = currentDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewComboBoxCell;
                string selectedSubject = comboBoxCell?.Value?.ToString();
                string selectedDay = currentDataGridView.Columns[e.ColumnIndex].HeaderText;
                int roomColumnIndex = e.ColumnIndex + 1;
                string currentRoomNumber = currentDataGridView.Rows[e.RowIndex].Cells[roomColumnIndex]?.Value?.ToString();

                if (selectedSubject != null && currentRoomNumber != null)
                {
                    // Проходим по всем вкладкам
                    foreach (TabPage tabPage in tab.TabPages)
                    {
                        DataGridView dataGridView = tabPage.Controls[0] as DataGridView;
                        string classNumber = tabPage.Text;

                        if (dataGridView != null)
                        {
                            // Проверяем все строки данного DataGridView
                            for (int rowIndex = 0; rowIndex < dataGridView.Rows.Count; rowIndex++)
                            {
                                if (dataGridView != currentDataGridView || rowIndex != e.RowIndex) // Пропускаем текущую строку в текущем DataGridView
                                {
                                    if (dataGridView.Rows[rowIndex].Cells[e.ColumnIndex] is DataGridViewComboBoxCell otherComboBoxCell)
                                    {
                                        string otherSubject = otherComboBoxCell?.Value?.ToString();
                                        string otherRoomNumber = dataGridView.Rows[rowIndex].Cells[roomColumnIndex]?.Value?.ToString();

                                        if (otherSubject == selectedSubject && otherRoomNumber == currentRoomNumber)
                                        {
                                            MessageBox.Show($"Предмет '{selectedSubject}' уже назначен для {classNumber} класса в день '{selectedDay}' в кабинете '{currentRoomNumber}'", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            else if (e.RowIndex >= 0 && e.ColumnIndex % 2 == 0) // Проверяем только ячейки с кабинетами (нечетные столбцы)
            {
                DataGridView currentDataGridView = (DataGridView)sender;
                DataGridViewTextBoxCell roomCell = currentDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewTextBoxCell;
                string currentRoomNumber = roomCell?.Value?.ToString();
                string selectedDay = currentDataGridView.Columns[e.ColumnIndex - 1].HeaderText;
                int subjectColumnIndex = e.ColumnIndex - 1;
                string selectedSubject = currentDataGridView.Rows[e.RowIndex].Cells[subjectColumnIndex]?.Value?.ToString();

                if (selectedSubject != null && currentRoomNumber != null)
                {
                    // Проходим по всем вкладкам
                    foreach (TabPage tabPage in tab.TabPages)
                    {
                        DataGridView dataGridView = tabPage.Controls[0] as DataGridView;
                        string classNumber = tabPage.Text;

                        if (dataGridView != null)
                        {
                            // Проверяем все строки данного DataGridView
                            for (int rowIndex = 0; rowIndex < dataGridView.Rows.Count; rowIndex++)
                            {
                                if (dataGridView != currentDataGridView || rowIndex != e.RowIndex) // Пропускаем текущую строку в текущем DataGridView
                                {
                                    if (dataGridView.Rows[rowIndex].Cells[subjectColumnIndex] is DataGridViewComboBoxCell otherComboBoxCell)
                                    {
                                        string otherSubject = otherComboBoxCell?.Value?.ToString();
                                        string otherRoomNumber = dataGridView.Rows[rowIndex].Cells[e.ColumnIndex]?.Value?.ToString();

                                        if (otherSubject == selectedSubject && otherRoomNumber == currentRoomNumber)
                                        {
                                            MessageBox.Show($"Кабинет '{currentRoomNumber}' уже назначен для {classNumber} класса в день '{selectedDay}' с предметом '{selectedSubject}'", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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


        /////////////////////////////////////////////////////////








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

        /*  private void UpdateComboBoxItems() //потом поменять
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
          }*/



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
                                        // Добавляем заголовок и балл в словарь для текущего номера класса
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


        private void предпросмотрРасписанияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            previewForm previewForm = new previewForm();
            previewForm.LoadSchedule(tab, daysOfWeek);
            previewForm.ShowDialog();
        }
    }

}

