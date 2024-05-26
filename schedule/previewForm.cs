using System.Drawing;
using System.Linq;
using System.Windows.Forms;


namespace schedule
{
    public partial class previewForm : Form
    {



        private DataGridView previewDataGridView;
        public previewForm()
        {
            InitializeComponent();
            InitializeDataGridView();
        }

        private void InitializeDataGridView()
        {
            previewDataGridView = new DataGridView();
            previewDataGridView.Dock = DockStyle.Fill;
            previewDataGridView.AllowUserToAddRows = false; // Запретить добавление строк
            previewDataGridView.AllowUserToDeleteRows = false; // Запретить удаление строк
            previewDataGridView.ReadOnly = true; // Сделать DataGridView только для чтения
            previewDataGridView.BackgroundColor = SystemColors.ActiveCaption;
            previewDataGridView.DefaultCellStyle.BackColor = SystemColors.GradientActiveCaption;
            this.Controls.Add(previewDataGridView);
        }

        public void LoadSchedule(TabControl mainTabControl, string[] daysOfWeek)
        {
            // Настраиваем столбцы DataGridView
            previewDataGridView.Columns.Add("Day", "День недели");
            previewDataGridView.Columns.Add("RowNumber", "#");

            for (int i = 1; i <= 11; i++)
            {
                previewDataGridView.Columns.Add($"Class{i}", $"{i} класс");
                previewDataGridView.Columns.Add($"Room{i}", "Кабинет");
            }

            // Заполняем строки данными из каждого TabControl
            for (int dayIndex = 0; dayIndex < daysOfWeek.Length; dayIndex++)
            {
                for (int rowIndex = 0; rowIndex < 7; rowIndex++) // Предполагая, что у нас 10 строк для каждого дня
                {
                    DataGridViewRow newRow = new DataGridViewRow();
                    newRow.CreateCells(previewDataGridView);

                    if (rowIndex == 0) // Только для первой строки каждого дня
                    {
                        newRow.Cells[0].Value = daysOfWeek[dayIndex]; // День недели
                    }

                    newRow.Cells[1].Value = rowIndex + 1; // Номер строки

                    for (int classIndex = 0; classIndex < mainTabControl.TabPages.Count; classIndex++)
                    {
                        DataGridView dgv = mainTabControl.TabPages[classIndex].Controls.OfType<DataGridView>().FirstOrDefault();
                        if (dgv != null && dgv.Rows.Count > rowIndex)
                        {
                            newRow.Cells[2 + classIndex * 2].Value = dgv.Rows[rowIndex].Cells[2 * dayIndex + 1].Value; // Четный столбец (предмет)
                            newRow.Cells[3 + classIndex * 2].Value = dgv.Rows[rowIndex].Cells[2 * dayIndex + 2].Value; // Нечетный столбец (кабинет)
                        }
                    }

                    previewDataGridView.Rows.Add(newRow);
                }
            }
        }
    }
}
