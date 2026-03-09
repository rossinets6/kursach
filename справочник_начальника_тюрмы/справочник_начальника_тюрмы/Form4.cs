using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace справочник_начальника_тюрмы
{
    public partial class Zeki : Form
    {
        private DatabaseService dbService = new DatabaseService();
        private DataTable dataTable = new DataTable();
        private string currentFilter = "";
        private string currentSort = "ASC";

        public Zeki()
        {
            InitializeComponent();
            LoadData();

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.ReadOnly = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;

            radioButton1.CheckedChanged += RadioButton_CheckedChanged;
            radioButton2.CheckedChanged += RadioButton_CheckedChanged;
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
            dataGridView1.CellBeginEdit += dataGridView1_CellBeginEdit;
            dataGridView1.CellDoubleClick += dataGridView1_CellDoubleClick;
            textBox1.KeyPress += textBox1_KeyPress;
        }

        private void LoadData(string filter = "", string sortOrder = "ASC")
        {
            try
            {
                dbService.openConnection();

                string query = @"SELECT [IDZak], [Имя], [Фамилия], [Отчество], [Дата_рождения],
                                        [Место_рождения], [Номер_дела], [Статья], [Суд],
                                        [Дата_приговора], [Срок], [Дата_заключения], [Дата_освобождения]
                                 FROM [spravochnik].[dbo].[Zaklichenie]";

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    query += " WHERE [Фамилия] LIKE @filter OR [Номер_дела] LIKE @filter";
                }

                query += " ORDER BY [Фамилия] " + sortOrder + ", [Имя] ASC";

                using (SqlCommand cmd = new SqlCommand(query, dbService.getConnection()))
                {
                    if (!string.IsNullOrWhiteSpace(filter))
                    {
                        cmd.Parameters.AddWithValue("@filter", $"%{filter}%");
                    }

                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        dataTable.Clear();
                        adapter.Fill(dataTable);
                    }
                }

                dataGridView1.DataSource = dataTable;

                ConfigureColumnHeaders();

                currentFilter = filter;
                currentSort = sortOrder;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                dbService.closeConnection();
            }
        }

        private void ConfigureColumnHeaders()
        {
            if (dataGridView1.Columns.Count == 0) return;

            dataGridView1.Columns["IDZak"].HeaderText = "Код";
            dataGridView1.Columns["Имя"].HeaderText = "Имя";
            dataGridView1.Columns["Фамилия"].HeaderText = "Фамилия";
            dataGridView1.Columns["Отчество"].HeaderText = "Отчество";
            dataGridView1.Columns["Дата_рождения"].HeaderText = "Дата рождения";
            dataGridView1.Columns["Место_рождения"].HeaderText = "Место рождения";
            dataGridView1.Columns["Номер_дела"].HeaderText = "Номер дела";
            dataGridView1.Columns["Статья"].HeaderText = "Статья";
            dataGridView1.Columns["Суд"].HeaderText = "Суд";
            dataGridView1.Columns["Дата_приговора"].HeaderText = "Дата приговора";
            dataGridView1.Columns["Срок"].HeaderText = "Срок (лет)";
            dataGridView1.Columns["Дата_заключения"].HeaderText = "Дата заключения";
            dataGridView1.Columns["Дата_освобождения"].HeaderText = "Дата освобождения";

            dataGridView1.Columns["IDZak"].ReadOnly = true;

            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                if (col.Name.Contains("Дата"))
                {
                    col.DefaultCellStyle.Format = "dd.MM.yyyy";
                }
            }
        }

        private void SaveCurrentChanges()
        {
            try
            {
                if (dataGridView1.IsCurrentCellInEditMode)
                    dataGridView1.EndEdit();

                if (dataTable?.GetChanges() == null)
                    return;

                string query = "SELECT [IDZak], [Имя], [Фамилия], [Отчество], [Дата_рождения], " +
                               "[Место_рождения], [Номер_дела], [Статья], [Суд], [Дата_приговора], " +
                               "[Срок], [Дата_заключения], [Дата_освобождения] FROM [spravochnik].[dbo].[Zaklichenie]";

                using (SqlCommand cmd = new SqlCommand(query, dbService.getConnection()))
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    SqlCommandBuilder commandBuilder = new SqlCommandBuilder(adapter);
                    adapter.InsertCommand = commandBuilder.GetInsertCommand();
                    adapter.UpdateCommand = commandBuilder.GetUpdateCommand();
                    adapter.DeleteCommand = commandBuilder.GetDeleteCommand();

                    dbService.openConnection();
                    int rowsUpdated = adapter.Update(dataTable);
                    dbService.closeConnection();

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                        row.DefaultCellStyle.BackColor = Color.White;

                    dataTable.AcceptChanges();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            string sortOrder = radioButton1.Checked ? "ASC" : "DESC";
            LoadData(currentFilter, sortOrder);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            DobZak dobzakForm = new DobZak();
            if (dobzakForm.ShowDialog() == DialogResult.OK)
            {
                LoadData(currentFilter, currentSort);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            MessageBox.Show("Изменения сохранены", "Успех",
                          MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите запись для удаления", "Внимание",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int id = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["IDZak"].Value);
            string fam = dataGridView1.SelectedRows[0].Cells["Фамилия"].Value?.ToString() ?? "";
            string name = dataGridView1.SelectedRows[0].Cells["Имя"].Value?.ToString() ?? "";

            if (MessageBox.Show($"Удалить заключённого {fam} {name}?", "Подтверждение",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    dbService.openConnection();
                    string query = "DELETE FROM [Zaklichenie] WHERE IDZak = @id";
                    SqlCommand cmd = new SqlCommand(query, dbService.getConnection());
                    cmd.Parameters.AddWithValue("@id", id);
                    int rows = cmd.ExecuteNonQuery();

                    if (rows > 0)
                    {
                        MessageBox.Show("Запись удалена", "Успех",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadData(currentFilter, currentSort);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления: {ex.Message}", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    dbService.closeConnection();
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            LoadData(textBox1.Text.Trim(), currentSort);
            if (dataTable.Rows.Count == 0)
                MessageBox.Show("Ничего не найдено", "Поиск",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataTable?.GetChanges() != null)
            {
                var result = MessageBox.Show("Сохранить изменения перед выходом?",
                                             "Несохранённые данные",
                                             MessageBoxButtons.YesNoCancel,
                                             MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                    SaveCurrentChanges();
                else if (result == DialogResult.Cancel)
                    return;
            }

            Vibor vibor = new Vibor();
            vibor.Show();
            Close();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                button4.PerformClick();
                e.Handled = true;
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "IDZak")
            {
                e.Cancel = true;
                MessageBox.Show("Это поле нельзя редактировать", "Внимание",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (dataGridView1.Columns[e.ColumnIndex].Name != "IDZak")
                dataGridView1.BeginEdit(true);
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name.Contains("Дата"))
            {
                if (e.FormattedValue != null && !string.IsNullOrWhiteSpace(e.FormattedValue.ToString()))
                {
                    if (!DateTime.TryParse(e.FormattedValue.ToString(), out _))
                    {
                        MessageBox.Show("Введите корректную дату (ДД.ММ.ГГГГ)", "Ошибка",
                                      MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                // Создаем объект приложения Excel
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // Экспортируем заголовки столбцов
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                    worksheet.Cells[1, i + 1].Font.Bold = true; // Делаем заголовки жирными
                }

                // Экспортируем данные
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++) // -1 чтобы не включать строку для добавления
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }

                // Автоматическая настройка ширины столбцов
                worksheet.Columns.AutoFit();

                // Открываем диалоговое окно сохранения файла
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.FileName = "Заключённые" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                    workbook.Close();
                    excelApp.Quit();

                    // Освобождаем ресурсы
                    ReleaseObject(worksheet);
                    ReleaseObject(workbook);
                    ReleaseObject(excelApp);

                    MessageBox.Show("Данные успешно экспортированы в Excel!");
                }
                else
                {
                    workbook.Close(false);
                    excelApp.Quit();

                    // Освобождаем ресурсы
                    ReleaseObject(worksheet);
                    ReleaseObject(workbook);
                    ReleaseObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}");
            }
        }
        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch
            {
                obj = null;
            }
        }
    }
}