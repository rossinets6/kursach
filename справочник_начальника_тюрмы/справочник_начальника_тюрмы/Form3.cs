using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace справочник_начальника_тюрмы
{
    public partial class Sotrudniki : Form
    {
        private DatabaseService dbService = new DatabaseService();
        private DataTable dataTable = new DataTable();
        private SqlDataAdapter dataAdapter;
        private string currentFilter = "";
        private string currentSort = "ASC";

        public Sotrudniki()
        {
            InitializeComponent();
            LoadData();

            // Настройка DataGridView
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.ReadOnly = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;

            // Подписка на события
            radioButton1.CheckedChanged += RadioButton_CheckedChanged;
            radioButton2.CheckedChanged += RadioButton_CheckedChanged;
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
            dataGridView1.CellBeginEdit += dataGridView1_CellBeginEdit;
            dataGridView1.CellDoubleClick += dataGridView1_CellDoubleClick;
            textBox1.KeyPress += textBox1_KeyPress;
        }

        // Загрузка данных с фильтром и сортировкой
        private void LoadData(string filter = "", string sortOrder = "ASC")
        {
            try
            {
                dbService.openConnection();

                string query = "SELECT * FROM sotrudniki";

                if (!string.IsNullOrWhiteSpace(filter))
                    query += " WHERE Фамилия LIKE @filter";

                query += " ORDER BY Фамилия " + sortOrder;

                dataAdapter = new SqlDataAdapter(query, dbService.getConnection());

                // CommandBuilder для автоматических INSERT/UPDATE/DELETE
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);

                if (!string.IsNullOrWhiteSpace(filter))
                    dataAdapter.SelectCommand.Parameters.AddWithValue("@filter", $"%{filter}%");

                dataTable.Clear();
                dataAdapter.Fill(dataTable);
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

        // Настройка заголовков и видимости столбцов
        private void ConfigureColumnHeaders()
        {
            if (dataGridView1.Columns.Count == 0) return;

            // ID (может называться ID или Код)
            if (dataGridView1.Columns.Contains("IDSot"))
                dataGridView1.Columns["IDSot"].HeaderText = "Код";
            else if (dataGridView1.Columns.Contains("Код"))
                dataGridView1.Columns["Код"].HeaderText = "Код";

            if (dataGridView1.Columns.Contains("Фамилия"))
                dataGridView1.Columns["Фамилия"].HeaderText = "Фамилия";
            if (dataGridView1.Columns.Contains("Имя"))
                dataGridView1.Columns["Имя"].HeaderText = "Имя";
            if (dataGridView1.Columns.Contains("Отчество"))
                dataGridView1.Columns["Отчество"].HeaderText = "Отчество";
            if (dataGridView1.Columns.Contains("Дата_рождения"))
                dataGridView1.Columns["Дата_рождения"].HeaderText = "Дата рождения";
            if (dataGridView1.Columns.Contains("Возраст"))
                dataGridView1.Columns["Возраст"].HeaderText = "Возраст";
            if (dataGridView1.Columns.Contains("Должность"))
                dataGridView1.Columns["Должность"].HeaderText = "Должность";
            if (dataGridView1.Columns.Contains("Номер"))
                dataGridView1.Columns["Номер"].HeaderText = "Номер";
            if (dataGridView1.Columns.Contains("Адрес"))
                dataGridView1.Columns["Адрес"].HeaderText = "Адрес";
            if (dataGridView1.Columns.Contains("Логин"))
                dataGridView1.Columns["Логин"].HeaderText = "Логин";
            if (dataGridView1.Columns.Contains("Пароль"))
            {
                dataGridView1.Columns["Пароль"].HeaderText = "Пароль";
                dataGridView1.Columns["Пароль"].Visible = false; // скрываем
            }

            // Делаем ID только для чтения
            if (dataGridView1.Columns.Contains("IDSot"))
                dataGridView1.Columns["IDSot"].ReadOnly = true;
            else if (dataGridView1.Columns.Contains("Код"))
                dataGridView1.Columns["Код"].ReadOnly = true;
        }

        // Сохранение изменений в БД
        private void SaveCurrentChanges()
        {
            try
            {
                if (dataGridView1.IsCurrentCellInEditMode)
                    dataGridView1.EndEdit();

                if (dataTable?.GetChanges() != null && dataAdapter != null)
                {
                    dataAdapter.Update(dataTable);
                    // Сброс подсветки изменённых строк
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                        row.DefaultCellStyle.BackColor = Color.White;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Обработчик переключения сортировки
        private void RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            string sortOrder = radioButton1.Checked ? "ASC" : "DESC";
            LoadData(currentFilter, sortOrder);
        }

        // Добавить сотрудника
        private void button1_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges(); // опционально
            Dobav dobavForm = new Dobav();
            dobavForm.ShowDialog();
            LoadData(currentFilter, currentSort); // обновить после добавления
        }

        // Сохранить изменения
        private void button2_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            MessageBox.Show("Изменения сохранены", "Успех",
                          MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Удалить сотрудника
        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите сотрудника для удаления", "Внимание",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string idColumn = dataGridView1.Columns.Contains("IDSot") ? "IDSot" : "Код";
            if (!dataGridView1.Columns.Contains(idColumn))
            {
                MessageBox.Show("Не найден идентификатор", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int id = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[idColumn].Value);
            string fam = dataGridView1.SelectedRows[0].Cells["Фамилия"].Value?.ToString() ?? "";
            string name = dataGridView1.SelectedRows[0].Cells["Имя"].Value?.ToString() ?? "";

            if (MessageBox.Show($"Удалить {fam} {name}?", "Подтверждение",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    dbService.openConnection();
                    string query = $"DELETE FROM sotrudniki WHERE {idColumn} = @idsot";
                    SqlCommand cmd = new SqlCommand(query, dbService.getConnection());
                    cmd.Parameters.AddWithValue("@idsot", id);
                    int rows = cmd.ExecuteNonQuery();

                    if (rows > 0)
                    {
                        MessageBox.Show("Сотрудник удалён", "Успех",
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

        // Поиск по фамилии
        private void button4_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            LoadData(textBox1.Text.Trim(), currentSort);
            if (dataTable.Rows.Count == 0)
                MessageBox.Show("Ничего не найдено", "Поиск",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Возврат на форму выбора
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

        // Enter в поле поиска
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                button4.PerformClick();
                e.Handled = true;
            }
        }

        // Подсветка изменённых ячеек
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
        }

        // Запрет редактирования ID
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            string idColumn = dataGridView1.Columns.Contains("IDSot") ? "IDSot" : "Код";
            if (dataGridView1.Columns[e.ColumnIndex].Name == idColumn)
            {
                e.Cancel = true;
                MessageBox.Show("Это поле нельзя редактировать", "Внимание",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Двойной клик – начать редактирование (если не ID)
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            string idColumn = dataGridView1.Columns.Contains("IDSot") ? "IDSot" : "Код";
            if (dataGridView1.Columns[e.ColumnIndex].Name != idColumn)
                dataGridView1.BeginEdit(true);
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
                saveFileDialog.FileName = "Сотрудники" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

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