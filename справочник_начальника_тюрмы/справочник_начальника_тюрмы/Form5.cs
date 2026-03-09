using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace справочник_начальника_тюрмы
{
    public partial class Meroprit : Form
    {
        private DatabaseService dbService = new DatabaseService();
        private DataTable dataTable = new DataTable();
        private SqlDataAdapter dataAdapter;
        private string currentFilter = "";

        public Meroprit()
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
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
            dataGridView1.CellBeginEdit += dataGridView1_CellBeginEdit;
            dataGridView1.CellDoubleClick += dataGridView1_CellDoubleClick;
            textBox1.KeyPress += textBox1_KeyPress;
        }

        // Загрузка данных с фильтром
        private void LoadData(string filter = "")
        {
            try
            {
                dbService.openConnection();

                string query = @"SELECT [ID], [Название_мероприятия], [Тип_мероприятия], 
                                [Дата_проведения], [IDSot], [IDZak]
                         FROM [spravochnik].[dbo].[Zyrnal]";

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    query += " WHERE [Название_мероприятия] LIKE @filter";
                }

                query += " ORDER BY [Название_мероприятия] ASC";

                dataAdapter = new SqlDataAdapter(query, dbService.getConnection());

                // Добавляем параметр ДО создания CommandBuilder
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    dataAdapter.SelectCommand.Parameters.AddWithValue("@filter", $"%{filter}%");
                }

                // Создаём CommandBuilder для автоматических команд обновления
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                dataAdapter.InsertCommand = commandBuilder.GetInsertCommand();
                dataAdapter.UpdateCommand = commandBuilder.GetUpdateCommand();
                dataAdapter.DeleteCommand = commandBuilder.GetDeleteCommand();

                dataTable.Clear();
                dataAdapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;

                // Не очищаем параметры – они больше не нужны, а адаптер будет заменён при следующей загрузке
                ConfigureColumnHeaders();

                currentFilter = filter;
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

        // Настройка заголовков и форматирования
        private void ConfigureColumnHeaders()
        {
            if (dataGridView1.Columns.Count == 0) return;

            dataGridView1.Columns["ID"].HeaderText = "Код";
            dataGridView1.Columns["Название_мероприятия"].HeaderText = "Название мероприятия";
            dataGridView1.Columns["Тип_мероприятия"].HeaderText = "Тип мероприятия";
            dataGridView1.Columns["Дата_проведения"].HeaderText = "Дата проведения";
            dataGridView1.Columns["IDSot"].HeaderText = "ID сотрудника";
            dataGridView1.Columns["IDZak"].HeaderText = "ID заключенного";

            // ID только для чтения
            dataGridView1.Columns["ID"].ReadOnly = true;

            // Форматирование даты
            if (dataGridView1.Columns["Дата_проведения"] != null)
            {
                dataGridView1.Columns["Дата_проведения"].DefaultCellStyle.Format = "dd.MM.yyyy";
            }
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

        // Добавить мероприятие
        private void button1_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges(); // опционально
            DobMer dobMerForm = new DobMer();
            if (dobMerForm.ShowDialog() == DialogResult.OK)
            {
                LoadData(currentFilter);
            }
        }

        // Сохранить изменения
        private void button2_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            MessageBox.Show("Изменения сохранены", "Успех",
                          MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Удалить мероприятие
        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите запись для удаления", "Внимание",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int id = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["ID"].Value);
            string name = dataGridView1.SelectedRows[0].Cells["Название_мероприятия"].Value?.ToString() ?? "";

            if (MessageBox.Show($"Удалить мероприятие \"{name}\"?", "Подтверждение",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    dbService.openConnection();
                    string query = "DELETE FROM [Zyrnal] WHERE ID = @id";
                    SqlCommand cmd = new SqlCommand(query, dbService.getConnection());
                    cmd.Parameters.AddWithValue("@id", id);
                    int rows = cmd.ExecuteNonQuery();

                    if (rows > 0)
                    {
                        MessageBox.Show("Запись удалена", "Успех",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadData(currentFilter);
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

        // Поиск по названию мероприятия
        private void button4_Click(object sender, EventArgs e)
        {
            SaveCurrentChanges();
            LoadData(textBox1.Text.Trim());
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

        // Кнопка "Подробно" – информация о сотруднике и заключённом
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите строку с мероприятием", "Внимание",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];

                if (selectedRow.IsNewRow)
                {
                    MessageBox.Show("Нельзя получить данные для новой строки.", "Внимание",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Проверяем наличие ID сотрудника и заключённого
                if (selectedRow.Cells["IDSot"].Value == null ||
                    selectedRow.Cells["IDSot"].Value == DBNull.Value ||
                    selectedRow.Cells["IDZak"].Value == null ||
                    selectedRow.Cells["IDZak"].Value == DBNull.Value)
                {
                    MessageBox.Show("В выбранной строке не указаны ID сотрудника или ID заключенного.",
                                  "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int idSot = Convert.ToInt32(selectedRow.Cells["IDSot"].Value);
                int idZak = Convert.ToInt32(selectedRow.Cells["IDZak"].Value);

                string employeeInfo = GetEmployeeInfo(idSot);
                string prisonerInfo = GetPrisonerInfo(idZak);

                string message = $"ИНФОРМАЦИЯ О СОТРУДНИКЕ (ID: {idSot})\n" +
                                 $"===============================\n" +
                                 $"{employeeInfo}\n\n" +
                                 $"ИНФОРМАЦИЯ О ЗАКЛЮЧЕННОМ (ID: {idZak})\n" +
                                 $"================================\n" +
                                 $"{prisonerInfo}";

                MessageBox.Show(message, "Данные о сотруднике и заключенном",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении данных: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод для получения информации о сотруднике по айди
        private string GetEmployeeInfo(int employeeId)
        {
            string result = "Информация не найдена";

            try
            {
                dbService.openConnection();

                string query = "SELECT [Имя], [Фамилия], [Отчество], [Должность], [Номер], [Адрес] " +
                              "FROM [spravochnik].[dbo].[Sotrudniki] WHERE [IDSot] = @idsot";

                using (SqlCommand command = new SqlCommand(query, dbService.getConnection()))
                {
                    command.Parameters.AddWithValue("@idsot", employeeId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string name = reader["Имя"] != DBNull.Value ? reader["Имя"].ToString() : "не указано";
                            string surname = reader["Фамилия"] != DBNull.Value ? reader["Фамилия"].ToString() : "не указано";
                            string patronymic = reader["Отчество"] != DBNull.Value ? reader["Отчество"].ToString() : "не указано";
                            string position = reader["Должность"] != DBNull.Value ? reader["Должность"].ToString() : "не указана";
                            string phone = reader["Номер"] != DBNull.Value ? reader["Номер"].ToString() : "не указан";
                            string address = reader["Адрес"] != DBNull.Value ? reader["Адрес"].ToString() : "не указан";

                            result = $"ФИО: {surname} {name} {patronymic}\n" +
                                    $"Должность: {position}\n" +
                                    $"Телефон: {phone}\n" +
                                    $"Адрес: {address}";
                        }
                        else
                        {
                            result = $"Сотрудник с ID {employeeId} не найден в базе данных.";
                        }
                    }
                }

                dbService.closeConnection();
            }
            catch (Exception ex)
            {
                result = $"Ошибка при получении данных о сотруднике: {ex.Message}";
            }

            return result; // гарантированный возврат
        }

        // Метод для получения информации о заключенном по айди
        private string GetPrisonerInfo(int prisonerId)
        {
            string result = "Информация не найдена";

            try
            {
                dbService.openConnection();

                string query = "SELECT [Имя], [Фамилия], [Отчество], [Дата_рождения], [Статья], [Срок], " +
                              "[Место_рождения], [Номер_дела], [Суд], [Дата_приговора], [Дата_заключения], [Дата_освобождения] " +
                              "FROM [spravochnik].[dbo].[Zaklichenie] WHERE [IDZak] = @idzak";

                using (SqlCommand command = new SqlCommand(query, dbService.getConnection()))
                {
                    command.Parameters.AddWithValue("@idzak", prisonerId);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string name = reader["Имя"] != DBNull.Value ? reader["Имя"].ToString() : "не указано";
                            string surname = reader["Фамилия"] != DBNull.Value ? reader["Фамилия"].ToString() : "не указано";
                            string patronymic = reader["Отчество"] != DBNull.Value ? reader["Отчество"].ToString() : "не указано";

                            string birthDate = "не указана";
                            if (reader["Дата_рождения"] != DBNull.Value)
                                birthDate = Convert.ToDateTime(reader["Дата_рождения"]).ToShortDateString();

                            string article = reader["Статья"] != DBNull.Value ? reader["Статья"].ToString() : "не указана";
                            string term = reader["Срок"] != DBNull.Value ? reader["Срок"].ToString() : "не указан";
                            string placeOfBirth = reader["Место_рождения"] != DBNull.Value ? reader["Место_рождения"].ToString() : "не указано";
                            string caseNumber = reader["Номер_дела"] != DBNull.Value ? reader["Номер_дела"].ToString() : "не указан";
                            string court = reader["Суд"] != DBNull.Value ? reader["Суд"].ToString() : "не указан";

                            string verdictDate = "не указана";
                            if (reader["Дата_приговора"] != DBNull.Value)
                                verdictDate = Convert.ToDateTime(reader["Дата_приговора"]).ToShortDateString();

                            string imprisonmentDate = "не указана";
                            if (reader["Дата_заключения"] != DBNull.Value)
                                imprisonmentDate = Convert.ToDateTime(reader["Дата_заключения"]).ToShortDateString();

                            string releaseDate = "не указана";
                            if (reader["Дата_освобождения"] != DBNull.Value)
                                releaseDate = Convert.ToDateTime(reader["Дата_освобождения"]).ToShortDateString();

                            result = $"ФИО: {surname} {name} {patronymic}\n" +
                                    $"Дата рождения: {birthDate}\n" +
                                    $"Место рождения: {placeOfBirth}\n" +
                                    $"Номер дела: {caseNumber}\n" +
                                    $"Статья: {article}\n" +
                                    $"Суд: {court}\n" +
                                    $"Дата приговора: {verdictDate}\n" +
                                    $"Срок: {term}\n" +
                                    $"Дата заключения: {imprisonmentDate}\n" +
                                    $"Дата освобождения: {releaseDate}";
                        }
                        else
                        {
                            result = $"Заключенный с ID {prisonerId} не найден в базе данных.";
                        }
                    }
                }

                dbService.closeConnection();
            }
            catch (Exception ex)
            {
                result = $"Ошибка при получении данных о заключенном: {ex.Message}";
            }

            return result; // гарантированный возврат
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
            if (dataGridView1.Columns[e.ColumnIndex].Name == "ID")
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
            if (dataGridView1.Columns[e.ColumnIndex].Name != "ID")
                dataGridView1.BeginEdit(true);
        }

        // Валидация ввода (опционально, можно добавить для даты)
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "Дата_проведения")
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

        private void button7_Click(object sender, EventArgs e)
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
                saveFileDialog.FileName = "Мероприятия" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

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