using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace справочник_начальника_тюрмы
{
    public partial class NeAdmin : Form
    {
        private DatabaseService dbService = new DatabaseService();

        // Для вкладки "Заключённые"
        private DataTable ZakTable = new DataTable();
        private DataView ZakView;
        private SqlDataAdapter ZakAdapter;

        // Для вкладки "Журнал"
        private DataTable ZyrTable = new DataTable();
        private DataView ZyrView;
        private SqlDataAdapter ZyrAdapter;

        public NeAdmin()
        {
            InitializeComponent();

            // Загрузка данных
            LoadPrisonersData();
            LoadJournalData();

            // Настройка DataGridView (только чтение)
            dataGridView1.ReadOnly = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;

            dataGridView2.ReadOnly = true;
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView2.AllowUserToAddRows = false;
            dataGridView2.AllowUserToDeleteRows = false;

            // Подписка на события сортировки
            radioButton1.CheckedChanged += RadioButton_CheckedChanged;
            radioButton2.CheckedChanged += RadioButton_CheckedChanged;
        }

        // Загрузка данных о заключённых
        private void LoadPrisonersData()
        {
            try
            {
                dbService.openConnection();

                string query = @"SELECT [IDZak], [Имя], [Фамилия], [Отчество], [Дата_рождения],
                                        [Место_рождения], [Номер_дела], [Статья], [Суд],
                                        [Дата_приговора], [Срок], [Дата_заключения], [Дата_освобождения]
                                 FROM [spravochnik].[dbo].[Zaklichenie]
                                 ORDER BY [Фамилия] ASC";

                ZakAdapter = new SqlDataAdapter(query, dbService.getConnection());
                ZakTable.Clear();
                ZakAdapter.Fill(ZakTable);
                ZakView = new DataView(ZakTable);
                dataGridView1.DataSource = ZakView;

                // Настройка заголовков
                ConfigurePrisonersColumns();

                dbService.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки заключённых: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Настройка столбцов для заключённых
        private void ConfigurePrisonersColumns()
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

            // Форматирование дат
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                if (col.Name.Contains("Дата"))
                    col.DefaultCellStyle.Format = "dd.MM.yyyy";
            }
        }

        // Загрузка данных журнала мероприятий
        private void LoadJournalData()
        {
            try
            {
                dbService.openConnection();

                string query = @"SELECT [ID], [Название_мероприятия], [Тип_мероприятия], 
                                        [Дата_проведения], [IDSot], [IDZak]
                                 FROM [spravochnik].[dbo].[Zyrnal]
                                 ORDER BY [Название_мероприятия] ASC";

                ZyrAdapter = new SqlDataAdapter(query, dbService.getConnection());
                ZyrTable.Clear();
                ZyrAdapter.Fill(ZyrTable);
                ZyrView = new DataView(ZyrTable);
                dataGridView2.DataSource = ZyrView;

                // Настройка заголовков
                ConfigureJournalColumns();

                dbService.closeConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки журнала: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Настройка столбцов журнала
        private void ConfigureJournalColumns()
        {
            if (dataGridView2.Columns.Count == 0) return;

            dataGridView2.Columns["ID"].HeaderText = "Код";
            dataGridView2.Columns["Название_мероприятия"].HeaderText = "Название мероприятия";
            dataGridView2.Columns["Тип_мероприятия"].HeaderText = "Тип мероприятия";
            dataGridView2.Columns["Дата_проведения"].HeaderText = "Дата проведения";
            dataGridView2.Columns["IDSot"].HeaderText = "ID сотрудника";
            dataGridView2.Columns["IDZak"].HeaderText = "ID заключённого";

            if (dataGridView2.Columns["Дата_проведения"] != null)
                dataGridView2.Columns["Дата_проведения"].DefaultCellStyle.Format = "dd.MM.yyyy";
        }

        // Сортировка заключённых
        private void RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (ZakView == null) return;

            if (radioButton1.Checked)
                ZakView.Sort = "[Фамилия] ASC";
            else if (radioButton2.Checked)
                ZakView.Sort = "[Фамилия] DESC";
        }

        // Кнопка 1: Поиск по заключённым (фамилия)
        private void button1_Click(object sender, EventArgs e)
        {
            string filterText = textBox1.Text.Trim();
            if (string.IsNullOrEmpty(filterText))
            {
                ZakView.RowFilter = "";
            }
            else
            {
                ZakView.RowFilter = $"[Фамилия] LIKE '%{filterText.Replace("'", "''")}%'";
            }

            if (ZakView.Count == 0)
                MessageBox.Show("Заключённые не найдены", "Поиск", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Кнопка 2: Назад (закрыть и открыть Vhid)
        private void button2_Click(object sender, EventArgs e)
        {
            ReturnToVhid();
        }

        // Кнопка 3: Поиск по журналу (название мероприятия)
        private void button3_Click(object sender, EventArgs e)
        {
            string filterText = textBox2.Text.Trim();
            if (string.IsNullOrEmpty(filterText))
            {
                ZyrView.RowFilter = "";
            }
            else
            {
                ZyrView.RowFilter = $"[Название_мероприятия] LIKE '%{filterText.Replace("'", "''")}%'";
            }

            if (ZyrView.Count == 0)
                MessageBox.Show("Мероприятия не найдены", "Поиск", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Кнопка 4: Подробно (информация о сотруднике и заключённом)
        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите строку в журнале", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = dataGridView2.SelectedRows[0];
            if (row.IsNewRow) return;

            // Получаем ID
            object idSotObj = row.Cells["IDSot"].Value;
            object idZakObj = row.Cells["IDZak"].Value;

            if (idSotObj == null || idSotObj == DBNull.Value || idZakObj == null || idZakObj == DBNull.Value)
            {
                MessageBox.Show("В выбранной записи отсутствуют ID сотрудника или заключённого", "Внимание",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int idSot = Convert.ToInt32(idSotObj);
            int idZak = Convert.ToInt32(idZakObj);

            string empInfo = GetEmployeeInfo(idSot);
            string prisInfo = GetPrisonerInfo(idZak);

            string message = $"СОТРУДНИК (ID: {idSot})\n{empInfo}\n\nЗАКЛЮЧЁННЫЙ (ID: {idZak})\n{prisInfo}";
            MessageBox.Show(message, "Подробная информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Кнопка 5: Назад (аналогично button2)
        private void button5_Click(object sender, EventArgs e)
        {
            ReturnToVhid();
        }

        // Возврат на форму входа
        private void ReturnToVhid()
        {
            Vhid vhid = new Vhid();
            vhid.Show();
            this.Close();
        }

        // Метод получения информации о сотруднике (аналогично предыдущим)
        private string GetEmployeeInfo(int employeeId)
        {
            string result = "Информация не найдена";
            try
            {
                dbService.openConnection();
                string query = "SELECT [Имя], [Фамилия], [Отчество], [Должность], [Номер], [Адрес] " +
                              "FROM [spravochnik].[dbo].[Sotrudniki] WHERE [IDSot] = @id";
                using (SqlCommand cmd = new SqlCommand(query, dbService.getConnection()))
                {
                    cmd.Parameters.AddWithValue("@id", employeeId);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string name = reader["Имя"]?.ToString() ?? "не указано";
                            string surname = reader["Фамилия"]?.ToString() ?? "не указано";
                            string patronymic = reader["Отчество"]?.ToString() ?? "не указано";
                            string position = reader["Должность"]?.ToString() ?? "не указана";
                            string phone = reader["Номер"]?.ToString() ?? "не указан";
                            string address = reader["Адрес"]?.ToString() ?? "не указан";
                            result = $"ФИО: {surname} {name} {patronymic}\nДолжность: {position}\nТелефон: {phone}\nАдрес: {address}";
                        }
                        else result = $"Сотрудник с ID {employeeId} не найден";
                    }
                }
                dbService.closeConnection();
            }
            catch (Exception ex)
            {
                result = $"Ошибка: {ex.Message}";
            }
            return result;
        }

        // Метод получения информации о заключённом
        private string GetPrisonerInfo(int prisonerId)
        {
            string result = "Информация не найдена";
            try
            {
                dbService.openConnection();
                string query = @"SELECT [Имя], [Фамилия], [Отчество], [Дата_рождения], [Статья], [Срок],
                                        [Место_рождения], [Номер_дела], [Суд], [Дата_приговора], 
                                        [Дата_заключения], [Дата_освобождения]
                                 FROM [spravochnik].[dbo].[Zaklichenie] WHERE [IDZak] = @id";
                using (SqlCommand cmd = new SqlCommand(query, dbService.getConnection()))
                {
                    cmd.Parameters.AddWithValue("@id", prisonerId);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string name = reader["Имя"]?.ToString() ?? "не указано";
                            string surname = reader["Фамилия"]?.ToString() ?? "не указано";
                            string patronymic = reader["Отчество"]?.ToString() ?? "не указано";
                            string birthDate = reader["Дата_рождения"] != DBNull.Value ? Convert.ToDateTime(reader["Дата_рождения"]).ToShortDateString() : "не указана";
                            string article = reader["Статья"]?.ToString() ?? "не указана";
                            string term = reader["Срок"]?.ToString() ?? "не указан";
                            string place = reader["Место_рождения"]?.ToString() ?? "не указано";
                            string caseNum = reader["Номер_дела"]?.ToString() ?? "не указан";
                            string court = reader["Суд"]?.ToString() ?? "не указан";
                            string verdictDate = reader["Дата_приговора"] != DBNull.Value ? Convert.ToDateTime(reader["Дата_приговора"]).ToShortDateString() : "не указана";
                            string imprisonmentDate = reader["Дата_заключения"] != DBNull.Value ? Convert.ToDateTime(reader["Дата_заключения"]).ToShortDateString() : "не указана";
                            string releaseDate = reader["Дата_освобождения"] != DBNull.Value ? Convert.ToDateTime(reader["Дата_освобождения"]).ToShortDateString() : "не указана";

                            result = $"ФИО: {surname} {name} {patronymic}\n" +
                                     $"Дата рождения: {birthDate}\n" +
                                     $"Место рождения: {place}\n" +
                                     $"Номер дела: {caseNum}\n" +
                                     $"Статья: {article}\n" +
                                     $"Суд: {court}\n" +
                                     $"Дата приговора: {verdictDate}\n" +
                                     $"Срок: {term}\n" +
                                     $"Дата заключения: {imprisonmentDate}\n" +
                                     $"Дата освобождения: {releaseDate}";
                        }
                        else result = $"Заключённый с ID {prisonerId} не найден";
                    }
                }
                dbService.closeConnection();
            }
            catch (Exception ex)
            {
                result = $"Ошибка: {ex.Message}";
            }
            return result;
        }

        // Обработка ошибок DataGridView
        private void dgvPrisoners_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Ошибка отображения данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.ThrowException = false;
        }

        private void dgvJournal_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Ошибка отображения данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.ThrowException = false;
        }
    }
}