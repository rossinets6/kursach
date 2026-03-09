using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace справочник_начальника_тюрмы
{
    public partial class DobMer : Form
    {
        private DatabaseService dbService = new DatabaseService();

        public DobMer()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Проверка обязательного поля "Название мероприятия"
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("Введите название мероприятия!", "Предупреждение",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                dbService.openConnection();

                string query = @"INSERT INTO [Zyrnal]
                                ([Название_мероприятия], [Тип_мероприятия], [Дата_проведения], [IDSot], [IDZak])
                         VALUES
                                (@name, @type, @date, @idSot, @idZak)";

                SqlCommand cmd = new SqlCommand(query, dbService.getConnection());

                cmd.Parameters.AddWithValue("@name", textBox1.Text.Trim());
                cmd.Parameters.AddWithValue("@type", textBox2.Text.Trim());

                // Дата проведения
                if (DateTime.TryParse(textBox3.Text.Trim(), out DateTime eventDate))
                    cmd.Parameters.AddWithValue("@date", eventDate);
                else
                    cmd.Parameters.AddWithValue("@date", DBNull.Value);

                // ID сотрудника (int, может быть NULL)
                if (int.TryParse(textBox4.Text.Trim(), out int idSot))
                    cmd.Parameters.AddWithValue("@idSot", idSot);
                else
                    cmd.Parameters.AddWithValue("@idSot", DBNull.Value);

                // ID заключённого (int, может быть NULL)
                if (int.TryParse(textBox5.Text.Trim(), out int idZak))
                    cmd.Parameters.AddWithValue("@idZak", idZak);
                else
                    cmd.Parameters.AddWithValue("@idZak", DBNull.Value);

                int rows = cmd.ExecuteNonQuery();

                if (rows > 0)
                {
                    MessageBox.Show("Мероприятие успешно добавлено!", "Успех",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Не удалось добавить мероприятие.", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                dbService.closeConnection();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}