using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace справочник_начальника_тюрмы
{
    public partial class DobZak : Form
    {
        private DatabaseService dbService = new DatabaseService();

        public DobZak()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Проверка обязательных полей (Фамилия и Имя)
            if (string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("Заполните фамилию и имя!", "Предупреждение",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                dbService.openConnection();

                string query = @"INSERT INTO [Zaklichenie]
                                ([Имя], [Фамилия], [Отчество], [Дата_рождения], [Место_рождения],
                                 [Номер_дела], [Статья], [Суд], [Дата_приговора], [Срок],
                                 [Дата_заключения], [Дата_освобождения])
                         VALUES
                                (@imya, @fam, @otch, @date_birth, @place_birth,
                                 @nom_dela, @statya, @sud, @date_prigovor, @srok,
                                 @date_zakl, @date_osvob)";

                SqlCommand cmd = new SqlCommand(query, dbService.getConnection());

                // Заполнение параметров
                cmd.Parameters.AddWithValue("@imya", textBox1.Text.Trim());
                cmd.Parameters.AddWithValue("@fam", textBox2.Text.Trim());
                cmd.Parameters.AddWithValue("@otch", textBox3.Text.Trim());

                // Дата рождения
                if (DateTime.TryParse(textBox4.Text.Trim(), out DateTime birthDate))
                    cmd.Parameters.AddWithValue("@date_birth", birthDate);
                else
                    cmd.Parameters.AddWithValue("@date_birth", DBNull.Value);

                cmd.Parameters.AddWithValue("@place_birth", textBox5.Text.Trim());
                cmd.Parameters.AddWithValue("@nom_dela", textBox6.Text.Trim());
                cmd.Parameters.AddWithValue("@statya", textBox7.Text.Trim());
                cmd.Parameters.AddWithValue("@sud", textBox8.Text.Trim());

                // Дата приговора
                if (DateTime.TryParse(textBox9.Text.Trim(), out DateTime prigovorDate))
                    cmd.Parameters.AddWithValue("@date_prigovor", prigovorDate);
                else
                    cmd.Parameters.AddWithValue("@date_prigovor", DBNull.Value);

                // Срок (int)
                if (int.TryParse(textBox10.Text.Trim(), out int srok))
                    cmd.Parameters.AddWithValue("@srok", srok);
                else
                    cmd.Parameters.AddWithValue("@srok", DBNull.Value);

                // Дата заключения
                if (DateTime.TryParse(textBox11.Text.Trim(), out DateTime zaklDate))
                    cmd.Parameters.AddWithValue("@date_zakl", zaklDate);
                else
                    cmd.Parameters.AddWithValue("@date_zakl", DBNull.Value);

                // Дата освобождения
                if (DateTime.TryParse(textBox12.Text.Trim(), out DateTime osvobDate))
                    cmd.Parameters.AddWithValue("@date_osvob", osvobDate);
                else
                    cmd.Parameters.AddWithValue("@date_osvob", DBNull.Value);

                int rows = cmd.ExecuteNonQuery();

                if (rows > 0)
                {
                    MessageBox.Show("Запись успешно добавлена!", "Успех",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Не удалось добавить запись.", "Ошибка",
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

        private void DobZak_Load(object sender, EventArgs e)
        {

        }
    }
}