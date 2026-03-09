using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace справочник_начальника_тюрмы
{
    public partial class Dobav : Form
    {
        private DatabaseService dbService = new DatabaseService();

        public Dobav()
        {
            InitializeComponent();
        }

        // Кнопка "Добавить"
        private void button1_Click(object sender, EventArgs e)
        {
            // Проверяем заполнение обязательных полей (например, Фамилия и Имя)
            if (string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("Заполните фамилию и имя!", "Предупреждение",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                dbService.openConnection();

                // Запрос на вставку (ID автоинкремент, поэтому его не указываем)
                string query = @"INSERT INTO sotrudniki 
                                (Имя, Фамилия, Отчество, Дата_рождения, Возраст, Должность, Номер, Адрес, Логин, Пароль) 
                         VALUES 
                                (@imya, @fam, @otch, @date, @age, @dolj, @nom, @adr, @login, @pass)";

                SqlCommand cmd = new SqlCommand(query, dbService.getConnection());

                // Параметры (соответствие: textBox1=Имя, textBox2=Фамилия, textBox3=Отчество, ...)
                cmd.Parameters.AddWithValue("@imya", textBox1.Text.Trim());
                cmd.Parameters.AddWithValue("@fam", textBox2.Text.Trim());
                cmd.Parameters.AddWithValue("@otch", textBox3.Text.Trim());

                // Дата рождения (textBox4): преобразуем в DateTime, если возможно
                if (DateTime.TryParse(textBox4.Text.Trim(), out DateTime birthDate))
                    cmd.Parameters.AddWithValue("@date", birthDate);
                else
                    cmd.Parameters.AddWithValue("@date", DBNull.Value);

                // Возраст (textBox5): пытаемся преобразовать в int
                if (int.TryParse(textBox5.Text.Trim(), out int age))
                    cmd.Parameters.AddWithValue("@age", age);
                else
                    cmd.Parameters.AddWithValue("@age", DBNull.Value);

                cmd.Parameters.AddWithValue("@dolj", textBox6.Text.Trim());   // Должность
                cmd.Parameters.AddWithValue("@nom", textBox7.Text.Trim());    // Номер
                cmd.Parameters.AddWithValue("@adr", textBox8.Text.Trim());    // Адрес
                cmd.Parameters.AddWithValue("@login", textBox9.Text.Trim());  // Логин
                cmd.Parameters.AddWithValue("@pass", textBox10.Text.Trim());  // Пароль

                int rows = cmd.ExecuteNonQuery();

                if (rows > 0)
                {
                    MessageBox.Show("Сотрудник успешно добавлен!", "Успех",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Не удалось добавить сотрудника.", "Ошибка",
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

        // Кнопка "Отмена" — просто закрыть форму
        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}