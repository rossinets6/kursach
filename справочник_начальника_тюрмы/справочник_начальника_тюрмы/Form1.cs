using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace справочник_начальника_тюрмы
{
    public partial class Vhid : Form
    {
        // Создаем экземпляр вашего класса для работы с БД
        private DatabaseService dbService = new DatabaseService();

        public Vhid()
        {
            InitializeComponent();

            // Дополнительные настройки
            textBox2.PasswordChar = '*'; // Скрываем ввод пароля
            textBox1.MaxLength = 50; // Ограничение длины логина
            textBox2.MaxLength = 50; // Ограничение длины пароля
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Проверяем, что поля не пустые
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Введите логин и пароль!", "Ошибка входа",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string login = textBox1.Text.Trim();
            string password = textBox2.Text;

            // Проверяем логин и пароль через БД
            if (CheckUserInDatabase(login, password))
            {
                // Проверяем, является ли пользователь админом (admin/1)
                if (login == "admin" && password == "1")
                {
                    // Админ - открываем Vibor
                    Vibor form = new Vibor();
                    form.Show();
                }
                else
                {
                    // Обычный пользователь - открываем NeAdmin
                    NeAdmin form = new NeAdmin();
                    form.Show();
                }
                this.Hide();
            }
            else
            {
                // Если логин или пароль неверные, показываем сообщение
                MessageBox.Show("Неверный логин или пароль!", "Ошибка входа",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Очищаем поле пароля для повторного ввода
                textBox2.Clear();
                textBox2.Focus();
            }
        }

        private bool CheckUserInDatabase(string login, string password)
        {
            bool isValid = false;

            try
            {
                // Открываем соединение с БД
                dbService.openConnection();

                // SQL запрос для проверки логина и пароля
                // ВНИМАНИЕ: В реальном проекте пароли должны храниться в захешированном виде!
                string query = "SELECT COUNT(*) FROM sotrudniki WHERE Логин = @login AND Пароль = @password";

                // Альтернативный вариант, если в таблице используются другие названия полей:
                // string query = "SELECT COUNT(*) FROM sotrudniki WHERE Логин = @login AND Пароль = @password";

                SqlCommand command = new SqlCommand(query, dbService.getConnection());
                command.Parameters.AddWithValue("@login", login);
                command.Parameters.AddWithValue("@password", password); // В реальности - хеш пароля

                int count = (int)command.ExecuteScalar();

                isValid = count > 0; // Если найден хотя бы один пользователь
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения к базе данных: {ex.Message}",
                              "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем соединение с БД
                dbService.closeConnection();
            }

            return isValid;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Закрываем программу
            Application.Exit();
        }

        // Обработчик нажатия Enter в полях ввода
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                button1.PerformClick();
            }
        }
    }
}