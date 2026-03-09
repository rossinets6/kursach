using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace справочник_начальника_тюрмы
{
    public partial class Vibor : Form
    {
        public Vibor()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Создаем и открываем форму
            Sotrudniki form3 = new Sotrudniki();
            form3.Show();

            // Скрываем текущую форму 
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Zeki form4 = new Zeki();
            form4.Show();

            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Vhid form1 = new Vhid();
            form1.Show();

            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Meroprit form5 = new Meroprit();
            form5.Show();

            this.Hide();
        }
    }
}