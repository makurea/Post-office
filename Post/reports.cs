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

namespace Post
{
    public partial class reports : Form
    {
        public reports()
        {
            InitializeComponent();
            //Занесение данных из таблиц в комбобоксы, для того чтобы пользователь мог просто выбрать значение            
            forComboBox("name", "newspaper", comboBox1);
            forComboBox("nameT", "postOffice", comboBox2);
            forComboBox("name", "newspaper", comboBox3);
            forComboBox("adrT","postOffice", comboBox4);
            forComboBox("name", "newspaper", comboBox5);

        }

        //Функция которая выполняет запрос на основании полученной строки
        void query(String sqlStr)
        {
            SqlDataAdapter da;
            da = new SqlDataAdapter(sqlStr, Ress.connect);
            DataSet ds = new DataSet();
            da.Fill(ds, "orders");
            dataGridView1.DataSource = ds.Tables["orders"];
            dataGridView1.Refresh();
            dataGridView1.ClearSelection();
        }
        private void label3_Click(object sender, EventArgs e)
        {

        }
        

        // Функция которая обрабатывает нажатие кнопки для выполнения первого запроса
        private void button1_Click(object sender, EventArgs e)
        {

            query("select distinct name as 'Газета', adrT as 'Адрес' from newspaper, postOffice, orders" +
                 " where postOffice.idT=orders.idT and newspaper.idn=orders.idN and newspaper.name ='" + comboBox1.Text + "'");
           
        }

        //Данная функция является копией оденоименной функции на форме добавления/изменения
        private void forComboBox(String field, String table, ComboBox cb)
        {


            SqlDataAdapter da = new SqlDataAdapter("Select " + field + " from " + table, Ress.connect);
            DataTable dataTable = new DataTable();
            DataSet ds = new DataSet();
            da.Fill(ds, table);
            for (int i = 0; i < ds.Tables[table].Rows.Count; i++)
            {
                cb.Items.Add(ds.Tables[table].Rows[i][0].ToString());
            }

        }

        //Обработка кнопки "Назад". Возвращает пользователя на главную форму
        private void button9_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 f = new Form1();
            f.ShowDialog();
        }

        // Функция которая обрабатывает нажатие кнопки для выполнения второго запроса
        private void button2_Click(object sender, EventArgs e)
        {
            query(" select top 1 name as 'Газета', FIO as 'ФИО', sum(count) as 'Количество' from newspaper, postOffice, orders" +
                 " where postOffice.idT=orders.idT and newspaper.idn=orders.idN and postOffice.nameT ='" + comboBox2.Text + "'" +
                 "group by name,FIO  order by Количество desc ");
    
        }

        // Функция которая обрабатывает нажатие кнопки для выполнения третьего запроса
        private void button3_Click(object sender, EventArgs e)
        {

            query("select distinct name as 'Газета',  price as 'Цена', adr as 'Адрес' from newspaper, mail, orders" +
                  " where mail.idp=orders.idP and newspaper.idn=orders.idN and newspaper.price>" + textBox1.Text);
           
        }

        // Функция которая обрабатывает нажатие кнопки для выполнения четвертого запроса
        private void button4_Click(object sender, EventArgs e)
        {

            query("select distinct name as 'Газета', orders.count as Количество, idP as 'Номер почтового отделения' from newspaper, orders" +
                 " where  newspaper.idn=orders.idN and count<" + textBox2.Text);
          
        }


        // Функция которая обрабатывает нажатие кнопки для выполнения пятого запроса
        private void button5_Click(object sender, EventArgs e)
        {

            query("Select distinct (Select name from newspaper where newspaper.idn=orders.idN) as 'Газета', mail.idp as 'Номер почтового отделения', adr as 'Адрес' from orders, mail" +
                " where orders.idP=mail.idp and idN=" + textToId(comboBox3.Text, "name", "newspaper") + " and idT=" + textToId(comboBox4.Text, "adrT", "postOffice"));
          
        }


        //Данная функция является копией оденоименной функции на форме добавления/изменения
        private String textToId(String data, String field, String table)
        {
            String id1 = "";
            switch (table)
            {
                case "newspaper": id1 = "idn"; break;
                case "postOffice": id1 = "idT"; break;
                case "mail": id1 = "idp"; break;
                case "ordersView": id1 = "idZ"; break;
                default: break;

            }
            SqlDataAdapter da = new SqlDataAdapter("Select " + id1 + " from " + table + " where " + field + " = '" + data + "'", Ress.connect);
            DataTable dt = new DataTable();
            da.Fill(dt);
            String s1 = dt.Rows[0][0].ToString();

            return s1;

        }


        // Функция которая обрабатывает нажатие кнопки для выполнения шестого запроса
        private void button6_Click(object sender, EventArgs e)
        {
            query("Select  top 1 name as 'Газета', max(count) as 'Количество', mail.idp as 'Номер почтового отделения', adr as 'Адрес' from mail, orders, newspaper " +
                "where newspaper.idn=orders.idN and mail.idp=orders.idP and name='" + comboBox5.Text + "' group by name, mail.idp, adr  order by max(count) desc");
         
        }

        // Функция которая обрабатывает нажатие кнопки для выполнения восьмого запроса
        private void button8_Click(object sender, EventArgs e)
        {
            query("Select t1.nameT, count(*) from " +
                " (Select  nameT, name  from newspaper, orders, postOffice " +
                " where postOffice.idT = orders.idT and newspaper.idn = orders.idN "+
                " group by nameT, name) as t1 group by t1.nameT");
          
        }


        // Функция которая обрабатывает нажатие кнопки для выполнения седьмого запроса
        private void button7_Click(object sender, EventArgs e)
        {
            query("Select  t1.name as 'Газета' from (Select name, nameT  from orders, postOffice, newspaper" +
                 " where newspaper.idn = orders.idN and postOffice.idT = orders.idT group by name, nameT) as t1 group by t1.name having count(*) > 1");
            

        }
    }
}
