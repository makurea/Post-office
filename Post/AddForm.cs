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
    public partial class AddForm : Form
    {
        //Глобальные переменные которые хранят имя таблицы и айди изменяемой записи
        String activeTable;
        String idVal;
        public AddForm()
        {
           
            InitializeComponent();
            
        }

        //Конструктор формы с 2-мя параметрами. Первый параметр это имя таблицы, передается всегда
        //Второй параметр - параметр по умолчанию. Это айди изменяемой записи. Передается только тогда, когда
        //была нажата кнопка изменить. По умолчанию значение false
        public AddForm(String table, String id = "false")
        {
            
            InitializeComponent();
            //Название выбранной таблицы при нажатии кнопки
            activeTable = table;
            idVal = "false";
            //Проверка какая кнопка была нажата. Добавить или Изменить
            //Если изменить, то флаг не будет равен False и мы заносим в idVal - ид записи которую выбрали для изменения
            if (!id.Equals("false"))
            {
                idVal = id;               
                initialFields();
            }
            //В зависимости от таблицы показываем на форме нужные нам значения лейбелов и поля
            switch (activeTable)
            {
                case "newspaper":    showForN();  break;
                case "postOffice":   showForP();  break;
                case "mail":         showForM();  break;
                case "ordersView":   showForO();  break;
                default:                            break;

            }
            label2.Text = "Добавление данных";
        }

        private void AddForm_Load(object sender, EventArgs e)
        {

        }


        //Подготовка к отображению формы для таблицы Газета
        void showForN()
        {
            label3.Text = "Добавление данных для таблицы 'Газета'";
            label3.Text = "Введите название газеты";
            label4.Text = "Введите индекс издания";
            label5.Text = "Введите ФИО редактора";
            label6.Text = "Введите цену одной газеты";
            comboBox1.Visible = false;
            comboBox2.Visible = false;
            comboBox3.Visible = false;
        }

        //Подготовка к отображению формы для таблицы Типография
        void showForP()
        {
            label3.Text = "Добавление данных для таблицы 'Типография'";
            label3.Text = "Введите название типографии";
            label4.Text = "Введите адрес";
            label5.Visible = false;
            label6.Visible = false;
            textBox3.Visible = false;
            textBox4.Visible = false;
            comboBox1.Visible = false;
            comboBox2.Visible = false;
            comboBox3.Visible = false;
            
        }

        //Подготовка к отображению формы для таблицы Почта
        void showForM()
        {
            label3.Text = "Добавление данных для таблицы 'почта'";
            label3.Text = "Введите адрес почтового отделения";
            label4.Visible = false;
            label5.Visible = false;
            label6.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false; 
            textBox4.Visible = false;
            comboBox1.Visible = false;
            comboBox2.Visible = false;
            comboBox3.Visible = false;

        }

        //Подготовка к отображению формы для таблицы Заказы
        void showForO()
        {
            label3.Text = "Добавление данных для таблицы 'Заказы'";
            label3.Text = "Выберите типографию";
            label4.Text = "Введите название газеты";
            label5.Text = "Выберите номер почтового отделения";
            label6.Text = "Введите количество экземпляров";
            forComboBox("nameT", "postOffice", comboBox1);
            forComboBox("name","newspaper", comboBox2);
            forComboBox("idP","mail", comboBox3);
            textBox1.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false;


        }

        //Кнопка назад
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 f = new Form1();
            
            f.ShowDialog();
        }

        //Кнопка сохранить
        private void button1_Click(object sender, EventArgs e)
        {
            String query = "";
            //Вызов функции в зависимости от того какая была нажата кнопка на форме 1, добавить или изменить
            if (idVal.Equals("false"))            {
                query = addInfo();              
            }
            else
            {
                query = changeInfo();              

            }

            try
            {
                //Занесение данных в бд
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = query;
                cmd.Connection = Ress.connect;
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();
                cmd.Connection.Close();
                //Уведомление пользователя о успешном добавлении или изменении записи
                if (idVal.Equals("false"))
                {
                    MessageBox.Show("Запись успешно добавлена");
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                }
                else
                {
                    MessageBox.Show("Запись успешно изменена");
                    this.Hide();
                    Form1 f = new Form1();
                    f.ShowDialog();
                }

            }
            catch
            {   //Уведомление о том, что перед сохранением не все поля были запонены
                MessageBox.Show("Заполнены не все поля"); ;
            }
           

        }


        //Функция которая генерирует запрос добавления данных
        String addInfo()
        {
            String query = "";
            //Запрос генерируется на основании того в какую таблицу нужно добавлять данные
            switch (activeTable)
            {
                case "newspaper":
                    query = "insert into newspaper (name, indexIzd , FIO, price) " +
                             "values ('" + textBox1.Text.ToString() + "', '" + textBox2.Text.ToString() + "', '" + textBox3.Text.ToString() + "'," + textBox4.Text.ToString() + ")";
                    break;
                case "postOffice":
                    query = " insert into postOffice (nameT, adrT)" +
                    " values ('" + textBox1.Text.ToString() + "', '" + textBox2.Text.ToString() + "')";
                    break;
                case "mail":
                    query = " insert into mail (adr)" +
                    " values ('" + textBox1.Text.ToString() + "')"; break;
                case "ordersView":
                    //запрос добавления данных в таблицу заказы несколько сложнее т.к. пользователь работает
                    // с представлением, а нам нужно заносить данные в исходную таблицу в которой хранятся
                    //коды записей. Поэтому записи нужно преобразовывать 
                    query = "insert into orders (idT, idN, idP, count) " +
                        "values (" + textToId(comboBox1.Text, "nameT","postOffice") + ", "
                        + textToId(comboBox2.Text, "name", "newspaper") + ", "
                        + textToId(comboBox3.Text, "idp", "mail") + ", "
                        + textBox4.Text.ToString() + ")";
                    break;
                default: break;

            }

            return query;
        }

        //Функция которая генерирует запрос изменения данных. Структура аналагична структуре функции выше
        String changeInfo()
        {
            String query = "";
           // Запрос генерируется на основании того в какой таблице  нужно изменять данные
            switch (activeTable)
            {
                case "newspaper":
                    query = "Update " + activeTable +
                            " set name='" + textBox1.Text.ToString() +
                            "', indexIzd='" + textBox2.Text.ToString() +
                            "',FIO ='" + textBox3.Text.ToString() +
                            "',price=" + textBox4.Text.ToString() +
                            "where idn=" + idVal;
                    break;

                case "postOffice":
                    query = "Update " + activeTable +
                            " set nameT='" + textBox1.Text.ToString() +
                            "', adrT='" + textBox2.Text.ToString() +
                            "' where idT=" + idVal;
                    break;
                case "mail":
                    query = "Update " + activeTable +
                            " set adr='" + textBox1.Text.ToString() +
                            "' where idp=" + idVal;
                    break;
                case "ordersView":
                    query = "Update orders set idT ="
                        + textToId(comboBox1.Text, "nameT", "postOffice") +
                        ", idN=" + textToId(comboBox2.Text, "name", "newspaper") +
                        ", idP=" + textToId(comboBox3.Text, "idp", "mail") +
                        ", count=" + textBox4.Text.ToString() +
                        " where idZ=" + idVal;

                    break;
                default: break;

            }
            return query;
        }

        //Функция инициализации полей. Вызывается если мы попали на форму при помощи нажатия
        //кнопки изменить. Заполняются поля начальными значениями из изменяемой записи
        void initialFields()
        {

            String id1 = "";
            //Заносим в строку id1 название id которое зависит от имени таблицы с которой мы будет работать
            switch (activeTable)
            {
                case "newspaper": id1 = "idn"; break;
                case "postOffice": id1 = "idT"; break;
                case "mail": id1 = "idp"; break;
                case "ordersView": id1 = "idZ"; break;
                default: break;

            }

       
            SqlDataAdapter da;
            //Получение данных
            da = new SqlDataAdapter("select * from " + activeTable + " where " + id1 + "=" + idVal, Ress.connect);
            DataSet ds = new DataSet();
            da.Fill(ds, activeTable);
            //Заполнение полей в зависимости от имени таблицы
            switch (activeTable)
            {
                case "newspaper":
                    textBox1.Text = ds.Tables[activeTable].Rows[0][1].ToString();
                    textBox2.Text = ds.Tables[activeTable].Rows[0][2].ToString();
                    textBox3.Text = ds.Tables[activeTable].Rows[0][3].ToString();
                    textBox4.Text = ds.Tables[activeTable].Rows[0][4].ToString();
                    break;

                case "postOffice":
                    textBox1.Text = ds.Tables[activeTable].Rows[0][1].ToString();
                    textBox2.Text = ds.Tables[activeTable].Rows[0][2].ToString();
                   

                    break;
                case "mail":
                    textBox1.Text = ds.Tables[activeTable].Rows[0][1].ToString();
                   
                    break;
                case "ordersView":

                    comboBox1.Text = ds.Tables[activeTable].Rows[0][1].ToString();
                    comboBox2.Text = ds.Tables[activeTable].Rows[0][2].ToString();
                    comboBox3.Text = ds.Tables[activeTable].Rows[0][3].ToString();
                    textBox4.Text = ds.Tables[activeTable].Rows[0][4].ToString();
                    
                    break;
                default: break;

            }
            
            
        }


        //Функция которая заполняет список комбобокса данными из бд.
        //Получает 3 параметра таблицу и столбец откуда нужно брать данные и комбобок в который их записывать
        private void forComboBox(String field, String table, ComboBox cb)
        {     
            //Запрос получения данных из бд   
            SqlDataAdapter da = new SqlDataAdapter("Select " + field + " from " + table, Ress.connect);
            DataTable dataTable = new DataTable();
            DataSet ds = new DataSet();
            da.Fill(ds, activeTable);
            //Занесение полученных данных в комбобокс
            for (int i = 0; i < ds.Tables[activeTable].Rows.Count; i++)
            {
                cb.Items.Add(ds.Tables[activeTable].Rows[i][0].ToString());
            }          

        }


        //Функция которая переводит текст в айди. Данная функция нужна в свзяи с тем,
        //что пользователь работает с представлением таблицы заказы. В данном представлении хранятся
        //названия, но в исходной таблицы хранятся коды. При добавлении или изменении записи
        // пользователь работает с названиями, а данная функция переопределяет в коды данные названия и
        //записывает их в исходную таблицу.
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
            //Запрос которые получает код исходя из конкретного названия
            SqlDataAdapter da = new SqlDataAdapter("Select " + id1 +" from " + table + " where " + field + " = '" + data + "'", Ress.connect);
            DataTable dt = new DataTable();
            da.Fill(dt);
            //Записывает данный код в строку и возвращает его
            String s1 = dt.Rows[0][0].ToString();
           
            return s1;

        }

      


    }
}
