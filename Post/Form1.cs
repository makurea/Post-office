using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;



namespace Post
{
    public partial class Form1 : Form
    {
        //Глобальная переменная которая хранит название выбранной таблицы
        private String activeTable = "none";
        public Form1()
        {
            
            InitializeComponent();
        
        }

        //Кнопка по нажатию на которую отображается данные таблицы "Газеты"
        private void button1_Click(object sender, EventArgs e)
        {
            //Задаем имя выбранной таблицы в глобальную перменную
            activeTable = "newspaper";
            //Вызываем функцию загрузки данных из бд в таблицу в приложении
            loadTable();
            //Скрываем первый столбец. Столбец с айди записей
            dataGridView1.Columns[0].Visible = false;
            //Устанавливаем заголовки датагрида
            dataGridView1.Columns[1].HeaderText = "Название газеты";
            dataGridView1.Columns[2].HeaderText = "Индекс издания";
            dataGridView1.Columns[3].HeaderText = "ФИО редактора";
            dataGridView1.Columns[4].HeaderText = "Цена одной газеты";


        }

        //Загрузка данных из бд в таблицу в приложении
        private void loadTable()
        {           
            
            SqlDataAdapter da;
            //Передаем запрос в датаадаптер 
            da = new SqlDataAdapter("select * from " + activeTable , Ress.connect);
            DataSet ds = new DataSet();
            da.Fill(ds, activeTable);   
            //Загружаем данные непосредственное в датагрид     
            dataGridView1.DataSource = ds.Tables[activeTable];
            //Обнавление датагрида
            dataGridView1.Refresh();
            //Сбрасываем выделение какой-либо записи. При загрузки данных ни одна запись не должны 
            //быть активной пока ее не выберет пользователь
            dataGridView1.ClearSelection();
        }

        //Кнопка по нажатию на которую отображается данные таблицы "Типография"
        private void button2_Click(object sender, EventArgs e)
        {
            activeTable = "postOffice";
            loadTable();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Название типографии";
            dataGridView1.Columns[2].HeaderText = "Адрес";
        }

        //Кнопка по нажатию на которую отображается данные таблицы "Почта"
        private void button3_Click(object sender, EventArgs e)
        {
            activeTable = "mail";
            loadTable();
            dataGridView1.Columns[0].HeaderText = "Номер почтового отделения";
            dataGridView1.Columns[1].HeaderText = "Адрес";

        }

        //Кнопка по нажатию на которую отображается данные таблицы "Заказы"
        private void button4_Click(object sender, EventArgs e)
        {
            activeTable = "ordersView";
            loadTable();
            dataGridView1.Columns[0].HeaderText = "код заказа";
            dataGridView1.Columns[1].HeaderText = "типография";
            dataGridView1.Columns[2].HeaderText = "газета";
            dataGridView1.Columns[3].HeaderText = "индекс почты";
            dataGridView1.Columns[4].HeaderText = "количество";

        }

        //Кнопка удаления записи
        private void button7_Click(object sender, EventArgs e)
        {
            
            try
            {
                     
                //Определяем имя айдишника которое будет использоваться в запросе. Зависит от имени таблицы         
                String id = "";
                switch (activeTable)
                {
                    case "newspaper":                        
                        id = "idn"; break;
                    case "postOffice":
                        id = "idT"; break;
                    case "mail":
                        id = "idp"; break;
                    case "ordersView":
                        //Здесь кроме присвоения айди также изменяем имя активной таблицы т.к. в 
                        //приложении данные отображаются из представления, а нам нужно удалять их из исходной таблицы
                        id = "idZ"; activeTable = "orders"; break;
                    default: break;

                }

              
                //Проверка на то, не содержится ли удаляемая запись в других таблицах
                //Например нельзя удалить почту если ее индекс используется в таблице заказы
                if (verify())
                {
                    //Само удаление из бд
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandText = "delete from " + activeTable + " where " + id + "=" + dataGridView1.CurrentRow.Cells[0].Value.ToString();
                    cmd.Connection = Ress.connect;
                    cmd.Connection.Open();
                    cmd.ExecuteNonQuery();
                    cmd.Connection.Close();
                    //Если запись была удалена из таблицы заказов, то меняем обратно параметр activetable
                    //чтобы вернуться к работе с представлением
                    if (activeTable.Equals("orders"))
                    {
                        activeTable = "ordersView";
                    }
                    //Перезагружаем таблицу с обнавленными данными
                    loadTable();
                }else
                {
                    MessageBox.Show("Данная запись содержится в других таблицах");
                }
                
           }
            catch(Exception) 
            {
                //Обработка ошибки на случай если пользователь нажал на кнопку удалить не выбрав запись
                MessageBox.Show("Выберите удаляемую строку");
            }
           
            

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        //Переход на форму добавления/изменения
        private void button5_Click(object sender, EventArgs e)
        {
            //Скрываем данную форму
            this.Hide();
            //Создаем обьект новой формы, передавая туда выбранную таблицу
            AddForm af = new AddForm(activeTable);
            //показываем
            af.ShowDialog();
          
        }

        //Обработка кнопки изменить
        private void button6_Click(object sender, EventArgs e)
        {
            
            try
            {
                //Записываем id записи которая быра выбрана для изменения
                String id = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                this.Hide();
                //Созздаем форму добавления/изменения передавая туда имя выбранной таблицы и айди выбранной записи
                AddForm cf = new AddForm(activeTable, id);
                cf.ShowDialog();
            }
            catch (Exception)
            {
                MessageBox.Show("Выберите изменяемую строку");
            }
            
        }

        //Функция которая проверяет можно ли удалить эту запись. А именно содержится ли она в других таблицах
        bool verify()
        {
            String id1 = "";
            switch (activeTable)
            {
                case "newspaper": id1 = "idN"; break;
                case "postOffice": id1 = "idT"; break;
                case "mail": id1 = "idP"; break;
                case "ordersView": id1 = "idZ"; break;
                default: break;

            }
            try
            {
                    SqlDataAdapter da;
                    //проверяем запросом есть ли указанная запись в таблице Заказы
                    da = new SqlDataAdapter("select * from orders where " + id1 + "=" + dataGridView1.CurrentRow.Cells[0].Value.ToString(), Ress.connect);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    //Присваиваем строке результат запроса. Если что-то вернуло,
                    //то такая запись содержится в другой таблице и удалять нам ее нельзя, не удалив
                    //предварительно из другой таблицы. Если не вернуло ничего, то вылетит ошибка которую мы ловим ниже
                    String s1 = dt.Rows[0][0].ToString();

                     return false;
            }
            catch (Exception) {
                //Если вылетеле ошибка, то запрос ничего не вернул, значит такой записи нет в инетерсующей
                //нас таблице, значит фун-ия должна вернуть true
                return true;
            }


        }

        //Кнопка оформления заказ. По сути таже кнопка добавления, только конкретно для таблицы заказы
        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            AddForm af = new AddForm("ordersView");
            af.ShowDialog();
        }


        //Кнопка выхода из приложения
        private void button10_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //Обработка кнопки "Отчеты". Открывает соответствующую форму
        private void button9_Click(object sender, EventArgs e)
        {
            this.Hide();
            reports r = new reports();
            r.ShowDialog();
        }


        //Функция обработки нажатия кнопки "Отчет". Генерирует отчет и выводит его в xml файле
        private void button11_Click(object sender, EventArgs e)
        {

           

            try
            {
                //Получение списка типографии
                SqlDataAdapter da = new SqlDataAdapter("Select nameT from postOffice", Ress.connect);
                DataSet ds = new DataSet();
                da.Fill(ds, "postOffice");

                Excel.Application exApp = new Excel.Application();      //создание объекта

                exApp.Workbooks.Add(Type.Missing);  //новая рабочая книга
                Excel.Worksheet workSheet = (Excel.Worksheet)exApp.ActiveSheet;         //получили активный лист

                //Добавление строки в экселе с текстом "Отчет по типографиям"
                workSheet.Cells[2, "F"] = "Отчет по типографиям ";
                
                
                int z = 0;
                //Проходим циклом по таблице с типографиями
                for (int i = 0; i < ds.Tables["postOffice"].Rows.Count; i++)
                {
                    //Пишем название типографии
                    workSheet.Cells[i + 6 + z, "D"] = ds.Tables["postOffice"].Rows[i][0].ToString();              
                    workSheet.Cells[i + 6 + z, "D"].Interior.ColorIndex = 43;

                    //Подсчитываем общее количество видов печатающихся в каждой типографии газет
                    SqlDataAdapter da1 = new SqlDataAdapter("Select count(*) from " +
                        "(Select distinct name from orders, newspaper" + 
                        " where idT = (Select idT from postOffice where nameT = '" + ds.Tables["postOffice"].Rows[i][0].ToString() + "')" + 
                        " and newspaper.idn = orders.idN) as t1", Ress.connect);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, "countAll");
                    //Записываем подсчитанную информацию в файл
                    workSheet.Cells[i + 7 + z, "D"] = "общее количество видов печатающихся в типографии газет";
                    workSheet.Cells[i + 7 + z, "E"] = ds1.Tables["countAll"].Rows[0][0].ToString();

                    //Подсчитываем общее количество напечатанных экзепляров газет в каждой типографии
                    da1 = new SqlDataAdapter("Select sum(count) from orders where" +
                        " idT=(Select idT from postOffice where nameT='" + ds.Tables["postOffice"].Rows[i][0].ToString() + "')", Ress.connect);
                    ds1 = new DataSet();
                    da1.Fill(ds1, "countAllNP");
                    //Записываем подсчитанную информацию в файл
                    workSheet.Cells[i + 8 + z, "D"] = "общее количество напечатанных экзепляров газет в типографии";
                    workSheet.Cells[i + 8 + z, "E"] = ds1.Tables["countAllNP"].Rows[0][0].ToString();

                    //Получение данных для цикла в котором будет подсчитываться количество газет каждого наименования
                    da1 = new SqlDataAdapter("select nameT, name, sum(count) from orders, newspaper,postOffice " +
                        "where newspaper.idn = orders.idN and postOffice.idT = orders.idT group by nameT, name", Ress.connect);
                    ds1 = new DataSet();
                    da1.Fill(ds1, "forLoop");
                    int k = 0;
                    workSheet.Cells[i + 9 + z, "D"] = "Количество газет каждого наименования";
                    //Цик в котором подсчитывается количество газет каждого наименования 
                    for (int j = 0; j < ds1.Tables["forLoop"].Rows.Count; j++)
                    {
                        if (ds1.Tables["forLoop"].Rows[j][0].ToString().Equals(ds.Tables["postOffice"].Rows[i][0].ToString()))
                        {
                            workSheet.Cells[i + 9 + z + k, "E"] = ds1.Tables["forLoop"].Rows[j][1].ToString();
                            workSheet.Cells[i + 9 + z + k, "F"] = ds1.Tables["forLoop"].Rows[j][2].ToString();
                            
                            k++;
                        }
                    }

                    //Получение данных для цикла в котором будет подсчитываться 
                    //какие газеты и в каком количестве типография отправляет в каждое почтовое отделение;

                    da1 = new SqlDataAdapter("select nameT, mail.idp, name, sum(count) from orders, newspaper, postOffice,mail" +
                        " where newspaper.idn = orders.idN and postOffice.idT = orders.idT and mail.idp = orders.idP" +
                        " group by nameT, mail.idp, name", Ress.connect);
                    ds1 = new DataSet();
                    da1.Fill(ds1, "forLoop1");

                    workSheet.Cells[i + 9 + z + k, "D"] = "какие газеты и в каком количестве типография отправляет в каждое почтовое отделение";
                   
                    workSheet.Cells[i + 10 + z + k, "D"] = "Номер почтового отделения ";
                    workSheet.Cells[i + 10 + z + k, "E"] = "Название газеты";
                    workSheet.Cells[i + 10 + z + k, "F"] = "Количество";

                    // Цикл в котором подсчитывается какие газеты и в каком количестве типография отправляет в каждое почтовое отделение";

                    for (int j = 0; j < ds1.Tables["forLoop1"].Rows.Count; j++)
                    {
                        if (ds1.Tables["forLoop1"].Rows[j][0].ToString().Equals(ds.Tables["postOffice"].Rows[i][0].ToString()))
                        {
                            workSheet.Cells[i + 11 + z + k, "D"] = ds1.Tables["forLoop1"].Rows[j][1].ToString();
                            workSheet.Cells[i + 11 + z + k, "E"] = ds1.Tables["forLoop1"].Rows[j][2].ToString();
                            workSheet.Cells[i + 11 + z + k, "F"] = ds1.Tables["forLoop1"].Rows[j][3].ToString();                            
                            k++;
                        }
                    }


                        z += 20;


                }
                workSheet.get_Range("A1", "X999").EntireColumn.AutoFit();

                //Создаем имя файла и сохраняем его
                String pathToXmlFile;
                pathToXmlFile = Path.GetTempPath() + "Отчет по типогрфиям за " + DateTime.Now.Year + " год.xlsx";
                workSheet.SaveAs(pathToXmlFile);
                Process.Start(pathToXmlFile);
                exApp.Visible = true;

                MessageBox.Show("Отчет составлен");
               
            }
            catch (COMException)
            {
                MessageBox.Show("Нет доступа к сектору либо файл уже существует и открыт", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }




        }
    }
}
