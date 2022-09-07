using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace bdays_app
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        ExcelTable Tab =  new ExcelTable("..\\bdays.xlsx");

        public DataTableCollection tableCollection = null;

        //при загрузке формы вывод 3х таблиц: 1 полной и 2 с выбранными данными
        public void Form1_Load(object sender, EventArgs e)
        {
            tableCollection = Tab.OpenFile(tableCollection);
            tableCollection[1].Clear();
            tableCollection[2].Clear();

            Tab.FindComing(tableCollection, 0); //ищем ближ др
            Tab.OpenSheet(tableCollection, dataGridView1, 1);   //др сегодня
            Tab.OpenSheet(tableCollection, dataGridView2, 2);   //др ближ 7 дней
            Tab.OpenSheet(tableCollection, dataGridView3, 0);   //все др
        }

        //добавление записи
        private void button1_Click(object sender, EventArgs e)
        {
            string[] record = {textBox1.Text, dateTimePicker1.Value.Date.ToString()};
            if (Tab.IsInTable(tableCollection, 0, record))
            {
                MessageBox.Show("Такая запись уже есть..", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.None);
                return;
            }
            Tab.AddRecord(tableCollection, 0, record); //добавить в осн табл
            Tab.AddToTables(tableCollection, record); //меняем доп таблицы
        }

        //редактирование даты рождения по имени
        private void button2_Click(object sender, EventArgs e)
        {
            string[] record = { textBox1.Text, dateTimePicker1.Value.Date.ToString() }; //введенные в форме данные
            bool flag = Tab.ChangeRecord(tableCollection, 0, record, true); //проверка на их существование в таблице
            if (!flag)
            {
                MessageBox.Show("Такой записи нет..", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.None);
                return;
            }
            //изменение доп таблиц
            if (Tab.IsInTable(tableCollection, 1, record))
                Tab.ChangeRecord(tableCollection, 1, record, true);
            else if (Tab.IsInTable(tableCollection, 2, record))
                Tab.ChangeRecord(tableCollection, 2, record, true);
            else
                Tab.AddToTables(tableCollection, record);
        }

        //удаление записи по имени и дате рождения
        private void button3_Click(object sender, EventArgs e)
        {
            string[] record = { textBox1.Text, dateTimePicker1.Value.Date.ToString() }; //требуемая запись
            bool flag = Tab.ChangeRecord(tableCollection, 0, record, false);//проверка на существование удаление в осн таблице
            if (!flag)
            {
                MessageBox.Show("Такой записи нет..", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.None);
                return;
            }
            //удаление в доп таблицах при наличии
            Tab.ChangeRecord(tableCollection, 1, record, false);
            Tab.ChangeRecord(tableCollection, 2, record, false);
        }

        //сохранение таблицы по листам
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                Tab.SaveSheet(dataGridView3, tableCollection, 0);
                Tab.SaveSheet(dataGridView1, tableCollection, 1);
                Tab.SaveSheet(dataGridView2, tableCollection, 2);
                MessageBox.Show("Успешно сохранено!!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.None);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
