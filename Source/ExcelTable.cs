using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using ClosedXML.Excel;
using System.Runtime.Remoting.Messaging;

namespace bdays_app
{
    public class ExcelTable
    {
        public string path = "";    //путь к файлу
        FileStream stream;
        IExcelDataReader reader;
        DataSet db;
        DataTable table;     


        public ExcelTable(string path = "")
        {
            this.path = path;
        }

        //открыть файл xlsx 
        public DataTableCollection OpenFile(DataTableCollection tableCollection)
        {
            stream = File.Open(path, FileMode.Open);

            reader = ExcelReaderFactory.CreateReader(stream);
            db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            }
            );

            tableCollection = db.Tables;

            return tableCollection;
        }

        //открыть лист в dataGridView
        public void OpenSheet(DataTableCollection tableCollection, DataGridView dataGridView1, int Sheet)
        {
            dataGridView1.DataSource = tableCollection[Sheet];
        }

        //считывание таблицы построчно, сортировка дат по степени близости
        public void FindComing(DataTableCollection tableCollection, int Sheet)
        {
            table = tableCollection[Sheet];
            foreach (DataRow row in table.Rows)
            {
                string[] record = {row[0].ToString(), Convert.ToDateTime(row[1]).ToShortDateString()};

                AddToTables(tableCollection, record); //проверяем основную таблицу на наличие необходимых дат 
            }
        }

        //изменение даты рождения или удаление записи 
        public bool ChangeRecord(DataTableCollection tableCollection, int Sheet, string[] record, bool action) //1 - change, 2 - delete
        {
            bool flag = false;
            table = tableCollection[Sheet];
            foreach (DataRow row in table.Rows)
            {
                if (record[0].ToString().ToLower().Trim() == row[0].ToString().ToLower().Trim())
                {
                    if (action)
                    {
                        row[1] = record[1];
                        Recheck(tableCollection, Sheet, record, row); //добавляем или удаляем из остальных таблиц
                        flag = true;
                        return flag;
                    }
                    else
                    {   if (Convert.ToDateTime(record[1]).ToShortDateString() == Convert.ToDateTime(row[1]).ToShortDateString())
                        table.Rows.Remove(row);
                        flag = true;
                        return flag;
                    }
                }
            }
            return flag;
        }

        //добавление записи
        public void AddRecord(DataTableCollection tableCollection, int Sheet, string[] record)
        {
            tableCollection[Sheet].Rows.Add(record);           
        }

        //добавление записи в табл ближайших или сегодняшних др
        public void AddToTables(DataTableCollection tableCollection, string[] record)
        {
            if (Convert.ToDateTime(record[1]).Day == DateTime.Today.Day
                    && Convert.ToDateTime(record[1]).Month == DateTime.Today.Month)
            {
                AddRecord(tableCollection, 1, record);
            }
            else if (CheckSoonBday(record))
            {
                AddRecord(tableCollection, 2, record);
            }
        }

        //провека измененной записи на выполнение условий
        public void Recheck(DataTableCollection tableCollection, int Sheet, string[] record, DataRow row)
        {
            table = tableCollection[Sheet];

            if (Sheet == 1 && (Convert.ToDateTime(record[1]).Day != DateTime.Today.Day || Convert.ToDateTime(record[1]).Month != DateTime.Today.Month))
            {
                table.Rows.Remove(row);
                if (!IsInTable(tableCollection, Sheet, record)) 
                    AddToTables(tableCollection, record);
            }
            else
                if (Sheet == 2 && !CheckSoonBday(record))
            {
                table.Rows.Remove(row);
                if (!IsInTable(tableCollection,Sheet,record))
                    AddToTables(tableCollection, record);
            }
        }

        //проверка, скоро ли др
        public bool CheckSoonBday(string[] record)
        {
            DateTime checkDate = new DateTime(DateTime.Today.Year, Convert.ToInt32(Convert.ToDateTime(record[1]).Month), Convert.ToInt32(Convert.ToDateTime(record[1]).Day));

            if (checkDate.Date > DateTime.Today && checkDate.Date < DateTime.Today.AddDays(7))
                return true;
            else return false;
        }

        //проверк, есть ли запись в таблице
        public bool IsInTable (DataTableCollection tableCollection, int Sheet, string[] record)
        {
            table = tableCollection[Sheet];

            foreach (DataRow row in table.Rows)
            {
                if (record[0].ToString().ToLower().Trim() == row[0].ToString().ToLower().Trim())
                    return true;
            }
            return false;
        }

        //сохранение таблицы
        public void SaveSheet(DataGridView dataGridView, DataTableCollection tableCollection, int Sheet)
        {
            stream.Close();

            int rowNum = dataGridView.RowCount;

            var workbook = new XLWorkbook(path);
            var ws = workbook.Worksheet(++Sheet);

            int column = 0;
            int row=0;

            ws.Clear();

            ws.Range("A1").Value = "Name";
            ws.Range("B1").Value = "BDay";

            if (rowNum == 0) {
                ws.Columns().AdjustToContents();
                workbook.SaveAs(path);
                return; 
            }

            var rngNumbers = ws.Range($"A2:A"+ rowNum);
            foreach (var cell in rngNumbers.Cells() )
            {
                string formattedString = dataGridView[column, row].Value.ToString();
                cell.DataType = XLDataType.Text;
                cell.Value = formattedString;
                row++;
            }
            column++;
            row = 0;
            rngNumbers = ws.Range($"B2:B" + rowNum);
            foreach (var cell in rngNumbers.Cells())
            {
                DateTime formattedString = Convert.ToDateTime(dataGridView[column, row].Value).Date;
                cell.DataType = XLDataType.DateTime;
                cell.Value = formattedString;
                row++;
            }

            ws.Columns().AdjustToContents();

            workbook.SaveAs(path);
        }
    }
}
