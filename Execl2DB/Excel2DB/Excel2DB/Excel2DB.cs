using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace Excel2DB
{
    public partial class Excel2DB: UserControl
    {
        private Excel2DBImporter excel2DBImporter;
        private System.Data.DataTable excelTable;
        public Excel2DB()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel2DBImporter.ConnectionParaHostName = textBox1.Text;
            excel2DBImporter.ConnectionParaDBName = textBox2.Text;
            excel2DBImporter.ConnectionParaUserName = textBox3.Text;
            excel2DBImporter.ConnectionParaPassWord = textBox4.Text;
            excel2DBImporter.ConnectionParaPort = textBox5.Text;

            comboBox1.Items.Clear();

            try
            {
                foreach (string tableName in excel2DBImporter.GetTables())
                {
                    comboBox1.Items.Add(tableName);
                }
                labelCurrentDB.Text = excel2DBImporter.ConnectionParaDBName;
                labelCurrentHost.Text = excel2DBImporter.ConnectionParaHostName;
                comboBox1.SelectedIndex = 0;
            }
            catch (Exception)
            {

                MessageBox.Show("连接不成功或者数据库内不包含任何表！");
            }


          
        }

        private void Excel2DB_Load(object sender, EventArgs e)
        {
            excel2DBImporter = new Excel2DBImporter();

            excel2DBImporter.ConnectionParaDataBaseType = DataBaseType.MySql;
            excel2DBImporter.ConnectionParaHostName = textBox1.Text;
            excel2DBImporter.ConnectionParaDBName = textBox2.Text;
            excel2DBImporter.ConnectionParaUserName = textBox3.Text;
            excel2DBImporter.ConnectionParaPassWord = textBox4.Text;
            excel2DBImporter.ConnectionParaPort = textBox5.Text;

            dataGridView1.Rows.Add(100);

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Rows.Add(100);
            excel2DBImporter.TransferParaTableName = comboBox1.Text;
            comboBox2.Items.Clear();
            try
            {
                int i = 0;
                
                foreach (string columnName in excel2DBImporter.GetKeys())
                {
                    comboBox2.Items.Add(columnName);
                    dataGridView1.Rows[i].Cells["DataBaseKeys"].Value = columnName;
                    i++;
                }
                labelCurrentTable.Text = excel2DBImporter.TransferParaTableName;
                comboBox2.SelectedIndex = 0;
            }
            catch (Exception)
            {

                MessageBox.Show("表内不包含任何列！");
            }



        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            excel2DBImporter.TransferParaPrimeKey = comboBox2.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> transferParaSourceKeys = new List<string>();
            List<string> transferParaDataBaseKeys = new List<string>();

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != null)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ","") != "")
                    {
                        transferParaSourceKeys.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                        transferParaDataBaseKeys.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    }
                }
            }
            excel2DBImporter.TransferParaDataBaseKeys = transferParaDataBaseKeys;
            excel2DBImporter.TransferParaSourceKeys = transferParaSourceKeys;

            if (!excel2DBImporter.ConnectionTest()|| labelCurrentDB.Text == "未选择")
            {
                MessageBox.Show("没有连接到数据库，请检查设置！");
                return;
            }

            if (labelCurrentFile.ForeColor == Color.Black)
            {
                MessageBox.Show("选择了源文件但是没有应用设置，请确认参数后点击\"应用源文件参数\"！");
                return;
            }

            if (excelTable == null||excel2DBImporter.SourceFileName == "")
            {
                MessageBox.Show("没有选择源文件，或者源文件内没有数据！");
                return;
            }

            foreach (string key in excel2DBImporter.TransferParaSourceKeys)
            {
                bool isExist = false;
                foreach (DataColumn column in excelTable.Columns)
                {
                    if (column.ColumnName == key)
                    {
                        isExist = true;
                        continue;
                    }
                }
                if (!isExist)
                {
                    MessageBox.Show("源数据表中不包含名为\"" + key + "\"的列！");
                    return;
                }
            }

            string dispStr = "你确定要从文件 " + excel2DBImporter.SourceFileName + " 向主机地址为 " + excel2DBImporter.ConnectionParaHostName + " \r\n的 " + excel2DBImporter.ConnectionParaDBName + "." + excel2DBImporter.TransferParaTableName + " 表中添加数据吗？";
            ConfirmWindow confirmWindow = new ConfirmWindow(dispStr, excel2DBImporter.TransferParaSourceKeys, transferParaDataBaseKeys);
            confirmWindow.ShowDialog();

            if (confirmWindow.DialogResult == DialogResult.Yes)
            {
                try
                {
                   
                    excel2DBImporter.Import(excelTable);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("数据导入出错，请检查源表格数据格式！");
                    MessageBox.Show(excel2DBImporter.CurrentSQLString);
                }


                MessageBox.Show("导入成功！");

            }
        


        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel文档|*.xlsx|Excel2003-2007文档|*.xls";
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName == "")
            {
                return;
            }
            excel2DBImporter.SourceFileName = openFileDialog.FileName;
            labelCurrentFile.Text = excel2DBImporter.SourceFileName;
            textBox6.Enabled = true;
            textBox7.Enabled = true;
            textBox8.Enabled = true;

            labelCurrentFile.ForeColor = Color.Black;



        }

        private void button4_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            excel2DBImporter.ExcelRangeInitRow = Convert.ToInt32(textBox6.Text);
            excel2DBImporter.ExcelRangeInitColumn = Convert.ToInt32(textBox7.Text);
            excel2DBImporter.ExcelItemsCount = Convert.ToInt32(textBox8.Text);

            excelTable = excel2DBImporter.GetExcelData();
            if (excelTable == null)
            {
                return;
            }
            int i = 0;
            foreach (DataColumn column in excelTable.Columns)
            {
                dataGridView1.Rows[i].Cells[0].Value = column.ColumnName;
                i++;
            }
            labelCurrentFile.ForeColor = Color.Green;
            button3.Enabled = true;

        }

     
    }


    public enum DataBaseType
    {
        MySql = 1,
        MSSql = 2,
        Oracle = 3,
        Access = 4,
        DB2 = 5,
        SyBase = 6
    }
    public class Excel2DBImporter 
    {
       
        private Dictionary<int, string> excelColumns = new Dictionary<int, string>();
       
     
        /// <summary>
        /// 数据库类型
        /// </summary>
        public DataBaseType ConnectionParaDataBaseType { get; set; }
        /// <summary>
        /// 用户名
        /// </summary>
        public string ConnectionParaUserName { get; set; }
        /// <summary>
        /// 密码
        /// </summary>
        public string ConnectionParaPassWord { get; set; }
        /// <summary>
        /// 主机名称
        /// </summary>
        public string ConnectionParaHostName { get; set; }
        /// <summary>
        /// 数据库名称
        /// </summary>
        public string ConnectionParaDBName { get; set; }
        /// <summary>
        /// （MySql）端口号
        /// </summary>
        public string ConnectionParaPort { get; set; }
        /// <summary>
        /// 源数据的Key
        /// </summary>
        public List<string> TransferParaSourceKeys { get; set; }
        /// <summary>
        /// 数据库的Key
        /// </summary>
        public List<string> TransferParaDataBaseKeys { get; set; }
        /// <summary>
        /// 主键，数据对比的依据
        /// </summary>
        public string TransferParaPrimeKey { get; set; }
        /// <summary>
        /// 要导入的表名
        /// </summary>
        public string TransferParaTableName { get; set; }
        /// <summary>
        /// Excel数据区的起始行
        /// </summary>
        public int ExcelRangeInitRow { get; set; }
        /// <summary>
        /// Excel数据区的起始列
        /// </summary>
        public int ExcelRangeInitColumn { get; set; }

        /// <summary>
        /// 要读取的excel表格中的多少列
        /// </summary>
        public int ExcelItemsCount { get; set; }
        /// <summary>
        /// 要导入的文件路径
        /// </summary>
        public string SourceFileName { get; set; }

        public string CurrentSQLString { get; set; }

        public Excel2DBImporter()
        {
            excelColumns.Add(1, "A");
            excelColumns.Add(2, "B");
            excelColumns.Add(3, "C");
            excelColumns.Add(4, "D");
            excelColumns.Add(5, "E");
            excelColumns.Add(6, "F");
            excelColumns.Add(7, "G");
            excelColumns.Add(8, "H");
            excelColumns.Add(9, "I");
            excelColumns.Add(10, "J");
            excelColumns.Add(11, "K");
            excelColumns.Add(12, "L");
            excelColumns.Add(13, "M");
            excelColumns.Add(14, "N");
            excelColumns.Add(15, "O");
            excelColumns.Add(16, "P");
            excelColumns.Add(17, "Q");
            excelColumns.Add(18, "R");
            excelColumns.Add(19, "S");
            excelColumns.Add(20, "T");
            excelColumns.Add(21, "U");
            excelColumns.Add(22, "V");
            excelColumns.Add(23, "W");
            excelColumns.Add(24, "X");
            excelColumns.Add(25, "Y");
            excelColumns.Add(26, "Z");     
        }

        /// <summary>
        /// 连接测试
        /// </summary>
        /// <returns></returns>
        public bool ConnectionTest()
        {
            switch (ConnectionParaDataBaseType)
            {
                case DataBaseType.MySql:
                    {
                        string connStr = "Database=" + ConnectionParaDBName + ";" +
                              "Data Source=" + ConnectionParaHostName + ";" +
                              "User Id=" + ConnectionParaUserName + ";" +
                              "Password=" + ConnectionParaPassWord + ";" +
                              "pooling=false;" +
                              "CharSet=utf8;" +
                              "port=" + ConnectionParaPort;

                        MySqlConnection mySqlConnection = new MySqlConnection(connStr);
                        bool isConnected = false;
                        try
                        {
                            mySqlConnection.Open();
                            isConnected = true;
                            mySqlConnection.Close();
                        }
                        catch (MySqlException e)
                        {
                            isConnected = false;
                        }

                        return isConnected;
                        
                    }
          
                case DataBaseType.MSSql:
                    {
                        return false;
                    }
                    
                case DataBaseType.Oracle:
                    {
                        return false;
                    }
                   
                case DataBaseType.Access:
                    {
                        return false;
                    }
                    
                case DataBaseType.DB2:
                    {
                        return false;
                    }
                   
                case DataBaseType.SyBase:
                    {
                        return false;
                    }
                    
                default:
                    return false;
                 
            }

            
        }
        /// <summary>
        /// 获得表名
        /// </summary>
        /// <returns></returns>
        public List<string> GetTables()
        {
            if (!ConnectionTest())
            {
                return null;
            }

            List<string> tables = new List<string>();

            string connStr = "Database=" + ConnectionParaDBName + ";" +
                               "Data Source=" + ConnectionParaHostName + ";" +
                               "User Id=" + ConnectionParaUserName + ";" +
                               "Password=" + ConnectionParaPassWord + ";" +
                               "pooling=false;" +
                               "CharSet=utf8;" +
                               "port=" + ConnectionParaPort;

            MySqlConnection mySqlConnection = new MySqlConnection(connStr);
            string sqlStr = "select table_name from information_schema.tables where table_schema = '" + ConnectionParaDBName + "'";
            MySqlDataAdapter mySqlDataAdapter = new MySqlDataAdapter(sqlStr, mySqlConnection);
            mySqlConnection.Close();
            System.Data.DataTable resultTable = new System.Data.DataTable();
            mySqlDataAdapter.Fill(resultTable);

            foreach (DataRow row in resultTable.Rows)
            {
                tables.Add(row[0].ToString());
            }
            return tables;
        }
        /// <summary>
        /// 查找表中的列名
        /// </summary>
        /// <returns></returns>
        public List<string> GetKeys()
        {
            if (!ConnectionTest())
            {
                return null;
            }

            List<string> keys = new List<string>();

            string connStr = "Database=" + ConnectionParaDBName + ";" +
                               "Data Source=" + ConnectionParaHostName + ";" +
                               "User Id=" + ConnectionParaUserName + ";" +
                               "Password=" + ConnectionParaPassWord + ";" +
                               "pooling=false;" +
                               "CharSet=utf8;" +
                               "port=" + ConnectionParaPort;

            MySqlConnection mySqlConnection = new MySqlConnection(connStr);
            string sqlStr = "select COLUMN_NAME from information_schema.columns where table_name='"+ TransferParaTableName + "'";
            MySqlDataAdapter mySqlDataAdapter = new MySqlDataAdapter(sqlStr, mySqlConnection);
            mySqlConnection.Close();
            System.Data.DataTable resultTable = new System.Data.DataTable();
            mySqlDataAdapter.Fill(resultTable);

            foreach (DataRow row in resultTable.Rows)
            {
                keys.Add(row[0].ToString());
            }
            return keys;
        }

        //private string[] LineDataSplit(string rawData)
        //{
        //    bool isInComment = false;
        //    List<string> data;

        //}

        //public System.Data.DataTable GetCSVData()
        //{
        //    StreamReader sw = new StreamReader(SourceFileName, Encoding.UTF8);


        //}

        /// <summary>
        /// 获取excel中的数据表
        /// </summary>
        /// <returns></returns>
        public System.Data.DataTable GetExcelData()
        {
            string[,] arry;
            int excelRowsCount;
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();//lauch excel application
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
            if (excel == null)
            {
                return null;
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;


                excelWorkbook = excel.Application.Workbooks.Open(SourceFileName, missing, true, missing, missing, missing,
                                missing, missing, missing, true, missing, missing, missing, missing, missing);



                Worksheet workSheet = (Worksheet)excelWorkbook.Worksheets.get_Item(1);



                int rowsint = workSheet.UsedRange.Cells.Rows.Count;
                int columnsint = workSheet.UsedRange.Cells.Columns.Count;



                Range range = workSheet.Cells.get_Range(excelColumns[ExcelRangeInitColumn] + ExcelRangeInitRow, excelColumns[ExcelItemsCount] + rowsint.ToString());

                object[,] arryItem = (object[,])range.Value2;

                excelRowsCount = range.Rows.Count;


                arry = new string[range.Rows.Count, range.Columns.Count];
                for (int i = 0; i < range.Rows.Count; i++)
                {
                    for (int j = 0; j < range.Columns.Count; j++)
                    {
                        if (arryItem[i + 1, j + 1] != null)
                        {
                            arry[i, j] = arryItem[i + 1, j + 1].ToString();
                        }

                    }
                }

            }
            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");

            foreach (Process pro in procs)
            {
                pro.Kill();
            }
            GC.Collect();

            System.Data.DataTable excelTable = new System.Data.DataTable();
            for (int i = 0; i < ExcelItemsCount; i++)
            {
                excelTable.Columns.Add(arry[0, i]);
            }

            for (int i = 1; i < excelRowsCount; i++)
            {
                DataRow dataRow = excelTable.NewRow();
                for (int j = 0; j < ExcelItemsCount; j++)
                {
                    if (arry[i, j] != null)
                    {
                        dataRow[j] = arry[i, j];
                    }
                }
                excelTable.Rows.Add(dataRow);
            }
            return excelTable;
        }

        /// <summary>
        /// 开始数据导入
        /// </summary>
        public void Import(System.Data.DataTable excelTable)
        {
            if (!ConnectionTest())
            {
                return;
            }

           

            string connStr = "Database=" + ConnectionParaDBName + ";" +
                                 "Data Source=" + ConnectionParaHostName + ";" +
                                 "User Id=" + ConnectionParaUserName + ";" +
                                 "Password=" + ConnectionParaPassWord + ";" +
                                 "pooling=false;" +
                                 "CharSet=utf8;" +
                                 "port=" + ConnectionParaPort;

            MySqlConnection mySqlConnection = new MySqlConnection(connStr);
            string sqlStr = "select * from " + ConnectionParaDBName + "." + TransferParaTableName + ";";
            MySqlDataAdapter mySqlDataAdapter = new MySqlDataAdapter(sqlStr, mySqlConnection);   
            System.Data.DataTable resultTable = new System.Data.DataTable();
            mySqlDataAdapter.Fill(resultTable);

            List<string> primeKeysList = new List<string>();
            string TransferParaPrimeKeyInExcel = TransferParaSourceKeys[TransferParaDataBaseKeys.IndexOf(TransferParaPrimeKey)];

            for (int i = 0; i < resultTable.Rows.Count; i++)
            {
                primeKeysList.Add(resultTable.Rows[i][TransferParaPrimeKey].ToString());
            }

            for (int i = 0; i < excelTable.Rows.Count; i++)
            {
                if (primeKeysList.Contains(excelTable.Rows[i][TransferParaPrimeKeyInExcel]))
                {
                    string updateSqlStr = "UPDATE " + TransferParaTableName + " SET `";
                    for (int j = 0; j < TransferParaDataBaseKeys.Count; j++)
                    {
                        if (TransferParaDataBaseKeys[j] != TransferParaPrimeKey)
                        {
                            if (j < TransferParaDataBaseKeys.Count - 1)
                            {
                                updateSqlStr += TransferParaDataBaseKeys[j] + "` = '" + excelTable.Rows[i][TransferParaSourceKeys[j]].ToString().Replace("'"," ").Replace(",", "，").Replace("\r\n"," ").Replace("\n"," ").Replace("\"", " ") + "',`";
                            }
                            else
                            {
                                updateSqlStr += TransferParaDataBaseKeys[j] + "` = '" + excelTable.Rows[i][TransferParaSourceKeys[j]].ToString().Replace("'", " ").Replace(",", "，").Replace("\r\n", " ").Replace("\n", " ").Replace("\""," ") + "'";
                            }
                            
                        }
                    }
                    updateSqlStr += " WHERE `" + TransferParaPrimeKey + "` = '" + excelTable.Rows[i][TransferParaPrimeKeyInExcel] + "'";
                    if (mySqlConnection.State == ConnectionState.Closed)
                    {
                        mySqlConnection.Open();
                    }
                    //CurrentSQLString = updateSqlStr;
                    //MySqlCommand mySqlCommand = new MySqlCommand(updateSqlStr, mySqlConnection);
                    //mySqlCommand.ExecuteNonQuery();
                    
                }
                else
                {
                    string insertStr = "INSERT INTO " + TransferParaTableName + " (";

                    string keys = "", values = "";

                    for (int j = 0; j < TransferParaDataBaseKeys.Count; j++)
                    {
                        
                        if (j < TransferParaDataBaseKeys.Count - 1)
                        {
                            keys += "`" + TransferParaDataBaseKeys[j] + "`,";
                            values += "'" + excelTable.Rows[i][TransferParaSourceKeys[j]].ToString().Replace("'", " ").Replace(",", "，").Replace("\r\n", " ").Replace("\n", " ").Replace("\"", " ") + "',";

                        }
                        else
                        {
                            keys += "`" + TransferParaDataBaseKeys[j] + "`";
                            values += "'" + excelTable.Rows[i][TransferParaSourceKeys[j]].ToString().Replace("'", " ").Replace(",", "，").Replace("\r\n", " ").Replace("\n", " ").Replace("\"", " ") + "'";
                        }

                    }
                    insertStr += keys + ") VALUES (" + values + ")";
                    if (mySqlConnection.State == ConnectionState.Closed)
                    {
                        mySqlConnection.Open();
                    }
                    CurrentSQLString = insertStr;
                    MySqlCommand mySqlCommand = new MySqlCommand(insertStr, mySqlConnection);
                    mySqlCommand.ExecuteNonQuery();
                }

            }

            mySqlConnection.Close();

        }


    


    }
}
