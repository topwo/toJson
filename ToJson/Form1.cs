using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;

namespace ToJson
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                this.Enabled = false;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = openFileDialog1.FileName;
                    comboBox1.DataSource = GetExcelSheets(openFileDialog1.FileName);
                    comboBox1.DisplayMember = "TABLE_NAME";
                    comboBox1.ValueMember = "TABLE_NAME";
                    button2.Enabled = true;
                    button3.Enabled = true;
                   /* if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        // Insert code to read the stream here.
                        myStream.Close();
                    }*/
                }
                this.Enabled = true;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                this.Enabled = true;
            }
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                this.Enabled = false;
                InitData();
                this.Enabled = true;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                this.Enabled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                this.Enabled = false;

                JsonSerializer serializer = JsonSerializer.Create(new JsonSerializerSettings()
                {
                    NullValueHandling = NullValueHandling.Ignore
                });
                serializer.Converters.Add(new DataTableConverter());
                StringWriter sw = new StringWriter();
                serializer.Serialize(new JsonTextWriter(sw), dataGridView1.DataSource);

                ParseJson(sw.ToString());


                this.Enabled = true;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                this.Enabled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            try
            {
                this.Enabled = false;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //获得文件路径
                    //localFilePath = saveFileDialog1.FileName.ToString();

                    //获取文件名，不带路径
                    //fileNameExt = localFilePath.Substring(localFilePath.LastIndexOf("\\") + 1);

                    //获取文件路径，不带文件名
                    //FilePath = localFilePath.Substring(0, localFilePath.LastIndexOf("\\"));

                    //给文件名前加上时间
                    //newFileName = DateTime.Now.ToString("yyyyMMdd") + fileNameExt;

                    //在文件名里加字符
                    //saveFileDialog1.FileName.Insert(1,"dameng");

                    StreamWriter sw = new StreamWriter(saveFileDialog1.FileName, false, Encoding.UTF8);
                    sw.Write(richTextBox1.Text);
                    sw.Close();
                    sw.Dispose();
                    System.Diagnostics.Process.Start(saveFileDialog1.FileName.Substring(0, saveFileDialog1.FileName.LastIndexOf("\\")));
                }
                this.Enabled = true;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                this.Enabled = true;
            }
        }

        private void InitData()
        {
            DataSet ds = GetExcelToDataTableBySheet(textBox1.Text, comboBox1.Text);
            dataGridView1.DataSource = ds.Tables[0];
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
        public  DataSet GetExcelToDataTableBySheet(string FileFullPath, string SheetName)
        {
            string strConn = GetConStr(FileFullPath);
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            OleDbDataAdapter odda = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", SheetName), conn);
            DataSet ds = new DataSet();
            odda.Fill(ds, SheetName);
            conn.Close();
            return ds;
        }
        public  DataTable GetExcelSheets(string FileFullPath)
        {
            string strConn = GetConStr(FileFullPath);
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            // 得到包含数据架构的数据表
            DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            conn.Close();
            return dt;
        }
        public DataTable FormatTable(DataTable data_table_src)
        {
            DataTable data_table_dest = new DataTable("Format_" + data_table_src.TableName); //dt.Clone();
            foreach (DataColumn dcVal in data_table_src.Columns)
            {
                DataColumn dc;
                if (!dcVal.ColumnName.Contains("'"))
                {
                    if (dcVal.ColumnName.Contains("."))
                    {
                        dc = new DataColumn(dcVal.ColumnName, typeof(float));
                    }
                    else if (dcVal.ColumnName.Contains("!"))
                    {
                        dc = new DataColumn(dcVal.ColumnName, System.Type.GetType("System.Boolean"));
                    }
                    else
                    {
                        dc = new DataColumn(dcVal.ColumnName, typeof(int));
                    }
                }
                else
                {
                    dc = new DataColumn(dcVal.ColumnName);//System.Type.GetType("System.Boolean")
                }
                data_table_dest.Columns.Add(dc);
            }
            foreach (DataRow drVal in data_table_src.Rows)
            {
                //向B中增加行
                data_table_dest.ImportRow(drVal);    //表结构（列类型）不同的话将复制失败，但不会有异常
            }
            return data_table_dest;
        }
        public void ParseJson(string json_string)
        {
            if (json_string.Contains("[") && json_string.Contains("]"))
            {
                var ja = JArray.Parse(json_string);
                richTextBox1.Text = ja.ToString();
            }
            else
            {
                var jo = JObject.Parse(json_string);
                richTextBox1.Text = jo.ToString();
            }
            if (richTextBox1.TextLength > 0)
            {
                button4.Enabled = true;
            }
        }
        public string GetConStr(string FileFullPath)
        {
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + FileFullPath + ";Extended Properties='Excel 12.0; HDR=YES; IMEX=1'"; //此連接可以操作.xls與.xlsx文件HDR=YES第一行为列名
            if (checkBox1.Checked)
            {
                strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + FileFullPath + ";Extended Properties='Excel 8.0; HDR=YES; IMEX=1'"; //此連接只能操作Excel2007之前(.xls)文件
            }
            return strConn;
        }
    }
}
