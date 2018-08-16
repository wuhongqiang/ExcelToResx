using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Resources;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string path = "";
        string fileName = "";
        OleDbConnection objConn = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                comboBox1.Items.Clear();
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = true;//该值确定是否可以选择多个文件
                dialog.Title = "请选择文件夹";
                dialog.Filter = "所有文件(*.*)|*.*";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    path = dialog.FileName;
                }
                label5.Text = path;
                String[] sheetList = GetExcelSheetNames(path);
                foreach (string str in sheetList)
                {
                    comboBox1.Items.Add(str);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        /// <summary>
        /// read excel file
        /// </summary>
        /// <param name="sheetName"></param>
        private void ReadExcel(string sheetName)
        {

            try
            {
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                DataSet ds = null;
                strExcel = "select 标签,日文,中文 from [" + sheetName + "$]";
                myCommand = new OleDbDataAdapter(strExcel, objConn);
                ds = new DataSet();
                myCommand.Fill(ds, "table1");
                dataGridView1.DataSource = ds.Tables[0].DefaultView;
                label2.Text="共"+ds.Tables[0].Rows.Count+"条记录";
            }
            finally
            {
               
            
            }
        }

       
        /// <summary> 
        /// 获取Excel工作薄中Sheet页(工作表)名集合
        /// </summary> 
        /// <param name="excelFile">Excel文件名及路径,EG:C:\Users\JK\Desktop\导入测试.xls</param> 
        /// <returns>Sheet页名称集合</returns> 
        private String[] GetExcelSheetNames(string fileName)
        {

            // 清理 
            if (objConn != null)
            {
                objConn.Close();
                objConn.Dispose();
            }
            System.Data.DataTable dt = null;
            try
            {
                string connString = string.Empty;
                string FileType = fileName.Substring(fileName.LastIndexOf("."));
                if (FileType == ".xls")
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                       "Data Source=" + fileName + ";Extended Properties=Excel 8.0;";
                else//.xlsx
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
                // 创建连接对象 
                objConn = new OleDbConnection(connString);
                // 打开数据库连接 
                objConn.Open();
                // 得到包含数据架构的数据表 
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    return null;
                }
                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;
                // 添加工作表名称到字符串数组 
                foreach (DataRow row in dt.Rows)
                {
                    string strSheetTableName = row["TABLE_NAME"].ToString();
                    //过滤无效SheetName
                    if (strSheetTableName.Contains("$") && strSheetTableName.Replace("'", "").EndsWith("$"))
                    {
                        excelSheets[i] = strSheetTableName.Substring(0, strSheetTableName.Length - 1);
                    }
                    i++;
                }
                return excelSheets;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
            finally
            {
               
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            ReadExcel(comboBox1.Text);
            textBox1.Text = comboBox1.Text + "Resource.ja-jp.resx";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            createResFile(textBox1.Text);
            MessageBox.Show(textBox1.Text+"文件生成成功");
        }
        /// <summary>
        /// create res file
        /// </summary>
        /// <param name="fileName"></param>
        private void createResFile(string fileName)
        {
            try
            {

                FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write);
                ResXResourceWriter writer = new ResXResourceWriter(fs);

                foreach (DataGridViewRow dgvr in dataGridView1.Rows)
                {
                    if (!String.IsNullOrEmpty(dgvr.Cells[0].Value.ToString()))
                    {
                        //MessageBox.Show(dgvr.Cells[0].Value.ToString() + "=" + dgvr.Cells[1].Value.ToString());
                        writer.AddResource(dgvr.Cells[0].Value.ToString(), dgvr.Cells[1].Value.ToString());
                    }
                }

                writer.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
