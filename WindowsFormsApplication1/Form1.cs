using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
 using System.Collections;
using System.Text;
 
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel=Microsoft.Office.Interop.Excel;
using System.IO;
using Finisar.Performance;
using System.Data.OleDb;
using HS.Platform;
namespace FinisarDataImport
{
    
    public partial class Form1 : Form
    {
        private string[] excelfiles;
        private ImportFinisardata Ifd = new ImportFinisardata();
        private Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        private string formsheet = "sheet1";
        private string strmainPath = "";
        private CmsPassport pst = new CmsPassport();
        DataSet dataSet1 = new DataSet();
        public Form1()
        {
            InitializeComponent();
            HS.Platform.CmsEnvironment.InitForClientApplication(Application.StartupPath, "finisar.log", false);
            pst = CmsPassport.GenerateCmsPassportBySysuser();
        }
        private string[] getrootfiles(string strPath, string ext)
        {
            string[] tmpstr2 = null;
            strmainPath = strPath;
        
            DirectoryInfo dir = new DirectoryInfo(strPath);
            if (dir.Exists)
            {
                tmpstr2 = getfiles(dir, ext);
            }
          
            return tmpstr2;
        }
        private DirectoryInfo[] getdirs(DirectoryInfo rootdir)
        {
            DirectoryInfo[] tmpdirs;
            tmpdirs=rootdir.GetDirectories();
            return tmpdirs;
        }
        private string[] getfiles(System.IO.DirectoryInfo Dir,string ext)
        {
            string[] tmpstr=null;
           
            FileInfo[] fis;
            fis= Dir.GetFiles(ext,SearchOption.AllDirectories);
            tmpstr=new string[fis.Length ];

            for (int i = 0; i < fis.Length; i++ )
            {
                tmpstr[i] = fis[i].FullName;
                
            }


            return tmpstr;
        }
        private void movefile(string filename)
        {

            //移动这个文件到备份目录下


            if (!Directory.Exists(strmainPath + "Error\\"))
            {
                Directory.CreateDirectory(strmainPath + "Error\\");
            }
            int i = strmainPath.LastIndexOf("\\");
            string Newfilename = TimeId.CurrentMilliseconds().ToString();

            string BakFileName = strmainPath + "Error\\" + Newfilename + ".xlsx";
            try
            { File.Move(filename, BakFileName); }
            catch
            { 

            }
          
        }

        private void logs(string filename,bool isok,string errtext)
        {
            Ifd.logs(filename, isok, errtext);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string[] excelfiles=null;
            string errtxt = "";
            string result = "";
            string tmpstr="";
            int okcount = 0;
            Hashtable ht = new Hashtable();
            if (txtPath.Text != "")
            {
                excelfiles= getrootfiles(txtPath.Text, "*.xls");
            }
            if (excelfiles != null)
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = excelfiles.Length;
                progressBar1.Step = 1;
               
                for (int i=0;i<excelfiles.Length ;i++)
                {

                    label2.Text = "共计表格:" + excelfiles.Length.ToString() + "当前第:" + (i+1).ToString();
                    progressBar1.PerformStep();

                    DataSet set = new DataSet();

                    if (excelfiles[i].IndexOf("$") > 0)
                    {
                        logs(excelfiles[i], true, "");
                        continue;
                    }
                    tmpstr = richTextBox1.Text;
                    tmpstr = excelfiles[i] + "\r" + tmpstr;

                    richTextBox1.Clear();
                        richTextBox1.AppendText(tmpstr);


                        if (getexceldata(excelfiles[i], ref set, ref errtxt))
                    {
       
           
                             dataSet1.Merge(set);

                             Application.DoEvents();
                    }
                    else
                    {
                        richTextBox1.AppendText(excelfiles[i] + ":fail" + errtxt + "\r");
                        logs(excelfiles[i], false, errtxt);
                        movefile(excelfiles[i]);
 
                    }
                }
                okcount=0;
                progressBar1.Value = 0;
                progressBar1.Maximum = dataSet1.Tables[0].Rows.Count;
                progressBar1.Step = 1;
                for (int i = 0; i < dataSet1.Tables[0].Rows.Count; i++)
                {
                    progressBar1.PerformStep();
                    label2.Text = "共计记录:" + dataSet1.Tables[0].Rows.Count.ToString() + "当前第:" + (i + 1).ToString();
                    DateTime Card_Date;
                    string CARD_ID;
                    DateTime CARD_TIME;
                    string DEPARTMENT;

                    string  EMP_ID;   
                    string EMP_NAME;
                    Application.DoEvents();
                    Hashtable ht2 = new Hashtable();
                    Hashtable ht2Empicflow = new Hashtable();
                    CmsTableParam cp=new CmsTableParam ();
                    ht2.Clear();
                    ht2Empicflow.Clear();
                    try
                    {
                        Card_Date = DateTime.Parse(Convert.ToString(dataSet1.Tables[0].Rows[i]["CARD_DATE"]));
                        CARD_ID = Convert.ToString(dataSet1.Tables[0].Rows[i]["CARD_ID"]);
                        CARD_TIME = DateTime.Parse(Card_Date.ToShortDateString() + ' ' + Convert.ToString(dataSet1.Tables[0].Rows[i]["CARD_TIME"]).Substring(0,8));
                        DEPARTMENT = Convert.ToString(dataSet1.Tables[0].Rows[i]["DEPARTMENT"]);
                        EMP_ID = Convert.ToString(dataSet1.Tables[0].Rows[i]["EMP_ID"]);
                        EMP_NAME = Convert.ToString(dataSet1.Tables[0].Rows[i]["EMP_NAME"]);
                        long pnid=0;
                        ht2.Add("Card_Date", Card_Date);
                        ht2.Add("CARD_ID", CARD_ID);
                        ht2.Add("CARD_TIME", CARD_TIME);
                        ht2.Add("DEPARTMENT", DEPARTMENT);
                        ht2.Add("EMP_ID", EMP_ID);
                        ht2.Add("EMP_NAME", EMP_NAME);
                        DbStatement ds = null;
                        string strWhere =" select count(*) from CT424198397948 where CARD_ID='" + CARD_ID + "' and CARD_TIME=Convert(datetime,'"+CARD_TIME.ToString()+"')";
                        if (CmsDbStatement.CountBySql(ref pst.Dbc ,strWhere,ref ds)==0)
                        {
                            CmsTable.AddRecord(ref pst,424198397948,ref ht2,ref cp);
                        }
                        string sql="select count(*) from  CT227186227531 where  C3_227192472953='"+EMP_ID.Trim()+"'";
                        if (CmsDbStatement.CountBySql(ref pst.Dbc, sql, ref ds) > 0)
                        {
                            pnid = CmsDbStatement.GetFieldLng(ref pst.Dbc, "C3_305737857578", "CT227186227531", " C3_227192472953 = '" + EMP_ID.Trim() + "'", ref ds);
                        }
                        else
                        {
                            sql = "select COUNT(*) from CT301050266340 where  CONVERT(bigint,C3_301064188869)=" + Convert.ToUInt32(CARD_ID).ToString();
                            if (CmsDbStatement.CountBySql(ref pst.Dbc, sql, ref ds) > 0)
                            {
                                pnid = CmsDbStatement.GetFieldLng(ref pst.Dbc, "C3_301064207767", "CT301050266340", " CONVERT(bigint,C3_301064188869) = " + Convert.ToUInt32(CARD_ID).ToString(), ref ds);
                            }
                            else
                            {
                                sql = "select COUNT(*) from CT301050266340 where  CONVERT(bigint,C3_422965961224)=" + Convert.ToUInt32(CARD_ID).ToString();
                                if (CmsDbStatement.CountBySql(ref pst.Dbc, sql, ref ds) > 0)
                                {
                                    pnid = CmsDbStatement.GetFieldLng(ref pst.Dbc, "C3_301064207767", "CT301050266340", " CONVERT(bigint,C3_301064188869)=" + Convert.ToUInt32(CARD_ID).ToString(), ref ds);
                                }
                                else
                                {
                                    pnid = 0;
                                }
 
                            }
                            
                        
                        }


                        ht2Empicflow.Add("PNID", pnid);
                        ht2Empicflow.Add("DATES", Card_Date.ToShortDateString());
                        ht2Empicflow.Add("Times", CARD_TIME);
                        ht2Empicflow.Add("C3_382963806129", CARD_TIME.Hour);
                        ht2Empicflow.Add("C3_382963812801", CARD_TIME.Minute);
                        ht2Empicflow.Add("C3_423706465924", CARD_TIME.Second);
                        ht2Empicflow.Add("REC_RESID", 375296681546);
                       
                        DbStatement dbs=new DbStatement (pst.Dbc );
                        sql="select count(*) from emp_icflowk where pnid="+pnid.ToString()+" and times="+"Convert(datetime,'"+CARD_TIME.ToString()+"')";
                        if (dbs.CountBySql(sql) == 0)
                        {
                            dbs.InsertRow(ref ht2Empicflow, "emp_icflowk");
                        }
                        
                        okcount++;
                        if (richTextBox1.Text .Length>8000) { richTextBox1.Clear();}
                        richTextBox1.AppendText("添加或更新刷卡记录:" + EMP_NAME+"   " + CARD_TIME.ToString());

                    }
                    catch  (Exception ex)
                    {
                     richTextBox1.AppendText(" 添加刷卡记录错误:" + ex.Message.ToString());
                    }

                }

                richTextBox1.AppendText("共计添加刷卡记录:" + okcount.ToString());
            }
        }
        private bool getexceldata(string filename, ref DataSet  set1, ref string errtxt)
        {
            Excel.Workbook wb=null;
            Excel.Worksheet ws ;
            try
            {
                string keyname = "";
                string keyvalue = "";
                string colname = "";

              
                if (isExistSheet(filename, formsheet))
                {
                   
                    wb = app.Workbooks.Open(filename);
                    ws = (Excel.Worksheet)wb.Sheets[formsheet]; }
                else
                {
                    errtxt = formsheet + "不存在";
                    return false;
                }
                OleDbConnection oledbconn1 = new OleDbConnection();
                oledbconn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" + filename + ";Extended Properties=Excel 8.0;Persist Security Info=False";
                oledbconn1.Open();
                string strSql = "Select * From [Sheet1$]";
                OleDbDataAdapter adapter=new OleDbDataAdapter ();
                adapter.SelectCommand = new  OleDbCommand(strSql, oledbconn1);
                adapter.SelectCommand.CommandType = CommandType.Text;
                
                DataSet dataSet1 =new DataSet () ;
                adapter.Fill(dataSet1);
                //richTextBox1.AppendText(excelfiles[i] + ":fail" + errtxt + "\r");
                set1 = dataSet1;
                 
               
              
            }
            catch (Exception ex)
            {
               

                errtxt = ex.ToString();
                return false;
            }
     
            return true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dr = new System.Windows.Forms.DialogResult();
            dr = folderBrowserDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                txtPath.Text = folderBrowserDialog1.SelectedPath;
            }
            else
            {
                txtPath.Text = "";
            }
        }
               
        

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        void Form1_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            app.Quit();
           
          //  app.DDETerminate();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private bool isExistSheet(string filename, string sheetname)
        {
           return true;
            OleDbConnection oledbconn1 = new OleDbConnection();
            oledbconn1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;data source=" + filename + ";Extended Properties==Excel 8.0;Persist Security Info=False";
            try
            {
                
                oledbconn1.Open();
                DataTable dt = oledbconn1.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                if (dt != null && dt.Rows.Count > 0)
                {
                    DataRow[] dr = dt.Select("TABLE_NAME = '" + sheetname + "'");
                    if (dr == null || dr.Length == 0)
                    {
                      
                        return false;
                    }
                    return true;
                }
             ;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString ());
                return false;
            }
           
           return false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }
    }
}
