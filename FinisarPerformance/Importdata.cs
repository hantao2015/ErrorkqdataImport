using System;
using System.Collections.Generic;
using System.Text;
using hsopPlatform;
using HS.Platform;
using System.Windows.Forms;
using System.Collections;
namespace Finisar.Performance
{
    public class ImportFinisardata
    {
        private  CmsPassport pst = new CmsPassport();
        public Hashtable htdbcolumn = new Hashtable();
        public Hashtable htdbcolumn2 = new Hashtable();
        public Hashtable htdbcolumn2Save = new Hashtable();
        public long resid = 420130498195;
        public string strTablename = "ct420130498195";
        string _strYgnoCol = "C3_227192472953";
        long   _personlngResID = 227186227531;
        string _strPersonidCol = "C3_305737857578";
        string _strInnerUniqueCol = "C3_420148203323";
        string _strInnerUniqueColVal = "";
       
        public ImportFinisardata()
        {
            try
            {
                HS.Platform.CmsEnvironment.InitForClientApplication(Application.StartupPath, "finisar.log", false);
                pst=CmsPassport.GenerateCmsPassportBySysuser();
                htdbcolumn=CmsColumn.GetColumnsByHashtable(ref pst, resid, true, true);
                IEnumerator enums = htdbcolumn.GetEnumerator();
                CmsColumnData CCD = null;
                DictionaryEntry DE = new DictionaryEntry();
                while (enums.MoveNext())
                {
                    DE = (DictionaryEntry)enums.Current;

                    CCD = (CmsColumnData)DE.Value;

                    if (CCD.ColOriginColName != "")
                    {
                        htdbcolumn2.Add(CCD.ColOriginColName, CCD.ColName);
                    }
                }
           

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public bool logs(string filename,bool isok ,string errtxt)
        {
            try
            {
                DbStatement dbs = null;
                string strWhere = "C3_421024798140='" + filename + "'";
                string strTable = "ct421024783119";
                Hashtable ht2savelog = new Hashtable();
                int intinterval = 0;
                ht2savelog.Add("C3_421024798140", filename);
                ht2savelog.Add("C3_421024839870", errtxt);
                if (isok)
                {
                    ht2savelog.Add("C3_421024833060", "Y");
                }
                else
                {
                    ht2savelog.Add("C3_421024833060", "N");
                }
                CmsTableParam cp = new CmsTableParam();
                if (CmsDbStatement.Count(ref pst.Dbc, strTable, strWhere, ref dbs) == 0)
                {
                    CmsTable.AddRecord(ref pst, 421024783119, ref ht2savelog, ref cp);
                }
                else
                {
                    CmsTable.EditRecordsByWhere(ref pst, 421024783119, ref ht2savelog, strWhere, ref cp, true, ref intinterval);
                }
            }
            catch
            {
                return true;
            }
            return true;
        }
        public bool dealonerecord(Hashtable ht, ref string result,string filesname)
        {
            DbStatement dbs=null;
            string strYear="";
            string strPersonid="";
            string strWhere="";
            int intinterval = 0;
            try
            {
                CmsTableParam cp = new CmsTableParam();
                if (checkdata(ht, ref result))
                {



                    exht2dbht(ht, ref result);
                    strYear = Convert.ToString(htdbcolumn2Save["C3_420150922019"]);
                    strPersonid = Convert.ToString(htdbcolumn2Save["C3_420148203323"]);
                    strWhere = "C3_420150922019='" + strYear + "' and C3_420148203323=" + strPersonid;

                    htdbcolumn2Save.Add("C3_420243053195", filesname);
                    DateTime theDate; //= new DateTime(Convert.ToInt64(htdbcolumn2Save["C3_420149827099"]));

                    if (!DateTime.TryParse(Convert.ToString(htdbcolumn2Save["C3_420149827099"]), out theDate))
                    {
                        htdbcolumn2Save["C3_420149827099"] = "2012-1-1";
                    }
                    if (CmsDbStatement.Count(ref pst.Dbc, strTablename, strWhere, ref dbs) == 0)
                    {
                        CmsTable.AddRecord(ref pst, resid, ref htdbcolumn2Save, ref cp);
                        result = "添加数据:财年=" + strYear + "人员编号=" + strPersonid + "工号=" + Convert.ToString(ht["ygno_val"]) + "姓名=" + Convert.ToString(ht["name_val"]);
                    }
                    else
                    {




                        CmsTable.EditRecordsByWhere(ref pst, resid, ref htdbcolumn2Save, strWhere, ref cp, true, ref intinterval);
                        result = "修改数据:财年=" + strYear + "人员编号=" + strPersonid + "工号=" + Convert.ToString(ht["ygno_val"]) + "姓名=" + Convert.ToString(ht["name_val"]);
                    }

                }
                else
                {

                    return false ;
                }
            }
            catch (Exception ex)
            {
                result = ex.Message.ToString();
                return false ;
            }
            return true ;
        }
        private void exht2dbht(Hashtable ht, ref string result)
        {
            htdbcolumn2Save.Clear();
            string keyname="";
            string keyval="";
            try
            {
                IEnumerator enums = htdbcolumn2.Keys.GetEnumerator();
                while (enums.MoveNext())
                {
                    keyname = Convert.ToString(enums.Current);
                    keyval = Convert.ToString(htdbcolumn2[keyname]);
                    htdbcolumn2Save.Add(keyval, ht[keyname]);
                }
                htdbcolumn2Save.Add(_strInnerUniqueCol, _strInnerUniqueColVal);
            }
            catch (Exception ex)
            {
                result=ex.Message.ToString();
                return ;
            }
            
        }
        public bool checkdata(Hashtable ht, ref string errtext)
        {
            //
            try
            {
                string strYgno = Convert.ToString(ht["ygno_val"]);
              
                if (!CmsTable.IsRecordExist(ref pst, _personlngResID, _strYgnoCol, strYgno))
                {
                    errtext = "人事档案中不存在工号=" + strYgno + "的员工.";
                   return   false;
                }
                 _strInnerUniqueColVal = getpersonid(strYgno, ref errtext);
              
            }
            catch (Exception ex)
            {
                errtext=ex.Message .ToString ();
                return false;
            }
            return true;
        }
        public string getpersonid(string strYgno,ref string errtxt)
        {
            string strPersonid = "";
            
            Hashtable ht=new Hashtable ();
            try
            {
                strYgno=(Convert.ToInt64(strYgno)).ToString();
               ht= CmsTable.GetRecordHashtableByUniqueColumn(ref pst, _personlngResID, _strYgnoCol, strYgno);
               return Convert.ToString(ht[_strPersonidCol]);
            }
            catch (Exception ex)
            {
                errtxt = ex.Message.ToString();
                return  "";
            }
            return "";
        }
       
    
    }
}
