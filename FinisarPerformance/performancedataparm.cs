using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
 

namespace Finisar.Performance
{
    public class performancedataparm
    {
        public string form_header_pos = "B3";
        public string name_pos = "U2";
        public string ygno_pos = "U3";
        public string supervisor_pos = "U4";
        public string deptname_pos = "U5";
        public string Hiredate_pos = "U6";
        public string Fisyear_pos = "U7";
        public string Goal1_pos = "E12";
        public string Goal1_Weight_pos = "V12";
        public string Goal1_Self_pos = "F14";
        public string Goal1_Self_txt_pos = "R14";
        public string Goal1_Line_pos = "F16";
        public string Goal1_Line_txt_pos = "R16";
        public string Goal2_pos = "E18";
        public string Goal2_Weight_pos = "V18";
        public string Goal2_Self_pos = "F20";
        public string Goal2_Self_txt_pos = "R20";
        public string Goal2_Line_pos = "F22";
        public string Goal2_Line_txt_pos = "R22";
        public string Goal3_pos = "E24";
        public string Goal3_Weight_pos = "V24";
        public string Goal3_Self_pos = "F26";
        public string Goal3_Self_txt_pos = "R26";
        public string Goal3_Line_pos = "F28";
        public string Goal3_Line_txt_pos = "R28";
        public string Goal4_pos = "E30";
        public string Goal4_Weight_pos = "V30";
        public string Goal4_Self_pos = "F32";
        public string Goal4_Self_txt_pos = "R32";
        public string Goal4_Line_pos = "F34";
        public string Goal4_Line_txt_pos = "R34";
        public string Goal5_pos = "E36";
        public string Goal5_Weight_pos = "V36";
        public string Goal5_Self_pos = "F38";
        public string Goal5_Self_txt_pos = "R38";
        public string Goal5_Line_pos = "F40";
        public string Goal5_Line_txt_pos = "R40";
        public string total_Self_pos = "B44";
        public string total_Line_pos = "M44";
        public string total_Goal1_pos = "C53";
        public string face2facedate_pos = "C55";
        public string Good_Self_pos = "B49";
        public string Good_Line_pos = "I49";
        public string Good_Else_pos = "R49";
        public string Bad_Self_pos = "B51";
        public string Bad_Line_pos = "I51";
        public string Bad_Else_pos = "R51";
        public string form_header_val = "B3";
        public string name_val = "U2";
        public string ygno_val = "";
        public string supervisor_val = "U4";
        public string deptname_val = "U5";
        public string Hiredate_val = "U6";
        public string Fisyear_val = "";
        public string Goal1_val = "E12";
        public string Goal1_Weight_val = "V12";
        public string Goal1_Self_val = "F14";
        public string Goal1_Self_txt_val = "R14";
        public string Goal1_Line_val = "F16";
        public string Goal1_Line_txt_val = "R16";
        public string Goal2_val = "E18";
        public string Goal2_Weight_val = "V18";
        public string Goal2_Self_val = "F20";
        public string Goal2_Self_txt_val = "R20";
        public string Goal2_Line_val = "F22";
        public string Goal2_Line_txt_val = "R22";
        public string Goal3_val = "E24";
        public string Goal3_Weight_val = "V24";
        public string Goal3_Self_val = "F26";
        public string Goal3_Self_txt_val = "R26";
        public string Goal3_Line_val = "F28";
        public string Goal3_Line_txt_val = "R28";
        public string Goal4_val = "E30";
        public string Goal4_Weight_val = "V30";
        public string Goal4_Self_val = "F32";
        public string Goal4_Self_txt_val = "R32";
        public string Goal4_Line_val = "F34";
        public string Goal4_Line_txt_val = "R34";
        public string Goal5_val = "E36";
        public string Goal5_Weight_val = "V36";
        public string Goal5_Self_val = "F38";
        public string Goal5_Self_txt_val = "R38";
        public string Goal5_Line_val = "F40";
        public string Goal5_Line_txt_val = "R40";
        //public string total_Self_val = "B44";
        //public string total_Line_val = "M44";
        //public string total_Goal1_val = "C53";
        //public string face2facedate_val = "C55";
        //public string Good_Self_val = "B49";
        //public string Good_Line_val = "I49";
        //public string Good_Else_val = "R49";
        //public string Bad_Self_val = "B51";
        //public string Bad_Line_val = "I51";
        //public string Bad_Else_val = "R51";
        //public string total_Self_pos_bak = "B46";
        //public string total_Line_pos_bak = "M46";
        //public string total_Goal1_pos_bak = "C55";
        //public string face2facedate_pos_bak = "C57";
        //public string Good_Self_pos_bak = "B51";
        //public string Good_Line_pos_bak = "I51";
        //public string Good_Else_pos_bak = "R51";
        //public string Bad_Self_pos_bak = "B53";
        //public string Bad_Line_pos_bak = "I53";
        //public string Bad_Else_pos_bak = "R53";
        public string total_Self_val = "B48";
        public string total_Line_val = "M48";
        public string total_Goal1_val = "C57";
        public string face2facedate_val = "C59";
        public string Good_Self_val = "B53";
        public string Good_Line_val = "I53";
        public string Good_Else_val = "R53";
        public string Bad_Self_val = "B55";
        public string Bad_Line_val = "I55";
        public string Bad_Else_val = "R55";
        public string total_Self_pos_bak = "B50";
        public string total_Line_pos_bak = "M50";
        public string total_Goal1_pos_bak = "C59";
        public string face2facedate_pos_bak = "C61";
        public string Good_Self_pos_bak = "B55";
        public string Good_Line_pos_bak = "I55";
        public string Good_Else_pos_bak = "R55";
        public string Bad_Self_pos_bak = "B57";
        public string Bad_Line_pos_bak = "I57";
        public string Bad_Else_pos_bak = "R57";
        public Hashtable ht = new Hashtable();
        public Hashtable htval = new Hashtable();
     
        public void Getperformancedataparm(string ver)
        {

            if (ver == "300")
            {
                form_header_pos = "B3";
                name_pos = "U2";
                ygno_pos = "U3";
                supervisor_pos = "U4";
                deptname_pos = "U5";
                Hiredate_pos = "U6";
                Fisyear_pos = "U7";
                Goal1_pos = "E12";
                Goal1_Weight_pos = "V12";
                Goal1_Self_pos = "F14";
                Goal1_Self_txt_pos = "R14";
                Goal1_Line_pos = "F16";
                Goal1_Line_txt_pos = "R16";
                Goal2_pos = "E18";
                Goal2_Weight_pos = "V18";
                Goal2_Self_pos = "F20";
                Goal2_Self_txt_pos = "R20";
                Goal2_Line_pos = "F22";
                Goal2_Line_txt_pos = "R22";
                Goal3_pos = "E24";
                Goal3_Weight_pos = "V24";
                Goal3_Self_pos = "F26";
                Goal3_Self_txt_pos = "R26";
                Goal3_Line_pos = "F28";
                Goal3_Line_txt_pos = "R28";
                Goal4_pos = "E30";
                Goal4_Weight_pos = "V30";
                Goal4_Self_pos = "F32";
                Goal4_Self_txt_pos = "R32";
                Goal4_Line_pos = "F34";
                Goal4_Line_txt_pos = "R34";
                Goal5_pos = "E36";
                Goal5_Weight_pos = "V36";
                Goal5_Self_pos = "F38";
                Goal5_Self_txt_pos = "R38";
                Goal5_Line_pos = "F40";
                Goal5_Line_txt_pos = "R40";
                total_Self_pos = "B44";
                total_Line_pos = "M44";
                total_Goal1_pos = "C53";
                face2facedate_pos = "C55";
                Good_Self_pos = "B49";
                Good_Line_pos = "I49";
                Good_Else_pos = "R49";
                Bad_Self_pos = "B51";
                Bad_Line_pos = "I51";
                Bad_Else_pos = "R51";
                form_header_val = "B3";
                name_val = "U2";
                ygno_val = "";
                supervisor_val = "U4";
                deptname_val = "U5";
                Hiredate_val = "U6";
                Fisyear_val = "";
                Goal1_val = "E12";
                Goal1_Weight_val = "V12";
                Goal1_Self_val = "F14";
                Goal1_Self_txt_val = "R14";
                Goal1_Line_val = "F16";
                Goal1_Line_txt_val = "R16";
                Goal2_val = "E18";
                Goal2_Weight_val = "V18";
                Goal2_Self_val = "F20";
                Goal2_Self_txt_val = "R20";
                Goal2_Line_val = "F22";
                Goal2_Line_txt_val = "R22";
                Goal3_val = "E24";
                Goal3_Weight_val = "V24";
                Goal3_Self_val = "F26";
                Goal3_Self_txt_val = "R26";
                Goal3_Line_val = "F28";
                Goal3_Line_txt_val = "R28";
                Goal4_val = "E30";
                Goal4_Weight_val = "V30";
                Goal4_Self_val = "F32";
                Goal4_Self_txt_val = "R32";
                Goal4_Line_val = "F34";
                Goal4_Line_txt_val = "R34";
                Goal5_val = "E36";
                Goal5_Weight_val = "V36";
                Goal5_Self_val = "F38";
                Goal5_Self_txt_val = "R38";
                Goal5_Line_val = "F40";
                Goal5_Line_txt_val = "R40";
                total_Self_val = "B48";
                total_Line_val = "M48";
                total_Goal1_val = "C57";
                face2facedate_val = "C59";
                Good_Self_val = "B53";
                Good_Line_val = "I53";
                Good_Else_val = "R53";
                Bad_Self_val = "B55";
                Bad_Line_val = "I55";
                Bad_Else_val = "R55";
                total_Self_pos_bak = "B46";
                total_Line_pos_bak = "M46";
                total_Goal1_pos_bak = "C55";
                face2facedate_pos_bak = "C57";
                Good_Self_pos_bak = "B51";
                Good_Line_pos_bak = "I51";
                Good_Else_pos_bak = "R51";
                Bad_Self_pos_bak = "B53";
                Bad_Line_pos_bak = "I53";
                Bad_Else_pos_bak = "R53";
            }
            else if (ver=="100")
            {
                form_header_pos = "B3";
                name_pos = "U2";
                ygno_pos = "U3";
                supervisor_pos = "U4";
                deptname_pos = "U5";
                Hiredate_pos = "U6";
                Fisyear_pos = "U7";
                Goal1_pos = "E12";
                Goal1_Weight_pos = "V12";
                Goal1_Self_pos = "F14";
                Goal1_Self_txt_pos = "R14";
                Goal1_Line_pos = "F16";
                Goal1_Line_txt_pos = "R16";
                Goal2_pos = "E18";
                Goal2_Weight_pos = "V18";
                Goal2_Self_pos = "F20";
                Goal2_Self_txt_pos = "R20";
                Goal2_Line_pos = "F22";
                Goal2_Line_txt_pos = "R22";
                Goal3_pos = "E24";
                Goal3_Weight_pos = "V24";
                Goal3_Self_pos = "F26";
                Goal3_Self_txt_pos = "R26";
                Goal3_Line_pos = "F28";
                Goal3_Line_txt_pos = "R28";
                Goal4_pos = "E30";
                Goal4_Weight_pos = "V30";
                Goal4_Self_pos = "F32";
                Goal4_Self_txt_pos = "R32";
                Goal4_Line_pos = "F34";
                Goal4_Line_txt_pos = "R34";
                Goal5_pos = "E36";
                Goal5_Weight_pos = "V36";
                Goal5_Self_pos = "F38";
                Goal5_Self_txt_pos = "R38";
                Goal5_Line_pos = "F40";
                Goal5_Line_txt_pos = "R40";
                total_Self_pos = "B44";
                total_Line_pos = "M44";
                total_Goal1_pos = "C53";
                face2facedate_pos = "C55";
                Good_Self_pos = "B49";
                Good_Line_pos = "I49";
                Good_Else_pos = "R49";
                Bad_Self_pos = "B51";
                Bad_Line_pos = "I51";
                Bad_Else_pos = "R51";
                form_header_val = "B3";
                name_val = "U2";
                ygno_val = "";
                supervisor_val = "U4";
                deptname_val = "U5";
                Hiredate_val = "U6";
                Fisyear_val = "";
                Goal1_val = "E12";
                Goal1_Weight_val = "V12";
                Goal1_Self_val = "F14";
                Goal1_Self_txt_val = "R14";
                Goal1_Line_val = "F16";
                Goal1_Line_txt_val = "R16";
                Goal2_val = "E18";
                Goal2_Weight_val = "V18";
                Goal2_Self_val = "F20";
                Goal2_Self_txt_val = "R20";
                Goal2_Line_val = "F22";
                Goal2_Line_txt_val = "R22";
                Goal3_val = "E24";
                Goal3_Weight_val = "V24";
                Goal3_Self_val = "F26";
                Goal3_Self_txt_val = "R26";
                Goal3_Line_val = "F28";
                Goal3_Line_txt_val = "R28";
                Goal4_val = "E30";
                Goal4_Weight_val = "V30";
                Goal4_Self_val = "F32";
                Goal4_Self_txt_val = "R32";
                Goal4_Line_val = "F34";
                Goal4_Line_txt_val = "R34";
                Goal5_val = "E36";
                Goal5_Weight_val = "V36";
                Goal5_Self_val = "F38";
                Goal5_Self_txt_val = "R38";
                Goal5_Line_val = "F40";
                Goal5_Line_txt_val = "R40";
                total_Self_val = "B44";
                total_Line_val = "M44";
                total_Goal1_val = "C53";
                face2facedate_val = "C55";
                Good_Self_val = "B49";
                Good_Line_val = "I49";
                Good_Else_val = "R49";
                Bad_Self_val = "B51";
                Bad_Line_val = "I51";
                Bad_Else_val = "R51";
                total_Self_pos_bak = "B46";
                total_Line_pos_bak = "M46";
                total_Goal1_pos_bak = "C55";
                face2facedate_pos_bak = "C57";
                Good_Self_pos_bak = "B51";
                Good_Line_pos_bak = "I51";
                Good_Else_pos_bak = "R51";
                Bad_Self_pos_bak = "B53";
                Bad_Line_pos_bak = "I53";
                Bad_Else_pos_bak = "R53";
               
            }
            else if (ver == "200")
            {
                form_header_pos = "B3";
                name_pos = "U2";
                ygno_pos = "U3";
                supervisor_pos = "U4";
                deptname_pos = "U5";
                Hiredate_pos = "U6";
                Fisyear_pos = "U7";
                Goal1_pos = "E12";
                Goal1_Weight_pos = "V12";
                Goal1_Self_pos = "F14";
                Goal1_Self_txt_pos = "R14";
                Goal1_Line_pos = "F16";
                Goal1_Line_txt_pos = "R16";
                Goal2_pos = "E18";
                Goal2_Weight_pos = "V18";
                Goal2_Self_pos = "F20";
                Goal2_Self_txt_pos = "R20";
                Goal2_Line_pos = "F22";
                Goal2_Line_txt_pos = "R22";
                Goal3_pos = "E24";
                Goal3_Weight_pos = "V24";
                Goal3_Self_pos = "F26";
                Goal3_Self_txt_pos = "R26";
                Goal3_Line_pos = "F28";
                Goal3_Line_txt_pos = "R28";
                Goal4_pos = "E30";
                Goal4_Weight_pos = "V30";
                Goal4_Self_pos = "F32";
                Goal4_Self_txt_pos = "R32";
                Goal4_Line_pos = "F34";
                Goal4_Line_txt_pos = "R34";
                Goal5_pos = "E36";
                Goal5_Weight_pos = "V36";
                Goal5_Self_pos = "F38";
                Goal5_Self_txt_pos = "R38";
                Goal5_Line_pos = "F40";
                Goal5_Line_txt_pos = "R40";
                total_Self_pos = "B44";
                total_Line_pos = "M44";
                total_Goal1_pos = "C53";
                face2facedate_pos = "C55";
                Good_Self_pos = "B49";
                Good_Line_pos = "I49";
                Good_Else_pos = "R49";
                Bad_Self_pos = "B51";
                Bad_Line_pos = "I51";
                Bad_Else_pos = "R51";
                form_header_val = "B3";
                name_val = "U2";
                ygno_val = "";
                supervisor_val = "U4";
                deptname_val = "U5";
                Hiredate_val = "U6";
                Fisyear_val = "";
                Goal1_val = "E12";
                Goal1_Weight_val = "V12";
                Goal1_Self_val = "F14";
                Goal1_Self_txt_val = "R14";
                Goal1_Line_val = "F16";
                Goal1_Line_txt_val = "R16";
                Goal2_val = "E18";
                Goal2_Weight_val = "V18";
                Goal2_Self_val = "F20";
                Goal2_Self_txt_val = "R20";
                Goal2_Line_val = "F22";
                Goal2_Line_txt_val = "R22";
                Goal3_val = "E24";
                Goal3_Weight_val = "V24";
                Goal3_Self_val = "F26";
                Goal3_Self_txt_val = "R26";
                Goal3_Line_val = "F28";
                Goal3_Line_txt_val = "R28";
                Goal4_val = "E30";
                Goal4_Weight_val = "V30";
                Goal4_Self_val = "F32";
                Goal4_Self_txt_val = "R32";
                Goal4_Line_val = "F34";
                Goal4_Line_txt_val = "R34";
                Goal5_val = "E36";
                Goal5_Weight_val = "V36";
                Goal5_Self_val = "F38";
                Goal5_Self_txt_val = "R38";
                Goal5_Line_val = "F40";
                Goal5_Line_txt_val = "R40";
              
                total_Self_pos = "B46";
                total_Line_pos = "M46";
                total_Goal1_pos = "C55";
                face2facedate_pos = "C57";
                Good_Self_pos = "B51";
                Good_Line_pos = "I51";
                Good_Else_pos = "R51";
                Bad_Self_pos = "B53";
                Bad_Line_pos = "I53";
                Bad_Else_pos = "R53";
            }
            else
            {
                return ;
            }
            ht.Clear();
            ht.Add("form_header_pos", form_header_pos);
            ht.Add("name_pos", name_pos);
            ht.Add("ygno_pos", ygno_pos);
            ht.Add("supervisor_pos", supervisor_pos);
            ht.Add("deptname_pos", deptname_pos);
            ht.Add("Hiredate_pos", Hiredate_pos);
            ht.Add("Fisyear_pos", Fisyear_pos);
            ht.Add("Goal1_pos", Goal1_pos);
            ht.Add("Goal1_Weight_pos", Goal1_Weight_pos);
            ht.Add("Goal1_Self_pos", Goal1_Self_pos);
            ht.Add("Goal1_Self_txt_pos", Goal1_Self_txt_pos);
            ht.Add("Goal1_Line_pos", Goal1_Line_pos);
            ht.Add("Goal1_Line_txt_pos", Goal1_Line_txt_pos);
            ht.Add("Goal2_pos", Goal2_pos);
            ht.Add("Goal2_Weight_pos", Goal2_Weight_pos);
            ht.Add("Goal2_Self_pos", Goal2_Self_pos);
            ht.Add("Goal2_Self_txt_pos", Goal2_Self_txt_pos);
            ht.Add("Goal2_Line_pos", Goal2_Line_pos);
            ht.Add("Goal2_Line_txt_pos", Goal2_Line_txt_pos);
            ht.Add("Goal3_pos", Goal3_pos);
            ht.Add("Goal3_Weight_pos", Goal3_Weight_pos);
            ht.Add("Goal3_Self_pos", Goal3_Self_pos);
            ht.Add("Goal3_Self_txt_pos", Goal3_Self_txt_pos);
            ht.Add("Goal3_Line_pos", Goal3_Line_pos);
            ht.Add("Goal3_Line_txt_pos", Goal3_Line_txt_pos);
            ht.Add("Goal4_pos", Goal4_pos);
            ht.Add("Goal4_Weight_pos", Goal4_Weight_pos);
            ht.Add("Goal4_Self_pos", Goal4_Self_pos);
            ht.Add("Goal4_Self_txt_pos", Goal4_Self_txt_pos);
            ht.Add("Goal4_Line_pos", Goal4_Line_pos);
            ht.Add("Goal4_Line_txt_pos", Goal4_Line_txt_pos);
            ht.Add("Goal5_pos", Goal5_pos);
            ht.Add("Goal5_Weight_pos", Goal5_Weight_pos);
            ht.Add("Goal5_Self_pos", Goal5_Self_pos);
            ht.Add("Goal5_Self_txt_pos", Goal5_Self_txt_pos);
            ht.Add("Goal5_Line_pos", Goal5_Line_pos);
            ht.Add("Goal5_Line_txt_pos", Goal5_Line_txt_pos);
            ht.Add("total_Self_pos", total_Self_pos);
            ht.Add("total_Line_pos", total_Line_pos);
            ht.Add("total_Goal1_pos", total_Goal1_pos);
            ht.Add("face2facedate_pos", face2facedate_pos);
            ht.Add("Good_Self_pos", Good_Self_pos);
            ht.Add("Good_Line_pos", Good_Line_pos);
            ht.Add("Good_Else_pos", Good_Else_pos);
            ht.Add("Bad_Self_pos", Bad_Self_pos);
            ht.Add("Bad_Line_pos", Bad_Line_pos);
            ht.Add("Bad_Else_pos", Bad_Else_pos);
           
        }
      
        public performancedataparm()
        {

            //ht.Clear();
            //ht.Add("form_header_pos", form_header_pos);
            //ht.Add("name_pos", name_pos);
            //ht.Add("ygno_pos", ygno_pos);
            //ht.Add("supervisor_pos", supervisor_pos);
            //ht.Add("deptname_pos", deptname_pos);
            //ht.Add("Hiredate_pos", Hiredate_pos);
            //ht.Add("Fisyear_pos", Fisyear_pos);
            //ht.Add("Goal1_pos", Goal1_pos);
            //ht.Add("Goal1_Weight_pos", Goal1_Weight_pos);
            //ht.Add("Goal1_Self_pos", Goal1_Self_pos);
            //ht.Add("Goal1_Self_txt_pos", Goal1_Self_txt_pos);
            //ht.Add("Goal1_Line_pos", Goal1_Line_pos);
            //ht.Add("Goal1_Line_txt_pos", Goal1_Line_txt_pos);
            //ht.Add("Goal2_pos", Goal2_pos);
            //ht.Add("Goal2_Weight_pos", Goal2_Weight_pos);
            //ht.Add("Goal2_Self_pos", Goal2_Self_pos);
            //ht.Add("Goal2_Self_txt_pos", Goal2_Self_txt_pos);
            //ht.Add("Goal2_Line_pos", Goal2_Line_pos);
            //ht.Add("Goal2_Line_txt_pos", Goal2_Line_txt_pos);
            //ht.Add("Goal3_pos", Goal3_pos);
            //ht.Add("Goal3_Weight_pos", Goal3_Weight_pos);
            //ht.Add("Goal3_Self_pos", Goal3_Self_pos);
            //ht.Add("Goal3_Self_txt_pos", Goal3_Self_txt_pos);
            //ht.Add("Goal3_Line_pos", Goal3_Line_pos);
            //ht.Add("Goal3_Line_txt_pos", Goal3_Line_txt_pos);
            //ht.Add("Goal4_pos", Goal4_pos);
            //ht.Add("Goal4_Weight_pos", Goal4_Weight_pos);
            //ht.Add("Goal4_Self_pos", Goal4_Self_pos);
            //ht.Add("Goal4_Self_txt_pos", Goal4_Self_txt_pos);
            //ht.Add("Goal4_Line_pos", Goal4_Line_pos);
            //ht.Add("Goal4_Line_txt_pos", Goal4_Line_txt_pos);
            //ht.Add("Goal5_pos", Goal5_pos);
            //ht.Add("Goal5_Weight_pos", Goal5_Weight_pos);
            //ht.Add("Goal5_Self_pos", Goal5_Self_pos);
            //ht.Add("Goal5_Self_txt_pos", Goal5_Self_txt_pos);
            //ht.Add("Goal5_Line_pos", Goal5_Line_pos);
            //ht.Add("Goal5_Line_txt_pos", Goal5_Line_txt_pos);
            //ht.Add("total_Self_pos", total_Self_pos);
            //ht.Add("total_Line_pos", total_Line_pos);
            //ht.Add("total_Goal1_pos", total_Goal1_pos);
            //ht.Add("face2facedate_pos", face2facedate_pos);
            //ht.Add("Good_Self_pos", Good_Self_pos);
            //ht.Add("Good_Line_pos", Good_Line_pos);
            //ht.Add("Good_Else_pos", Good_Else_pos);
            //ht.Add("Bad_Self_pos", Bad_Self_pos);
            //ht.Add("Bad_Line_pos", Bad_Line_pos);
            //ht.Add("Bad_Else_pos", Bad_Else_pos);
           
        }
        public void addvalue(object key, object  strval)
        {

            htval.Add(key, strval);
        }
        public bool checkdata(ref string errtext)
        {
            try
            {
                string strHeader = Convert.ToString(htval["form_header_val"]);
                string strYgno = Convert.ToString(htval["ygno_val"]);
                string strFisyear = Convert.ToString(htval["Fisyear_val"]);
                string strGoal1_Weight_val = Convert.ToString(htval["Goal1_Weight_val"]);
                string strGoal2_Weight_val = Convert.ToString(htval["Goal2_Weight_val"]);
                string strGoal3_Weight_val = Convert.ToString(htval["Goal3_Weight_val"]);
                string strGoal4_Weight_val = Convert.ToString(htval["Goal4_Weight_val"]);
                string strGoal5_Weight_val = Convert.ToString(htval["Goal5_Weight_val"]);
                if (strGoal1_Weight_val==""  ) strGoal1_Weight_val="0";
                if (strGoal2_Weight_val == "") strGoal2_Weight_val = "0";
                if (strGoal3_Weight_val == "") strGoal3_Weight_val = "0";
                if (strGoal4_Weight_val == "") strGoal4_Weight_val = "0";
                if (strGoal5_Weight_val == "") strGoal5_Weight_val = "0";
                double dGoal1_Weight_val=0;
                double dGoal2_Weight_val = 0;
                double dGoal3_Weight_val = 0;
                double dGoal4_Weight_val = 0;
                double dGoal5_Weight_val = 0;

                double.TryParse(strGoal1_Weight_val, out dGoal1_Weight_val);
                double.TryParse(strGoal2_Weight_val, out dGoal2_Weight_val);
                double.TryParse(strGoal3_Weight_val, out dGoal3_Weight_val);
                double.TryParse(strGoal4_Weight_val, out dGoal4_Weight_val);
                double.TryParse(strGoal5_Weight_val, out dGoal5_Weight_val);
                htval["Goal1_Weight_val"] = Convert.ToDouble(dGoal1_Weight_val) * 100;
                htval["Goal2_Weight_val"] = Convert.ToDouble(dGoal2_Weight_val) * 100;
                htval["Goal3_Weight_val"] = Convert.ToDouble(dGoal3_Weight_val) * 100;
                htval["Goal4_Weight_val"] = Convert.ToDouble(dGoal4_Weight_val) * 100;
                htval["Goal5_Weight_val"] = Convert.ToDouble(dGoal5_Weight_val) * 100;

                if (strHeader.IndexOf("菲尼萨绩效考评表格") < 0)
                {
                    errtext = "表头:菲尼萨绩效考评表格,不存在";
                    return false;
                }
                if (strYgno == "" || strYgno == null)
                {
                    errtext = "工号:不能为空";
                    return false;
                }
                if (strFisyear == "" || strFisyear == null)
                {
                    errtext = "财年:不能为空";
                    return false;
                }

            }
            catch (Exception ex)

            {
                errtext = ex.Message.ToString();
                return false;
            }


            return true;
        }

       


    }
}