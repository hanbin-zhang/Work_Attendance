using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;

namespace WorkAttendance
{
    public partial class FMain : Form
    {

        public static bool NeedSavePrompt =false ;
        public List<String> letters;
        public FMain()
        {
            List<string> lletters = new List<string>();
            for (char i = 'A'; i <= 'Z'; i++)
            {
                lletters.Add(i.ToString());
            }
            letters = lletters;
            InitializeComponent();

            
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            spreadsheetControl1.SaveDocumentAs();
           
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if(NeedSavePrompt)
            {
                if (MessageBox.Show("当前数据尚未保存，继续生成新数据吗?.", "WorkAttendance",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;

            }
           
            spreadsheetControl1.Document.CreateNewDocument();
            NeedSavePrompt = false;

            splashScreenManager1.ShowWaitForm();

            string D1 = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string D2 = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            string SQLForMorning = string.Format("SELECT vr.realname, yg_no, vr.department, min(vr.CIO_Time), CIO_Type FROM V_RealList AS vr where CIO_Type = 1 AND CIO_Time>='{0} 0:00:00' AND CIO_Time<='{1} 23:59:59' Group by realname, cast(CIO_Time as date), CIO_Type, department,yg_no order by department, realname, cast(CIO_Time as date)", D1, D2);
            DataTable DTMorning = DAL.LoadData(SQLForMorning);
            string SQLForAfternoon = string.Format("SELECT vr.realname, yg_no, vr.department, max(vr.CIO_Time), CIO_Type FROM V_RealList AS vr where CIO_Type = -1 AND CIO_Time>='{0} 0:00:00' AND CIO_Time<='{1} 23:59:59' Group by realname, cast(CIO_Time as date), CIO_Type, department,yg_no order by department, realname, cast(CIO_Time as date)", D1, D2);
            DataTable DTAfternoon = DAL.LoadData(SQLForAfternoon);
            if (DTMorning.Rows.Count > 0)
            {
                //生成日期列表
                List<DateTime> DList = new List<DateTime>();
                DateTime d = dateTimePicker1.Value.Date;
                while (d <= dateTimePicker2.Value.Date)
                {
                    DList.Add(d);
                    d= d.AddDays(1);
                }

                if (radioGroup1.SelectedIndex == 0)
                {
                    Worksheet ws = spreadsheetControl1.Document.Worksheets[0];
                    ws.Name = "打卡记录";
                    ws.Cells[0, 0].SetValue("打卡记录 " + " 统计日期:" + string.Format("{0} 至 {1}", D1, D2));
                    ws.Cells[0, 0].Font.Bold = true;
                    ws.Cells[0, 0].Font.Size = 24;
                    ws.Cells[0, 0].Font.Color = Color.FromArgb(008080);
                    ws.FreezePanes(2, 0);
                    ws.MergeCells(ws.Range[string.Format("A1:{0}1", excelColumnConverter(3+DList.Count))]);
                    ws.MergeCells(ws.Range[string.Format("A2:{0}2", excelColumnConverter(3 + DList.Count))]);

                    ws.Cells[1, 0].SetValue(string.Format("生成时间：{0}", DateTime.Now.ToString()));

                    ws.Cells[2, 0].SetValue("姓名");
                    ws.Cells[2, 0].Font.Bold = true;
                    ws.Cells[2, 1].SetValue("部门");
                    ws.Cells[2, 1].Font.Bold = true;
                    ws.Cells[2, 2].SetValue("工号");
                    ws.Cells[2, 2].Font.Bold = true;
                    //此列起是日期
                    for (int k=0;k<DList.Count;k++)
                    {
                        ws.Cells[2, 3 + k].Font.Bold = true;
                        ws.Cells[2, 3+k].SetValue(DList[k].Day + "(" + GetCNWeekday( DList[k]) + ")");
                        if (DList[k].DayOfWeek == DayOfWeek.Saturday || DList[k].DayOfWeek ==  DayOfWeek.Sunday)
                        {   
                            ws.Cells[2, 3 + k].FillColor = Color.Green;
                        }
                    }

                    Hashtable morningHashtable = timeHashtableHelper(DTMorning);

                    Hashtable afternoonHashTable = timeHashtableHelper(DTAfternoon);

                    MessageBox.Show("aaaa");
                    // 从这里开始
                }
                if (radioGroup1.SelectedIndex == 1 || radioGroup1.SelectedIndex == 2)
                {
                    spreadsheetControl1.Document.Worksheets.Add("月度汇总");
                }

                NeedSavePrompt = true;
            }
            else
            {
                MessageBox.Show("指定日期范围没有数据.", "WorkAttendance", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            splashScreenManager1.CloseWaitForm();
        }

        private string GetCNWeekday(DateTime D)
        {
            const string Day = "日一二三四五六";

            return Day[Convert.ToInt16(D.DayOfWeek)].ToString();

        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;//确保日期2大于日期1
        }

        private void FMain_Load(object sender, EventArgs e)
        {
            dateTimePicker1_ValueChanged(null,null);//初次调用，避免日期2小于日期1

        }

        private void radioGroup1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private string excelColumnConverter(int i)
        {
            if (i > 26) return letters[i / 26 - 1] + excelColumnConverter(i-26);
            else return letters[i - 1];
        }

        private Hashtable timeHashtableHelper(DataTable DT)
        {
            Hashtable timeHashtable = new Hashtable();

            for (int i = 0; i < DT.Rows.Count; i++)
            {
                if (!timeHashtable.ContainsKey(DT.Rows[i][0]))
                {
                    List<DateTime> timelist = new List<DateTime>();
                    timelist.Add(Convert.ToDateTime(DT.Rows[i][3]));
                    timeHashtable.Add(DT.Rows[i][0], timelist);
                }
                else
                {
                    List<DateTime> timelist = (List<DateTime>)timeHashtable[DT.Rows[i][0]];
                    timelist.Add(Convert.ToDateTime(DT.Rows[i][3]));
                    timeHashtable[DT.Rows[i][0]] = timelist;
                }
            }

            return timeHashtable;
        }
    }
}
