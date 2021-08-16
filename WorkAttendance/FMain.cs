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

namespace WorkAttendance
{
    public partial class FMain : Form
    {

        public static bool NeedSavePrompt =false ;
        public FMain()
        {
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
            DataTable dt = DAL.LoadData(D1, D2);
            if (dt.Rows.Count > 0)
            {
                //生成日期列表
                List<DateTime> DList = new List<DateTime>();
                DateTime d = dateTimePicker1.Value.Date;
                while (d <= dateTimePicker2.Value.Date)
                {
                    DList.Add(d);
                    d= d.AddDays(1);
                }

                if (radioGroup1.SelectedIndex == 0 || radioGroup1.SelectedIndex == 2)
                {
                    Worksheet ws = spreadsheetControl1.Document.Worksheets[0];
                    ws.Name = "打卡记录";
                    ws.Cells[0, 0].SetValue("打卡记录 " + " 生成时间:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    ws.Cells[1, 0].SetValue("姓名");
                    ws.Cells[1, 1].SetValue("部门");
                    ws.Cells[1, 2].SetValue("工号");
                    //此列起是日期
                    for(int k=0;k<DList.Count;k++)
                    {

                        ws.Cells[1, 3+k].SetValue(DList[k].Day + "(" + GetCNWeekday( DList[k]) + ")");
                        if (DList[k].DayOfWeek == DayOfWeek.Saturday || DList[k].DayOfWeek ==  DayOfWeek.Sunday)
                        {
                            ws.Cells[1, 3 + k].FillColor = Color.Green;
                        }
                    }     

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                    }

                   
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
    }
}
