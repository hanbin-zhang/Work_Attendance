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

                Hashtable morningHashtable = timeHashtableHelper(DTMorning);

                Hashtable afternoonHashTable = timeHashtableHelper(DTAfternoon);

                Hashtable yg_info = yg_info_helper();

                if (radioGroup1.SelectedIndex == 0)
                {
                    Workbook wb = new Workbook();
                    Worksheet ws = wb.Worksheets[0];

                    generateHeader0(ws, DList);

                   


                    int row_number = 3;

                    foreach (String keys in morningHashtable.Keys)
                    {
                        ws.Cells[row_number, 0].SetValue(keys);
                        // 你亲爱的记得推galgame小助手：本段代码用于添加部门和员工号
                        List<string> info_list = (List<string>)yg_info[keys];
                        ws.Cells[row_number, 1].SetValue(info_list[0]);
                        ws.Cells[row_number, 2].SetValue(info_list[1]);

                        List<DateTime> morningtimelist = (List<DateTime>)morningHashtable[keys];
                        List<DateTime> afternoontimelist = (List<DateTime>)afternoonHashTable[keys];
                        int morningattend = -1;
                        int afternoonattend = -1;
                        if (morningtimelist != null)
                        {
                            morningattend = morningtimelist.Count-1;
                        }
                        if (afternoontimelist != null)
                        {
                            afternoonattend = afternoontimelist.Count-1;
                        }

                        for (int i = DList.Count-1; i>=0; i--)
                        {
                            String cell = "";
                            if (morningattend >= 0)
                            {
                                if (morningtimelist[morningattend].Day.Equals(DList[i].Day)&& morningtimelist[morningattend].Month.Equals(DList[i].Month))
                                {
                                    cell = morningtimelist[morningattend].ToShortTimeString();// + ":" + morningtimelist[morningattend].Minute;
                                    morningattend--;
                                }
                            }
                            cell = cell + "\n";
                            if (afternoonattend >= 0)
                            {
                                if (afternoontimelist[afternoonattend].Day.Equals(DList[i].Day)&& afternoontimelist[afternoonattend].Month.Equals(DList[i].Month))
                                {
                                    cell = cell + afternoontimelist[afternoonattend].ToShortTimeString();// + ":" + afternoontimelist[afternoonattend].Minute;
                                    afternoonattend--;
                                }
                                else
                                {
                                    cell = cell + " ";
                                }
                            }
                            
                            ws.Cells.RowHeight = 200;
                            ws.Cells.Alignment.WrapText = true;
                            ws.Cells[row_number, 3 + i].SetValue(cell);
                        }
                        row_number++;
                    }


                    // 从这里开始
                    spreadsheetControl1.Document.Worksheets[0].CopyFrom(ws);
                    spreadsheetControl1.Document.Worksheets[0].Name = "打卡时间表";
                }
                if (radioGroup1.SelectedIndex == 1)
                {
                    Hashtable ygStastisticInfo = new Hashtable();
                    Hashtable ygalldaysInfo = new Hashtable();

                    Workbook wb = new Workbook();
                    Worksheet ws = wb.Worksheets[0];
                    generateHeader1(ws,DList);
                    
                    foreach(string name in yg_info.Keys) {
                        Hashtable lookingUpTable = new Hashtable();

                        foreach (DateTime dt in DList)
                        {
                            lookingUpTable.Add(dt, new time_storer());
                        }

                        if (morningHashtable[name] != null)
                        {
                            foreach (DateTime dt in (List<DateTime>)morningHashtable[name])
                            {
                                time_storer ts = (time_storer)lookingUpTable[dt.Date];
                                ts.morningTime = dt;
                                //lookingUpTable[dt.Date] = ts;
                            }
                        }

                        if (afternoonHashTable[name] != null)
                        {
                            foreach (DateTime dt in (List<DateTime>)afternoonHashTable[name])
                            {
                                time_storer ts = (time_storer)lookingUpTable[dt.Date];
                                ts.afternoonTime = dt;
                                //lookingUpTable[dt.Date] = ts;
                            }
                        }

                        int attend = 0;
                        int weekends = 0;
                        int late = 0;
                        int leave_early = 0;
                        int noCheckMorning = 0;
                        int noCheckAfternoon = 0;
                        Hashtable everydayStatus = new Hashtable();

                        
                        foreach(DateTime dt in DList)
                        {
                            time_storer ts = (time_storer)lookingUpTable[dt];
                            if (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday)
                            {
                                if (ts.morningTime != default(DateTime) || ts.afternoonTime != default(DateTime))
                                {
                                    everydayStatus.Add(dt, "休息并打卡");
                                    weekends += 1;
                                    attend += 1;
                                }
                                else
                                {
                                    everydayStatus.Add(dt, "休息");
                                    weekends += 1;
                                }
                            }
                            else
                            {
                                string morning_status = "";
                                string afternoon_status = "";
                                string final_status = "";

                                if (ts.morningTime == default(DateTime) && ts.afternoonTime == default(DateTime))
                                {
                                    final_status = "旷工";
                                }

                                else {
                                    if (ts.morningTime == default(DateTime))
                                    {
                                        morning_status = "上班缺卡";
                                        noCheckMorning += 1;
                                    }
                                    else if (ts.morningTime.TimeOfDay > Convert.ToDateTime("08:30:00").TimeOfDay)
                                    {
                                        morning_status = "上班迟到";
                                        late += 1;
                                    }

                                    if (ts.afternoonTime == default(DateTime))
                                    {
                                        afternoon_status = "下班缺卡";
                                        noCheckAfternoon += 1;
                                    }
                                    else if (ts.afternoonTime.TimeOfDay < Convert.ToDateTime("17:00:00").TimeOfDay)
                                    {
                                        afternoon_status = "下班早退";
                                        leave_early += 1;
                                    }
                                }
                                if(final_status == "")
                                {
                                    attend += 1;
                                    if (morning_status != "" && afternoon_status != "")
                                    {
                                        final_status = morning_status + "，" + afternoon_status;
                                    }
                                    else if (afternoon_status != "" || morning_status != "")
                                    {
                                        final_status = morning_status + afternoon_status;
                                    }
                                    else
                                    {
                                        final_status = "正常";
                                    }
                                }

                                everydayStatus.Add(dt, final_status);
                            }
                        }
                        Hashtable stastistics = new Hashtable();

                        stastistics.Add("attend", attend);
                        stastistics.Add("weekends", weekends);
                        stastistics.Add("late", late);
                        stastistics.Add("leave_early", leave_early);
                        stastistics.Add("noCheckMorning", noCheckMorning);
                        stastistics.Add("noCheckAfternoon", noCheckAfternoon);

                        ygalldaysInfo.Add(name, everydayStatus);
                        ygStastisticInfo.Add(name, stastistics);

                    }

                    List<String> contents = new List<string> { "attend", "weekends", "late", "leave_early", "noCheckMorning", "noCheckAfternoon" };
                    int row = 4;
                    foreach(String name in ygStastisticInfo.Keys)
                    {
                        addcontents(row, 0, ws, name);
                        List<string> info_list = (List<string>)yg_info[name];
                        //考勤组
                        ws.Cells[row, 1].SetValue(info_list[0]);
                        //部门
                        ws.Cells[row, 2].SetValue(info_list[0]);
                        //工号
                        ws.Cells[row, 3].SetValue(info_list[1]);
                        //出勤，休息，迟到，早退，上班缺卡，下班缺卡
                        for (int i = 0; i < 6; i++)
                        {
                            Hashtable table = (Hashtable)ygStastisticInfo[name];
                            addcontents(row, 5 + i, ws, table[contents[i]].ToString());
                        }
                        //每日内容
                        for(int i = 0; i < DList.Count; i++)
                        {
                            Hashtable allday = (Hashtable)ygalldaysInfo[name];
                            addcontents(row, 11 + i, ws, allday[(DateTime)DList[i]].ToString());
                            if(allday[(DateTime)DList[i]] == "旷工")
                            {
                                ws.Cells[row,11+i].Fill.BackgroundColor = Color.FromArgb(0xff99cc);
                            }else if(allday[(DateTime)DList[i]] == "上班迟到")
                            {
                                ws.Cells[row, 11 + i].Fill.BackgroundColor = Color.FromArgb(0xccffcc);
                            }
                            else if (allday[(DateTime)DList[i]] == "下班早退")
                            {
                                ws.Cells[row, 11 + i].Fill.BackgroundColor = Color.FromArgb(0xffffcc);
                            }
                            else if (allday[(DateTime)DList[i]] == "下班缺卡" || allday[(DateTime)DList[i]] == "上班缺卡")
                            {
                                ws.Cells[row, 11 + i].Fill.BackgroundColor = Color.FromArgb(0xff8080);
                            }
                        }
                        row++;
                    }

                    spreadsheetControl1.Document.Worksheets[0].CopyFrom(ws);
                    spreadsheetControl1.Document.Worksheets[0].Name = "月度统计表";




                    // Workbook wb = new Workbook();
                    //Worksheet ws = wb.Worksheets[0];
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

        private Hashtable yg_info_helper()
        {
            string sqlcommand = string.Format("SELECT realname, department, yg_no FROM [Wechat1].[dbo].[V_RealList] where CIO_Time>='{0} 0:00:00' AND CIO_Time<='{1} 23:59:59' group by realname, department, yg_no;", dateTimePicker1.Value.ToString("yyyy-MM-dd"), dateTimePicker2.Value.ToString("yyyy-MM-dd"));
            DataTable dt = DAL.LoadData(sqlcommand);
            Hashtable ht = new Hashtable();
            for (int i=0;i<dt.Rows.Count;i++)
            {
                List<string> info_list = new List<string>();
                info_list.Add(dt.Rows[i][1].ToString());
                info_list.Add(dt.Rows[i][2].ToString());
                ht.Add(dt.Rows[i][0], info_list);
            }

            return ht;
        }
        //Write headers for file when 0 is been selected
        private void generateHeader0(Worksheet ws, List<DateTime>DList)
        {
            string D1 = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string D2 = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            ws.Name = "打卡记录";
            ws.Cells[0, 0].SetValue("打卡记录 " + " 统计日期:" + string.Format("{0} 至 {1}", D1, D2));
            ws.Cells[0, 0].Font.Bold = true;
            ws.Cells[0, 0].Font.Size = 24;
            ws.Cells[0, 0].Fill.BackgroundColor = Color.FromArgb(0xccffff);
            ws.Cells[0, 0].Font.Color = Color.FromArgb(0x008080);
            ws.Cells[0, 0].Borders.BottomBorder.LineStyle = BorderLineStyle.Thick;
            ws.FreezePanes(2, 0);

            CellRange cellrange1 = ws.Range[string.Format("A1:{0}1", excelColumnConverter(3 + DList.Count))];
            CellRange cellrange2 = ws.Range[string.Format("A2:{0}2", excelColumnConverter(3 + DList.Count))];
            ws.MergeCells(cellrange1);
            ws.MergeCells(cellrange2);

            //尝试更改一整个range里单元格的border
            //cellrange1.BeginUpdateFormatting().Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thick;
            //cellrange2.BeginUpdateFormatting().Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Medium;
            Formatting cellrange1formatting = cellrange1.BeginUpdateFormatting();
            cellrange1formatting.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin;
            //cellrange1formatting.Borders.RightBorder.LineStyle = BorderLineStyle.Thin;

            ws.Cells[1, 0].SetValue(string.Format("生成时间：{0}", DateTime.Now.ToString()));
            ws.Cells[1, 0].Font.Size = 14;
            ws.Cells[1, 0].FillColor = Color.FromArgb(0xccffff);
            ws.Cells[1, 0].Font.Color = Color.FromArgb(0x008080);
            ws.Cells[1, 0].Borders.TopBorder.Color = Color.Black;
            ws.Cells[1, 0].Borders.BottomBorder.Color = Color.Black;
            //cellrange2.BeginUpdateFormatting().Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Medium;
            Formatting cellrange2formatting = cellrange2.BeginUpdateFormatting();
            cellrange2formatting.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin;
            //cellrange2formatting.Borders.RightBorder.LineStyle = BorderLineStyle.Thin;

            ws.Cells[2, 0].SetValue("姓名");
            ws.Cells[2, 0].Font.Bold = true;
            ws.Cells[2, 0].FillColor = Color.FromArgb(0xffffcc);
            ws.Cells[2, 0].Borders.RightBorder.LineStyle = BorderLineStyle.Thin;

            ws.Cells[2, 1].SetValue("部门");
            ws.Cells[2, 1].Font.Bold = true;
            ws.Cells[2, 1].FillColor = Color.FromArgb(0xffffcc);
            ws.Cells[2, 1].Borders.RightBorder.LineStyle = BorderLineStyle.Thin;

            ws.Cells[2, 2].SetValue("工号");
            ws.Cells[2, 2].Font.Bold = true;
            ws.Cells[2, 2].FillColor = Color.FromArgb(0xffffcc);
            ws.Cells[2, 2].Borders.RightBorder.LineStyle = BorderLineStyle.Thin;

            //此列起是日期
            for (int k = 0; k < DList.Count; k++)
            {
                ws.Cells[2, 3 + k].Font.Bold = true;
                ws.Cells[2, 3 + k].SetValue(DList[k].Day + "(" + GetCNWeekday(DList[k]) + ")");
                if (DList[k].DayOfWeek == DayOfWeek.Saturday || DList[k].DayOfWeek == DayOfWeek.Sunday)
                {
                    ws.Cells[2, 3 + k].FillColor = Color.FromArgb(0xc9edd9);
                }
                else
                {
                    ws.Cells[2, 3 + k].FillColor = Color.FromArgb(0xffffcc);
                }
                ws.Cells[2, 3 + k].Borders.RightBorder.LineStyle = BorderLineStyle.Thin;
            }

        }
        //Write headers for file when 1 is been selected
        private void generateHeader1(Worksheet ws, List<DateTime>DList)
        {
            string D1 = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string D2 = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            ws.Name = "打卡记录";
            ws.Cells[0, 0].SetValue("打卡记录 " + " 统计日期:" + string.Format("{0} 至 {1}", D1, D2));
            ws.Cells[0, 0].Font.Bold = true;
            ws.Cells[0, 0].Font.Size = 24;
            ws.Cells[0, 0].Fill.BackgroundColor = Color.FromArgb(0xccffff);
            ws.Cells[0, 0].Font.Color = Color.FromArgb(0x008080);
            ws.Cells[0, 0].Borders.BottomBorder.LineStyle = BorderLineStyle.Thick;

            ws.Cells[1, 0].SetValue(string.Format("生成时间：{0}", DateTime.Now.ToString()));
            ws.Cells[1, 0].Font.Size = 14;
            ws.Cells[1, 0].FillColor = Color.FromArgb(0xccffff);
            ws.Cells[1, 0].Font.Color = Color.FromArgb(0x008080);
            ws.Cells[1, 0].Borders.TopBorder.Color = Color.Black;
            ws.Cells[1, 0].Borders.BottomBorder.Color = Color.Black;

            //Set headers for subtitles before "考勤结果"
            List<String> subtitles = new List<string>{ "姓名", "考勤组", "部门", "工号", "职位", "出勤天数", "休息天数", "迟到次数", "早退次数" ,"上班缺卡次数","下班缺卡次数"};
            for (int i = subtitles.Count - 1; i >= 0; i--)
            {
                addcontents(3, i, ws, subtitles[i]);
                CellRange range = ws[string.Format("{0}3:{0}4",excelColumnConverter(i+1))];
                range.Merge();
            }
            ws.FreezePanes(3, 0);

            //"考勤结果" Title
            addcontents(2, subtitles.Count, ws, "考勤结果");
            ws.Cells[2, subtitles.Count].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            ws.Cells[2, subtitles.Count].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            //Merge "考勤结果" cell
            CellRange kq_range = ws[string.Format("{0}3:{1}3",excelColumnConverter(subtitles.Count+1),excelColumnConverter(subtitles.Count + DList.Count))];
            kq_range.Merge();

            //output dates
            for(int i = subtitles.Count; i < subtitles.Count + DList.Count; i++)
            {
                ws.Cells[3,i].Font.Bold = true;
                ws.Cells[3,i].SetValue(DList[i-subtitles.Count].Day + "(" + GetCNWeekday(DList[i-subtitles.Count]) + ")");
                if (DList[i-subtitles.Count].DayOfWeek == DayOfWeek.Saturday || DList[i-subtitles.Count].DayOfWeek == DayOfWeek.Sunday)
                {
                    ws.Cells[3,i].FillColor = Color.FromArgb(0xc9edd9);
                }
                else
                {
                    ws.Cells[3,i].FillColor = Color.FromArgb(0xffffcc);
                }
                ws.Cells[3,i].Borders.RightBorder.LineStyle = BorderLineStyle.Thin;
            }
        }

        // Function That can write contets into specific cell with Bolding;
        private void addcontents(int row, int col, Worksheet ws,String contents)
        {
            ws.Cells[row, col].SetValue(contents);
            ws.Cells[row, col].Font.Bold = true;
            ws.Cells[row, col].Borders.RightBorder.LineStyle = BorderLineStyle.Thin;
        }
        
    }
}
