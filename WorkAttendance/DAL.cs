﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkAttendance
{
    public static class DAL
    {
        public static DataTable LoadData(string D1,string D2)
        {
            DataTable DT = new DataTable();
            try
            {
                SqlConnection conn = new SqlConnection(Comm.ConnString);

                string SQL = "SELECT 1ss FROM [V_RealList] Where CIO_Time>='" + D1  + " 0:00:00' AND CIO_Time<='" + D2 + " 23:59:59'";
                using (SqlCommand sc = new SqlCommand(SQL, conn))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter(sc))
                    {
                        sda.Fill(DT);
                    }
                }
            }
            //catch (SqlException ex) 此代码捕获SQL错误，用于执行不可控语句时
             catch (Exception ex)
            {
                Comm.WriteTextLog("LoadData", ex.Message);
            }
            return DT;
        }

      
    }
}
