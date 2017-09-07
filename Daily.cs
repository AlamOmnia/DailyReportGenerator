using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Demo_Excel_Export.Helpers;
using MySql.Data.MySqlClient;
using System.Windows.Forms;

namespace Demo_Excel_Export
{
    class Daily
    {
       
        public DataTable GetIncomingDataTable(DateTime start, DateTime end, DateTime ans1, DateTime ans2)
        {
            
            #region Incoming DataTable
            DatabaseHelper.Instance.OpenConnection();
            DataTable incomingDataTable = new DataTable();
          
            // Open connection to the database
            DatabaseHelper.Instance.OpenConnection();
            MySqlCommand command = DatabaseHelper.Instance.CreateCommand();

            command.CommandText = Properties.Resources.Int_Incom_IOS_Wise;

            command.Parameters.AddWithValue("@startTime", start);
            command.Parameters.AddWithValue("@endTime", end);
            command.Parameters.AddWithValue("@AnsTime1", ans1);
            command.Parameters.AddWithValue("@AnsTime2", ans2);
            using (var dataReader = command.ExecuteReader())
            {
                incomingDataTable.Load(dataReader);
            }

            // Close the connection after finish using
            DatabaseHelper.Instance.CloseConnection();
            
            #endregion

            return incomingDataTable;
        }
        public DataTable GetOutgoingDataTable(DateTime start, DateTime end, DateTime ans1, DateTime ans2)
        {
            #region Outgoing DataTable
            DatabaseHelper.Instance.OpenConnection();
            DataTable outgoingDataTable = new DataTable();
           
            DatabaseHelper.Instance.OpenConnection();
            MySqlCommand command = DatabaseHelper.Instance.CreateCommand();

            command.CommandText = Properties.Resources.Int_Outgoing_IOS_Wise;

            command.Parameters.AddWithValue("@startTime", start);
            command.Parameters.AddWithValue("@endTime", end);
            command.Parameters.AddWithValue("@AnsTime1", ans1);
            command.Parameters.AddWithValue("@AnsTime2", ans2);
            using (var dataReader = command.ExecuteReader())
            {
                outgoingDataTable.Load(dataReader);
            }

            // Close the connection after finish using
            DatabaseHelper.Instance.CloseConnection();
            
            #endregion


            return outgoingDataTable;
        }

        public DataTable GetDomDataTable(DateTime start, DateTime end,DateTime ans1,DateTime ans2)
        {
            #region Dom DataTable
            DatabaseHelper.Instance.OpenConnection();
            DataTable domDataTable = new DataTable();

            DatabaseHelper.Instance.OpenConnection();
            MySqlCommand command = DatabaseHelper.Instance.CreateCommand();

            command.CommandText = Properties.Resources.Dom_Monthly;

            command.Parameters.AddWithValue("@startTime", start);
            command.Parameters.AddWithValue("@endTime", end);
            command.Parameters.AddWithValue("@AnsTime1",ans1);
            command.Parameters.AddWithValue("@AnsTime2", ans2);
            using (var dataReader = command.ExecuteReader())
            {
                domDataTable.Load(dataReader);
            }

            // Close the connection after finish using
            DatabaseHelper.Instance.CloseConnection();

            #endregion


            return domDataTable;
        }


        public void ExportReport(DateTime start, DateTime end,DateTime ans1,DateTime ans2)
        {
            DailyReport dailyReport = new DailyReport(GetIncomingDataTable(start,end,ans1,ans2), GetOutgoingDataTable(start,end,ans1,ans2), start.AddDays(+1).ToString("d/MM/yyyy"));
            string EXPORT_EXCEL_FILE_NAME = @"C:/Users/Omnia/Desktop/Daily International incoming and outgoingTraffic Report of Purple ICX for BTRC(" + start.AddDays(+1).ToString("dd-MMM") + ").xlsx";
            dailyReport.ExportToExcel(EXPORT_EXCEL_FILE_NAME);
            MessageBox.Show("Done!!!");
        }
        //public void ExportReportMonthly(DateTime start, DateTime end)
        //{
        //    MonthlyIncomOutReport monthlyReport = new MonthlyIncomOutReport(GetIncomingDataTable(start, end), GetOutgoingDataTable(start, end), start.ToString("MMM,yyyy"));
        //    string EXPORT_EXCEL_FILE_NAME = "D:/Billing Reports Purple/Purple/BTRC Report/Fig-2.1.xlsx";
        //    monthlyReport.ExportToExcel(EXPORT_EXCEL_FILE_NAME);
        //}
        //public void ExportReportMonthlyDom(DateTime start, DateTime end)
        //{
        //    MonthlyDomReport monthlyReport = new MonthlyDomReport(GetDomDataTable(start, end), start.ToString("MMM,yyyy"));
        //    string EXPORT_EXCEL_FILE_NAME = "D:/Billing Reports Purple/Purple/BTRC Report/Fig-2.2.xlsx";
        //    monthlyReport.ExportToExcel(EXPORT_EXCEL_FILE_NAME);
        //}
    }
}
