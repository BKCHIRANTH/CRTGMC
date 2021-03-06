using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Data.Common;
using System.Xml.Linq;
using System.Net.Mime;

namespace clientmail
{
    public partial class Form1 : Form
    {
        string path = null;
        SqlConnection con = new SqlConnection(@"Data Source=CHIRANTH\BKCHIRANTH;Initial Catalog=TimeManagement;Persist Security Info=True;User ID=sa;Password=sawasdee@23");
        public Form1()
        {
            InitializeComponent();

        }
        String year = "2010";
        String month = "January";

        private void pickyear(object sender, EventArgs e)
        {
            year = comboBox2.SelectedItem.ToString();
        }

        private void pickmonth(object sender, EventArgs e)
        {
            month = comboBox3.SelectedItem.ToString();
        }



        Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();



        private void Form1_Load(object sender, EventArgs e)
        {
            string selectedValue = comboBox1.SelectedItem as string;
            SqlCommand cmdd;
            comboBox1.Items.Clear();
            con.Open();
            cmdd = con.CreateCommand();
            cmdd.CommandType = CommandType.Text;
            cmdd.CommandText = "Execute MyProc";
            cmdd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            foreach (DataRow dr in dt.Rows)
            {
                comboBox1.Items.Add(dr["ClientName"].ToString());
                con.Close();
            }
        }
        //mail sending code
        private void button1_Click(object sender, EventArgs e)
        {



        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }
        //excel template generate code
        private void button2_Click(object sender, EventArgs e)
        {
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            int num = 1;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Client Name";

            xlWorkSheet.Cells[2, 1] = "Resource";
            xlWorkSheet.Columns[1].ColumnWidth = 15;

            xlWorkSheet.Cells[2, 2] = "      ";

            xlWorkSheet.Cells[3, 1] = "Type";
            xlWorkSheet.Cells[3, 2] = "        ";

            xlWorkSheet.Cells[4, 1] = "Dates";
            xlWorkSheet.Cells[4, 1].Font.Bold = true;

            xlWorkSheet.Cells[4, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGoldenrodYellow);

            xlWorkSheet.Cells[4, 2] = "  ";
            xlWorkSheet.Cells[4, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGoldenrodYellow);

            xlWorkSheet.Cells[4, 3] = "Description";
            xlWorkSheet.Columns[3].ColumnWidth = 20;
            xlWorkSheet.Cells[4, 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGoldenrodYellow);
            xlWorkSheet.Cells[4, 3].Font.Bold = true;

            xlWorkSheet.Cells[4, 4] = "Time in hours";
            xlWorkSheet.Cells[4, 4].Font.Bold = true;
            xlWorkSheet.Columns[4].ColumnWidth = 16;
            xlWorkSheet.Cells[4, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGoldenrodYellow);

            xlWorkSheet.Cells[5, 2] = "Week " + num;
            xlWorkSheet.Cells[5, 2].Font.Bold = true;

            xlWorkSheet.Columns[2].ColumnWidth = 18;


            xlWorkSheet.Cells[4, 2] = month + "'" + year;                    //displaying month and year




            int montha;
            if (month == "Febraury")
            {
                montha = 2;
            }
            else
            {
                montha = DateTime.ParseExact(month, "MMMM", System.Globalization.CultureInfo.InvariantCulture).Month;
            }

            int yyear = Convert.ToInt32(year);
            DateTime first = new DateTime(yyear, montha, 1);

            String firstdayname = first.DayOfWeek.ToString();

            String[] days = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };

            int m = 0, monthnumber = 1, y, yearcode;
            int code;
            if (month == "January") { monthnumber = 1; }
            else if (month == "Febraury") { monthnumber = 2; }
            else if (month == "March") { monthnumber = 3; }
            else if (month == "April") { monthnumber = 4; }
            else if (month == "May") { monthnumber = 5; }
            else if (month == "June") { monthnumber = 6; }
            else if (month == "July") { monthnumber = 7; }
            else if (month == "August") { monthnumber = 8; }
            else if (month == "September") { monthnumber = 9; }
            else if (month == "October") { monthnumber = 10; }
            else if (month == "November") { monthnumber = 11; }
            else if (month == "December") { monthnumber = 12; }

            int yr = Convert.ToInt32(year);
            y = yr % 100;

            yearcode = (y + (y / 4)) % 7;
            code = 1 + m + yearcode + 6;


            xlWorkSheet.Cells[6, 1] = "01-" + month + "-" + yyear;
            xlWorkSheet.Cells[6, 2].Value = firstdayname;

            xlWorkSheet.Cells[6, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LemonChiffon);
            int r, d, end;
            for (r = 7, d = 2; r < 34; r++, d++)
            {
                xlWorkSheet.Cells[r, 1] = d + "-" + month + "-" + y;
            }
            end = 28;
            //30 days  or 31 days calculation and leap year
            if (monthnumber % 2 == 0)
            {
                if (monthnumber == 2)
                {
                    if (((yr % 4 == 0) && (yr % 100 != 0)) || (yr % 400 == 0))
                    {
                        xlWorkSheet.Cells[r, 1] = d + "-" + month + "-" + y;
                        end = r;
                    }
                    else
                        goto next;
                }

                else if (monthnumber == 8)
                {
                    xlWorkSheet.Cells[r, 1] = d + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 1, 1] = (d + 1) + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 2, 1] = (d + 2) + "-" + month + "-" + y;
                    end = r + 2;
                }
                else if (monthnumber == 10)
                {
                    xlWorkSheet.Cells[r, 1] = d + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 1, 1] = (d + 1) + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 2, 1] = (d + 2) + "-" + month + "-" + y;
                    end = r + 2;
                }
                else if (monthnumber == 12)
                {
                    xlWorkSheet.Cells[r, 1] = d + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 1, 1] = (d + 1) + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 2, 1] = (d + 2) + "-" + month + "-" + y;
                    end = r + 2;
                }
                else
                {
                    xlWorkSheet.Cells[r, 1] = d + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 1, 1] = (d + 1) + "-" + month + "-" + y;
                    end = r + 1;
                }

            }
            else
            {
                if (monthnumber < 8)
                {
                    xlWorkSheet.Cells[r, 1] = d + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 1, 1] = (d + 1) + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 2, 1] = (d + 2) + "-" + month + "-" + y;
                    end = r + 2;
                }
                else
                {
                    xlWorkSheet.Cells[r, 1] = d + "-" + month + "-" + y;
                    xlWorkSheet.Cells[r + 1, 1] = (d + 1) + "-" + month + "-" + y;
                    end = r + 1;
                }
            }

            next:
            int k = 7, j = 0;
            //print all days .... using for loop and array!! month[]={ } ......incriment

            xlWorkSheet.Cells[k, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LemonChiffon);
            xlWorkSheet.Cells[k, 2].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

            xlWorkSheet.Cells[k, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
            xlWorkSheet.Cells[k, 4].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

            for (int i = 0; i < days.Length; i++)
            {
                if (days[i] == firstdayname)
                {
                    j = i + 1;
                    for (int n = 0; n < end - 6; n++)  //end-6
                    {
                        if (j >= days.Length) j = 0;
                        xlWorkSheet.Cells[k, 2] = days[j];
                        k++; j++;

                        xlWorkSheet.Cells[k, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LemonChiffon);
                        xlWorkSheet.Cells[k, 2].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                        xlWorkSheet.Cells[k, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
                        xlWorkSheet.Cells[k, 4].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    }
                }
            }

            int[] satnum = new int[5];
            int s = 0;
            //eliminating saturday and sundays
            if (firstdayname == "Sunday") num = 0;
            int last = 0, index = 0;
            int[] arr = new int[5];
            for (int i = 6; i < end; i++)
            {

                if (Convert.ToString(xlWorkSheet.Cells[i, 2].Value) == "Saturday")
                {
                    arr[index] = i; index++;
                    xlWorkSheet.Cells[i, 2] = null;
                    xlWorkSheet.Cells[i, 3] = "Total of the week";
                    xlWorkSheet.Cells[i, 3].Font.Bold = true;

                    xlWorkSheet.Cells[i, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    xlWorkSheet.Cells[i, 2].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    xlWorkSheet.Cells[i, 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    xlWorkSheet.Cells[i, 3].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    xlWorkSheet.Cells[i, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    xlWorkSheet.Cells[i, 4].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    satnum[s] = i;
                    s++;
                    xlWorkSheet.Cells[i, 4].Value = "=SUM(D" + (i - 5) + ":D" + (i - 1) + ")";  //formula..!
                    last = i;
                }



                if (Convert.ToString(xlWorkSheet.Cells[i, 2].Value) == "Sunday")
                {
                    xlWorkSheet.Cells[i, 2] = "Week " + (++num);
                    xlWorkSheet.Cells[i, 2].Font.Bold = true;
                    xlWorkSheet.Cells[i, 4] = "Time in hours";
                    xlWorkSheet.Cells[i, 4].Font.Bold = true;

                    xlWorkSheet.Cells[i, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    xlWorkSheet.Cells[i, 2].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    xlWorkSheet.Cells[i, 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    xlWorkSheet.Cells[i, 3].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                    xlWorkSheet.Cells[i, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    xlWorkSheet.Cells[i, 4].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    last = i;
                }

            }


            xlWorkSheet.Columns[1].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            xlWorkSheet.Columns[2].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            xlWorkSheet.Columns[3].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            xlWorkSheet.Columns[4].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

            xlWorkSheet.Cells[1, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            xlWorkSheet.Cells[1, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            xlWorkSheet.Cells[1, 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);


            //last day is less then saturday 


            if (Convert.ToString(xlWorkSheet.Cells[end, 2].Value) == "Monday" || Convert.ToString(xlWorkSheet.Cells[end, 2].Value) == "Tuesday" || Convert.ToString(xlWorkSheet.Cells[end, 2].Value) == "Wednesday" || Convert.ToString(xlWorkSheet.Cells[end, 2].Value) == "Thursday" || Convert.ToString(xlWorkSheet.Cells[end, 2].Value) == "Friday")
            {

                xlWorkSheet.Cells[end + 1, 3] = "Total Of the week";
                xlWorkSheet.Cells[end + 1, 4] = "=SUM(D" + (last + 1) + ":D" + end + ")";

                xlWorkSheet.Cells[end + 1, 3].Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                xlWorkSheet.Cells[end + 1, 3].Font.Bold = true;

                xlWorkSheet.Cells[end + 1, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                xlWorkSheet.Cells[end + 1, 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                xlWorkSheet.Cells[end + 1, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);



                xlWorkSheet.Cells[end + 2, 3] = "Total of the month";

                xlWorkSheet.Cells[end + 2, 3].Font.Bold = true;
                xlWorkSheet.Cells[end + 2, 4] = "=D" + arr[0] + "+D" + arr[1] + "+D" + arr[2] + "+D" + arr[3] + "+D" + (end + 1);


            }

            if (firstdayname == "Sunday")
            {
                xlWorkSheet.Cells[5, 2] = null;
            }

            if (Convert.ToString(xlWorkSheet.Cells[end, 2].Value) == "Saturday")
            {
                xlWorkSheet.Cells[end, 2] = null;
                xlWorkSheet.Cells[end, 3] = "Total of week";
                xlWorkSheet.Cells[end, 3].Font.Bold = true;
                xlWorkSheet.Cells[end, 4] = "=SUM(D" + (last + 1) + ":D" + (end - 1) + ")";

                xlWorkSheet.Cells[end + 2, 3] = "Total of the month";

                xlWorkSheet.Cells[end + 2, 3].Font.Bold = true;
                xlWorkSheet.Cells[end + 2, 4] = "=D" + arr[0] + "+D" + arr[1] + "+D" + arr[2] + "+D" + arr[3] + "+D" + (end);
            }


            if (Convert.ToString(xlWorkSheet.Cells[end, 2].Value) == "Sunday")
            {
                xlWorkSheet.Cells[end, 2] = null;
                xlWorkSheet.Cells[end, 3] = "Total of month";
                xlWorkSheet.Cells[end, 3].Font.Bold = true;
                xlWorkSheet.Cells[end, 4] = "=D" + arr[0] + "+D" + arr[1] + "+D" + arr[2] + "+D" + arr[3] + "+D" + (end - 1);
            }

            //SaveFileDialog save = new SaveFileDialog();

            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);
            //MessageBox.Show("SUCCESSFULLY DOWNLOADED");



            try
            {

                StringBuilder s2 = new StringBuilder();
                string selectedValue = comboBox1.SelectedItem as string;
                con.Open();
                SqlCommand cmddd;
                cmddd = con.CreateCommand();
                cmddd.CommandType = CommandType.Text;
                cmddd.CommandText = "SELECT DISTINCT ', ' + COALESCE(E.Email, '') from Resource E,Client C WHERE E.StatusId=C.StatusId and ClientName='" + selectedValue + "'";
                cmddd.ExecuteNonQuery();
                DataTable dt1 = new DataTable();
                SqlDataAdapter da1 = new SqlDataAdapter(cmddd);
                da1.Fill(dt1);

                foreach (DataRow dr1 in dt1.Rows)
                {
                    s2.Append(dr1[0].ToString());

                }
                con.Close();
                string h = s2.ToString();


                MailMessage mail = new MailMessage();
                SmtpClient server = new SmtpClient("smtp.gmail.com");

                mail.From = new MailAddress("example@gmail.com");
                mail.To.Add(h);

                mail.Subject = "THIS IS THE TEMPLATE!!!!";
                mail.Body = "Please find Attachment";

                MemoryStream stream = new MemoryStream();
                Attachment attachment = new Attachment(stream, new ContentType("text/xlsx"));
                attachment.Name = "xlApp.xlsx";  // generated excel file sent to email...
                mail.Attachments.Add(attachment);



                server.Port = 587;
                server.Credentials = new System.Net.NetworkCredential("example@gmail.com", "examplepassword");
                server.EnableSsl = true;
                server.Send(mail);
                MessageBox.Show("FILE SENT!!!!", "Response", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "FILE NOT SENT!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

