using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using RestSharp;
using RestSharp.Authenticators;
using RestSharp.Serialization.Json;
using System.Globalization;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace ReportDDC
{
    public partial class Form1 : Form
    {
        long var_inn;
        long var_kpp;
        string login_ddc;
        string pass_ddc;
        string login_dev;
        string token_dev;
        string token_ddc;
        string load = null;
        string find_guid;
        int ind1;
        int ind2;
        string guid;
        string boxid;
        string st;
        string user_st;
        string user_st1;

        public Form1()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var client = new RestClient("https://diadoc-api.kontur.ru/GetOrganization?inn=" + var_inn + "&kpp=" + var_kpp );
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "DiadocAuth, ddauth_api_client_id=Test(" + login_dev+ ")-"+token_dev+", ddauth_token=" +token_ddc+"");
            request.AddHeader("Accept", "application/json");
            request.AddHeader("Content-Type", "application/json charset=utf-8");
            IRestResponse response = client.Execute(request);            
            user_st = Convert.ToString(response.Content);
            for (int i = 0; i < user_st.Length; i++)
            {
                user_st1 = Convert.ToString(user_st[i]);
                if (user_st1.Contains(",")) { st = st + Environment.NewLine + Environment.NewLine; }

                st = st + user_st1;
            }
            var re = new Regex(@",");
            st = re.Replace(st, "");

            textBox8.Text = st;
            st = null;
            user_st = null;
            user_st1 = null;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try { var_inn = Convert.ToInt64(textBox1.Text); }
            catch   {
                MessageBox.Show("Введите числовое значение", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                var_inn = 0;
                textBox1.Clear();
                    }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            try { var_kpp = Convert.ToInt64(textBox7.Text); }
            catch
            {
                MessageBox.Show("Введите числовое значение", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                var_inn = 0;
                textBox7.Clear();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            MessageBox.Show(Convert.ToString(var_inn));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var client = new RestClient("https://diadoc-api.kontur.ru/V3/Authenticate?type=password");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("Authorization", "DiadocAuth, ddauth_api_client_id=Test("+login_dev+")-"+token_dev+"");
            request.AddHeader("Accept", "application/json");
            request.AddHeader("Content-Type", "application/json charset=utf-8");
            request.AddParameter("application/json charset=utf-8", "{'login':'"+login_ddc+"', 'password': '"+pass_ddc+"'}", ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            token_ddc = (response.Content);
            textBox6.Text = token_ddc;
            if (token_ddc.Length > 130)
            {
                if (tabControl1.TabPages.Contains(SrchCaPage) || tabControl1.TabPages.Contains(CaListPage))
                {
                    
                }
                else
                {
                    tabControl1.TabPages.Add(SrchCaPage);
                    tabControl1.TabPages.Add(CaListPage);
                }
            }
            var client1 = new RestClient("https://diadoc-api.kontur.ru//V2/GetMyUser");
            client1.Timeout = -1;
            var request1 = new RestRequest(Method.GET);
            request1.AddHeader("Authorization", "DiadocAuth, ddauth_api_client_id=Test(" + login_dev + ")-" + token_dev + ", ddauth_token=" + token_ddc + "");
            request1.AddHeader("Accept", "application/json");
            request1.AddHeader("Content-Type", "application/json charset=utf-8");
            IRestResponse response1 = client1.Execute(request1);
            user_st = Convert.ToString(response1.Content);
            for (int i = 0; i < user_st.Length; i++)
            {
                user_st1 = Convert.ToString(user_st[i]);
                if (user_st1.Contains(",")) { st = st + Environment.NewLine + Environment.NewLine; }
                
                st = st + user_st1;                
            }            
            var re = new Regex(@",");
            st = re.Replace(st, "");            

            textBox11.Text = st;
            st = null;
            user_st = null;
            user_st1 = null;
        }
                
        private void textBox2_TextChanged(object sender, EventArgs e)
        {            
            login_ddc = textBox2.Text; 
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            pass_ddc = textBox3.Text;
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            login_dev = textBox5.Text;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            token_dev = textBox4.Text;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            load = "word";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            load = "excel";
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            load = "txt";
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            var client = new RestClient("https://diadoc-api.kontur.ru/GetBox?boxId="+boxid);
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "DiadocAuth, ddauth_api_client_id=Test(" + login_dev + ")-" + token_dev + ", ddauth_token=" + token_ddc + "");
            request.AddHeader("Accept", "application/json");
            request.AddHeader("Content-Type", "application/json charset=utf-8");
            IRestResponse response = client.Execute(request);
            find_guid = Convert.ToString(response.Content);

            try
            {

                ind1 = (find_guid.IndexOf("OrgId")) + 12;
                ind2 = (find_guid.LastIndexOf("OrgId")) - 3;
                guid = find_guid.Substring(ind1, (ind2 - ind1));
                MessageBox.Show("Выполнен импорт BoxId, OrgId:" + guid, "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

                var client1 = new RestClient("https://diadoc-api.kontur.ru//V2/GetCounteragents?myOrgId=" + guid);
                client1.Timeout = -1;
                var request1 = new RestRequest(Method.GET);
                request1.AddHeader("Authorization", "DiadocAuth, ddauth_api_client_id=Test(" + login_dev + ")-" + token_dev + ", ddauth_token=" + token_ddc + "");
                request1.AddHeader("Accept", "application/json");
                request1.AddHeader("Content-Type", "application/json charset=utf-8");
                IRestResponse response1 = client1.Execute(request1);
                user_st = Convert.ToString(response1.Content);

                for (int i = 0; i < user_st.Length; i++)
                {
                    user_st1 = Convert.ToString(user_st[i]);
                    if (user_st1.Contains(",")) { st = st + Environment.NewLine + Environment.NewLine; }

                    st = st + user_st1;
                }
                var re = new Regex(@",");
                st = re.Replace(st, "");

                textBox10.Text = st;
                st = null;
                user_st = null;
                user_st1 = null;
            }

            catch { MessageBox.Show("Введите корректное значение BoxID", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly); }

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            boxid = textBox9.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (load == "word")
            {
                Word.Application wordapp = new Word.Application();
                wordapp.Visible = true;
                Word.Document worddoc;
                object wordobj = System.Reflection.Missing.Value;
                worddoc = wordapp.Documents.Add(ref wordobj);
                wordapp.Selection.TypeText(textBox10.Text);
                wordapp = null;
            }
            else if (load == "excel")
            {
                // Создаём экземпляр нашего приложения
                Excel.Application excelApp = new Excel.Application();
                // Создаём экземпляр рабочий книги Excel
                Excel.Workbook workBook;
                // Создаём экземпляр листа Excel
                Excel.Worksheet workSheet;

                workBook = excelApp.Workbooks.Add();
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                workSheet.Cells[1, 1] = textBox10.Text;

                // Открываем созданный excel-файл
                excelApp.Visible = true;
                excelApp.UserControl = true;
            }
            else if (load == "txt")
            {
                StreamWriter wrtr = new StreamWriter(@"C:\Games\test_file.txt");
                {
                    wrtr.WriteLine(textBox10.Text);
                    wrtr.Close();
                }                
                Process.Start("notepad.exe", @"C:\Games\test_file.txt");
            }
            MessageBox.Show("Выберите способ выгрузки", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (load == "word")
            {
                Word.Application wordapp = new Word.Application();
                wordapp.Visible = true;
                Word.Document worddoc;
                object wordobj = System.Reflection.Missing.Value;
                worddoc = wordapp.Documents.Add(ref wordobj);
                wordapp.Selection.TypeText(textBox8.Text);
                wordapp = null;
            }
            else if (load == "excel")
            {
                // Создаём экземпляр нашего приложения
                Excel.Application excelApp = new Excel.Application();
                // Создаём экземпляр рабочий книги Excel
                Excel.Workbook workBook;
                // Создаём экземпляр листа Excel
                Excel.Worksheet workSheet;

                workBook = excelApp.Workbooks.Add();
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                workSheet.Cells[1, 1] = textBox8.Text;
                                                
                // Открываем созданный excel-файл
                excelApp.Visible = true;
                excelApp.UserControl = true;
            }
            else if (load == "txt")
            {
                StreamWriter wrtr = new StreamWriter(@"C:\Games\test_file.txt");
                {
                    wrtr.WriteLine(textBox8.Text);
                    wrtr.Close();
                }
                Process.Start("notepad.exe", @"C:\Games\test_file.txt");
            }            
            MessageBox.Show("Выберите способ выгрузки", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            load = "word";
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            load = "excel";
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            load = "txt";
        }
        
    }

}
