using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DILIMIZINZENGINLIKLERI
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();

        }
        public string sozcuk1, sozcukanlam1, sozcukcumle1, deyim1, deyimanlam1, deyimcumle1, atasozu1, atasozuyazar1,sozcukcumlekoyu1,deyimcumlekoyu1;
        public int sayi1;
        public DateTime tarih2 = DateTime.Now;
        baglantisinif con=new baglantisinif();
        //OleDbConnection excel = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Combine(Application.StartupPath, "KELIMELER.xlsx") + ";Extended Properties='Excel 12.0 Xml;HDR=YES;'");

        private void AddToStartup()
        {
            try
            {
                string appName = "DILIMIZINZENGINLIKLERI"; // Başlangıç adı
                string exePath = Application.ExecutablePath;

                RegistryKey key = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);

                // Aynı anahtar varsa tekrar yazma
                if (key.GetValue(appName) == null)
                {
                    key.SetValue(appName, exePath);
                }

                key.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

        private void Form2_Shown(object sender, EventArgs e)
        {
            OleDbConnection excel = new OleDbConnection(con.baglan);
            AddToStartup();
            try
            {
                excel.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [KELIMELER$] WHERE TARIH=@p1", excel);
                cmd.Parameters.AddWithValue("@p1", tarih2.ToString("dd.MM.yyyy"));
                OleDbDataReader dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    sozcuk1 = dr[1].ToString();
                    sozcukanlam1 = dr[2].ToString();
                    sozcukcumle1 = dr[3].ToString();
                    sozcukcumlekoyu1 = dr[4].ToString();
                    deyim1 = dr[5].ToString();
                    deyimanlam1 = dr[6].ToString();
                    deyimcumle1 = dr[7].ToString();
                    deyimcumlekoyu1 = dr[8].ToString();
                    atasozu1 = dr[9].ToString();
                    atasozuyazar1 = dr[10].ToString();
                    sayi1 = int.Parse(dr[12].ToString());

                }
                else
                {
                    Application.Exit();
                }
                excel.Close();
            }
            catch (Exception hata)
            {

                MessageBox.Show("Hata oluştu: " + hata.Message);
                Application.Exit();
            }
           

            Form1 frm1 = new Form1();
            frm1.sozcuk = sozcuk1;
            frm1.sozcukanlam = sozcukanlam1;
            frm1.sozcukcumle = sozcukcumle1;
            frm1.sozcukcumlekoyu = sozcukcumlekoyu1;
            frm1.deyim = deyim1;
            frm1.deyimanlam = deyimanlam1;
            frm1.deyimcumle = deyimcumle1;
            frm1.deyimcumlekoyu = deyimcumlekoyu1;
            frm1.atasozu = atasozu1;
            frm1.atasozuyazar = atasozuyazar1;
            frm1.sayi = sayi1;
            frm1.tarih = tarih2;

            frm1.Show();
            this.Hide();
        }
           
     
    }
}
