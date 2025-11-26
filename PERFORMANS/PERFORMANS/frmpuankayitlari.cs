using Microsoft.Reporting.WinForms;
using Microsoft.ReportingServices.Diagnostics.Internal;
using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml.Linq;


namespace PERFORMANS
{
    public partial class frmpuankayitlari : Form
    {
        public frmpuankayitlari()
        {
            InitializeComponent();
        }
        public int haftaal(DateTime dtPassed)
        {

            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dtPassed, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;

        }
        public int brans;
        public string datasec;
        public int sinif;
        baglantisinif conn=new baglantisinif();
        private void frmpuankayitlari_Load(object sender, EventArgs e)
        {
            
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            //puanlanmamisderssayi
            con.Open();
            OleDbCommand puanlanmamisderssayi=new OleDbCommand("select count(*) from TBLDERSPROGRAMI where OLCDURUM=@O1",con);
            puanlanmamisderssayi.Parameters.AddWithValue("@O1", false);
            OleDbDataReader rd1=puanlanmamisderssayi.ExecuteReader();
            while (rd1.Read())
            {
                lblpuansizderssayi.Text = "Puanlanmamış Ders Sayısı: " + rd1[0].ToString();
            }
            con.Close();
            //puanlanmiş ders sayisi
            con.Open();
            OleDbCommand puanlanmisderssayi = new OleDbCommand("select count(*) from TBLDERSPROGRAMI where OLCDURUM=@O1", con);
            puanlanmisderssayi.Parameters.AddWithValue("@O1", true);
            OleDbDataReader rd2 = puanlanmisderssayi.ExecuteReader();
            while (rd2.Read())
            {
                lblpuanderssayi.Text = "Puanlanmış Ders Sayısı: " + rd2[0].ToString();
            }
            con.Close();

            //puanlamış öğretmens sayısı

            con.Close();

            //puanlamiş öğretmen sayisi
            con.Open();
            OleDbCommand puanlamisogretmensayi = new OleDbCommand("select count(*) from (select DISTINCT OGRETMEN FROM TBLDERSPROGRAMI WHERE OLCDURUM=@O1)", con);
            puanlamisogretmensayi.Parameters.AddWithValue("@O1", 1);
            OleDbDataReader rd3 = puanlamisogretmensayi.ExecuteReader();
            while (rd3.Read())
            {
                lblpuanogretmensayi.Text = "Puanlamayı Tamamlamış Öğretmen Sayısı: " + rd3[0].ToString();
            }
            con.Close();

            //puanlamamis ogretmens sayisi

            con.Open();
            OleDbCommand puanlamamisogretmensayi = new OleDbCommand("select count(*) from (SELECT DISTINCT OGRETMEN FROM TBLDERSPROGRAMI WHERE OLCDURUM=@K)", con);
            puanlamamisogretmensayi.Parameters.AddWithValue("@K", 0);
            OleDbDataReader rd4 = puanlamamisogretmensayi.ExecuteReader();
            while (rd4.Read())
            {
                lblpuansizogretmensayi.Text = "Puanlamayı Tamamlamamış Öğretmen Sayısı: " + rd4[0].ToString();
            }
            con.Close();
            

            //this.reportViewer1.RefreshReport();
            //this.reportViewer2.RefreshReport();

            
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
                
                    reportViewer2.Visible = false;
                    reportViewerpuanlanmamisdersler.Visible = false;
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbDataAdapter puankayitlari = new OleDbDataAdapter("SELECT     NOTID AS 'NOT NUMARASI',    BRANSADI AS 'DERS ADI',    SAATAD AS 'DERS SAATİ',    SINIFAD AS 'SINIF',    OGRENCIADISOYADI AS 'ÖĞRENCİNİN ADI SOYADI',    OGRENCINUMARASI AS 'ÖĞRENCİ NUMARASI',   TARIH AS 'TARİH',    [1_OLCUT],   [2_OLCUT],    [3_OLCUT],    [4_OLCUT],    [5_OLCUT],    UYARIVARYOK,    UYARI,   TOPLAMPUAN,    HAFTA  FROM    (        (            (               TBLNOTLAR                INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLNOTLAR.BRANS            )           INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLNOTLAR.DERSSAATI        )       INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLNOTLAR.SINIF    )    INNER JOIN TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID = TBLNOTLAR.OGRENCININID ", con);
            DataTable dataTable = new DataTable();
            puankayitlari.Fill(dataTable);
            dataGridView1.DataSource = dataTable;




        }

        private void btnpuanlanmamisders_Click(object sender, EventArgs e)
        {
            reportViewerpuanlanmamisdersler.Visible = false;
            reportViewer2.Visible = false;  
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand puanlanmamisdersler = new OleDbCommand("SELECT KAYITID AS 'SIRA NUMARASI', GUNLER AS 'GÜN', SINIFAD AS 'SINIF', SAATAD AS 'DERS SAATİ', BRANSADI AS 'DERS ADI', ADISOYADI AS 'ÖĞRETMEN', OLCDURUM AS 'ÖLÇÜM DURUMU' FROM ((   ( (   TBLDERSPROGRAMI INNER JOIN TBLGUN ON TBLGUN.HAFTANINGUNU = TBLDERSPROGRAMI.TARIH   ) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLDERSPROGRAMI.SINIF) INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLDERSPROGRAMI.DERSSAATI) INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLDERSPROGRAMI.DERS) INNER JOIN TBLOGRETMENLER ON TBLOGRETMENLER.OGRETMENID = TBLDERSPROGRAMI.OGRETMEN WHERE OLCDURUM=@O1 ORDER BY KAYITID ASC", con);
            puanlanmamisdersler.Parameters.AddWithValue("@O1", 0);
            OleDbDataAdapter puanlanmamisderslerdata = new OleDbDataAdapter();
            puanlanmamisderslerdata.SelectCommand = puanlanmamisdersler;
            DataTable dataTable = new DataTable();
            puanlanmamisderslerdata.Fill(dataTable);
            dataGridView1.DataSource = dataTable;
            con.Close();



        }

        private void btnpuanlanmamisdersrapor_Click(object sender, EventArgs e)
        {
            if (reportViewer2.Visible == true)
            {
                reportViewerpuanlanmamisdersler.Visible = true;
                OleDbConnection con = new OleDbConnection(conn.baglan);
                con.Open();
                OleDbCommand puanlanmamisdersler = new OleDbCommand("SELECT KAYITID AS 'SIRA NUMARASI', GUNLER AS 'GÜN', SINIFAD AS 'SINIF', SAATAD AS 'DERS SAATİ', BRANSADI AS 'DERS ADI', ADISOYADI AS 'ÖĞRETMEN', OLCDURUM AS 'ÖLÇÜM DURUMU' FROM ((   ( (   TBLDERSPROGRAMI INNER JOIN TBLGUN ON TBLGUN.HAFTANINGUNU = TBLDERSPROGRAMI.TARIH   ) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLDERSPROGRAMI.SINIF) INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLDERSPROGRAMI.DERSSAATI) INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLDERSPROGRAMI.DERS) INNER JOIN TBLOGRETMENLER ON TBLOGRETMENLER.OGRETMENID = TBLDERSPROGRAMI.OGRETMEN WHERE OLCDURUM=@O1 ORDER BY KAYITID ASC", con);
                puanlanmamisdersler.Parameters.AddWithValue("@O1", false);
                OleDbDataAdapter puanlanmamisderslerdata = new OleDbDataAdapter();

                puanlanmamisderslerdata.SelectCommand = puanlanmamisdersler;
                DataTable dt2 = new DataTable();
                puanlanmamisderslerdata.Fill(dt2);

                reportViewerpuanlanmamisdersler.LocalReport.DataSources.Clear();
                ReportDataSource rds2 = new ReportDataSource("DataSet1", dt2);
                reportViewerpuanlanmamisdersler.LocalReport.ReportPath = ".//Report1.rdlc";
                reportViewerpuanlanmamisdersler.LocalReport.DataSources.Add(rds2);
                this.reportViewerpuanlanmamisdersler.RefreshReport();
                con.Close();

            }
            else
            {
                reportViewerpuanlanmamisdersler.Visible = true;
                OleDbConnection con = new OleDbConnection(conn.baglan);
                con.Open();
                OleDbCommand puanlanmamisdersler = new OleDbCommand("SELECT KAYITID AS 'SIRA NUMARASI', GUNLER AS 'GÜN', SINIFAD AS 'SINIF', SAATAD AS 'DERS SAATİ', BRANSADI AS 'DERS ADI', ADISOYADI AS 'ÖĞRETMEN', OLCDURUM AS 'ÖLÇÜM DURUMU' FROM ((   ( (   TBLDERSPROGRAMI INNER JOIN TBLGUN ON TBLGUN.HAFTANINGUNU = TBLDERSPROGRAMI.TARIH   ) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLDERSPROGRAMI.SINIF) INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLDERSPROGRAMI.DERSSAATI) INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLDERSPROGRAMI.DERS) INNER JOIN TBLOGRETMENLER ON TBLOGRETMENLER.OGRETMENID = TBLDERSPROGRAMI.OGRETMEN WHERE OLCDURUM=@O1 ORDER BY KAYITID ASC", con);
                puanlanmamisdersler.Parameters.AddWithValue("@O1", false);
                OleDbDataAdapter puanlanmamisderslerdata = new OleDbDataAdapter();

                puanlanmamisderslerdata.SelectCommand = puanlanmamisdersler;
                DataTable dt2 = new DataTable();
                puanlanmamisderslerdata.Fill(dt2);

                reportViewerpuanlanmamisdersler.LocalReport.DataSources.Clear();
                ReportDataSource rds2 = new ReportDataSource("DataSet1", dt2);
                reportViewerpuanlanmamisdersler.LocalReport.ReportPath = ".//Report1.rdlc";
                reportViewerpuanlanmamisdersler.LocalReport.DataSources.Add(rds2);
                this.reportViewerpuanlanmamisdersler.RefreshReport();
                con.Close();
            }

               
           
           
        }

        private void raporalnot_Click(object sender, EventArgs e)
        {
            if(reportViewerpuanlanmamisdersler.Visible == true)
            {
                reportViewerpuanlanmamisdersler.Visible = false;
                reportViewer2.Visible = true;
                
                OleDbConnection con = new OleDbConnection(conn.baglan);
                con.Open();
                OleDbCommand notgor = new OleDbCommand("SELECT NOTID, BRANSADI, SAATAD, SINIFAD, OGRENCIADISOYADI, OGRENCINUMARASI, TARIH, [1_OLCUT] AS OLCUTBIR, [2_OLCUT] AS OLCUTIKI, [3_OLCUT] AS OLCUTUC, [4_OLCUT] AS OLCUTDORT, [5_OLCUT] AS OLCUTBES, UYARIVARYOK, UYARI, TOPLAMPUAN, HAFTA FROM (((TBLNOTLAR INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLNOTLAR.BRANS) INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLNOTLAR.DERSSAATI) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLNOTLAR.SINIF) INNER JOIN TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID = TBLNOTLAR.OGRENCININID ORDER BY NOTID ASC", con);
                OleDbDataAdapter notgoradp = new OleDbDataAdapter();
                notgoradp.SelectCommand = notgor;
                DataTable dt = new DataTable();
                notgoradp.Fill(dt);

                reportViewer2.LocalReport.DataSources.Clear();
                ReportDataSource rds = new ReportDataSource("DataSet1", dt);
                reportViewer2.LocalReport.ReportPath = ".//Report2.rdlc";
                reportViewer2.LocalReport.DataSources.Add(rds);
                this.reportViewer2.RefreshReport();

                con.Close();
            }
            else
            {
                reportViewerpuanlanmamisdersler.Visible = false;
                reportViewer2.Visible = true;

                OleDbConnection con = new OleDbConnection(conn.baglan);
                con.Open();
                OleDbCommand notgor = new OleDbCommand("SELECT NOTID, BRANSADI, SAATAD, SINIFAD, OGRENCIADISOYADI, OGRENCINUMARASI, TARIH, [1_OLCUT] AS OLCUTBIR, [2_OLCUT] AS OLCUTIKI, [3_OLCUT] AS OLCUTUC, [4_OLCUT] AS OLCUTDORT, [5_OLCUT] AS OLCUTBES, UYARIVARYOK, UYARI, TOPLAMPUAN, HAFTA FROM (((TBLNOTLAR INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLNOTLAR.BRANS) INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLNOTLAR.DERSSAATI) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLNOTLAR.SINIF) INNER JOIN TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID = TBLNOTLAR.OGRENCININID ORDER BY NOTID ASC", con);
                OleDbDataAdapter notgoradp = new OleDbDataAdapter();
                notgoradp.SelectCommand = notgor;
                DataTable dt = new DataTable();
                notgoradp.Fill(dt);

                reportViewer2.LocalReport.DataSources.Clear();
                ReportDataSource rds = new ReportDataSource("DataSet1", dt);
                reportViewer2.LocalReport.ReportPath = ".//Report2.rdlc";
                reportViewer2.LocalReport.DataSources.Add(rds);
                this.reportViewer2.RefreshReport();

                con.Close();
            }
        }
    }
}
