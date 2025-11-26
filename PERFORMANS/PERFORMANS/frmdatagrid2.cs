using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PERFORMANS
{
    public partial class frmdatagrid2 : Form
    {
        public frmdatagrid2()
        {
            InitializeComponent();
        }
        public string birinciolcut;
        public string ikinciolcut;
        public string ucuncuolcut;
        public string dorduncuolcut;
        public string besinciolcut;
        public string toplampuan;
        public int sec;
        public int sinif;
        public int ders;
        baglantisinif conn= new baglantisinif();
        public int haftaal(DateTime dtPassed)
        {

            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dtPassed, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;

        }
        private void frmdatagrid2_Load(object sender, EventArgs e)
        { //ÖLÇÜT DATAGRID VİEW
            if(sec == 0)
            {
                OleDbConnection con = new OleDbConnection(conn.baglan);
                frmistatistikler frmistatistikler = new frmistatistikler();
                OleDbCommand komutpuangetir = new OleDbCommand("SELECT OGRENCIADISOYADI, SUM([1_OLCUT]) AS 'BİRİNCİ ÖLÇÜT', SUM([2_OLCUT]) AS 'İKİNCİ ÖLÇÜT', SUM([3_OLCUT]) AS 'ÜÇÜNCÜ ÖLÇÜT', SUM([4_OLCUT]) AS 'DÖRDÜNCÜ ÖLÇÜT', SUM([5_OLCUT]) AS 'BEŞİNCİ ÖLÇÜT', SUM(TOPLAMPUAN) AS 'TOPLAM PUAN' FROM TBLNOTLAR  INNER JOIN  TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID=TBLNOTLAR.OGRENCININID WHERE SINIF=@P1 GROUP BY OGRENCIADISOYADI ORDER BY  SUM(TOPLAMPUAN) DESC", con);
                komutpuangetir.Parameters.AddWithValue("@P1", sinif);
                OleDbDataAdapter da = new OleDbDataAdapter(komutpuangetir);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[1].HeaderText = birinciolcut;
                dataGridView1.Columns[2].HeaderText = ikinciolcut;
                dataGridView1.Columns[3].HeaderText = ucuncuolcut;
                dataGridView1.Columns[4].HeaderText = dorduncuolcut;
                dataGridView1.Columns[5].HeaderText = besinciolcut;
            }
            else if(sec == 1) 
                    {
                OleDbConnection con = new OleDbConnection(conn.baglan);
                frmistatistikler frmistatistikler = new frmistatistikler();
                OleDbDataAdapter da2 = new OleDbDataAdapter("SELECT SINIFAD, SUM([1_OLCUT]) AS 'BİRİNCİ ÖLÇÜT', SUM([2_OLCUT]) AS 'İKİNCİ ÖLÇÜT', SUM([3_OLCUT]) AS 'ÜÇÜNCÜ ÖLÇÜT', SUM([4_OLCUT]) AS 'DÖRDÜNCÜ ÖLÇÜT', SUM([5_OLCUT]) AS 'BEŞİNCİ ÖLÇÜT', SUM(TOPLAMPUAN) AS 'TOPLAM PUAN' FROM TBLNOTLAR  INNER JOIN  TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLNOTLAR.SINIF GROUP BY SINIFAD ORDER BY  SUM(TOPLAMPUAN) DESC ", con);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                dataGridView1.DataSource = dt2;
                dataGridView1.Columns[1].HeaderText = birinciolcut;
                dataGridView1.Columns[2].HeaderText = ikinciolcut;
                dataGridView1.Columns[3].HeaderText = ucuncuolcut;
                dataGridView1.Columns[4].HeaderText = dorduncuolcut;
                dataGridView1.Columns[5].HeaderText = besinciolcut;
            }

            else if(sec == 3)
            {
                int hafta=haftaal(DateTime.Now);
                OleDbConnection con = new OleDbConnection(conn.baglan);
                OleDbCommand olumsuzuyarihafta= new OleDbCommand("SELECT   NOTID,    BRANSADI AS 'DERS ADI',   OGRENCINUMARASI AS 'ÖĞRENCİ NUMARASI',    OGRENCIADISOYADI AS 'ÖĞRENCİADISOYADI',   SAATAD AS 'DERS SAATİ',  SINIFAD AS 'SINIFI',    TARIH,   UYARITIPI,    UYARI,   UYARIVARYOK,   HAFTA FROM    (        (         (              TBLNOTLAR                INNER JOIN TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID = TBLNOTLAR.OGRENCININID          )            INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLNOTLAR.SINIF      )        INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLNOTLAR.DERSSAATI  )   INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLNOTLAR.BRANS  WHERE    HAFTA = @H    AND UYARIVARYOK = TRUE   AND UYARITIPI = FALSE",con);
                olumsuzuyarihafta.Parameters.AddWithValue("@H", hafta);
                OleDbDataAdapter olumsuzuyarida = new OleDbDataAdapter(olumsuzuyarihafta);
                DataTable dt2 = new DataTable();
                olumsuzuyarida.Fill(dt2);
                dataGridView1.DataSource = dt2;
            }

            else if( sec == 4)
            {
                OleDbConnection con = new OleDbConnection(conn.baglan);
                OleDbCommand olumsuzuyarihepsi = new OleDbCommand("SELECT   NOTID,    BRANSADI AS 'DERS ADI',   OGRENCINUMARASI AS 'ÖĞRENCİ NUMARASI',    OGRENCIADISOYADI AS 'ÖĞRENCİADISOYADI',   SAATAD AS 'DERS SAATİ',  SINIFAD AS 'SINIFI',    TARIH,   UYARITIPI,    UYARI,   UYARIVARYOK,   HAFTA FROM    (        (         (              TBLNOTLAR                INNER JOIN TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID = TBLNOTLAR.OGRENCININID          )            INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLNOTLAR.SINIF      )        INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLNOTLAR.DERSSAATI  )   INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLNOTLAR.BRANS  WHERE     UYARIVARYOK = TRUE   AND UYARITIPI = FALSE", con);
                OleDbDataAdapter olumsuzuyarida = new OleDbDataAdapter(olumsuzuyarihepsi);
                DataTable dt2 = new DataTable();
                olumsuzuyarida.Fill(dt2);
                dataGridView1.DataSource = dt2;
            }
            else if (sec == 5)
            {
                int hafta = haftaal(DateTime.Now);
                OleDbConnection con = new OleDbConnection(conn.baglan);
                frmistatistikler frmistatistikler = new frmistatistikler();
                OleDbCommand komutpuangetir5 = new OleDbCommand("SELECT OGRENCIADISOYADI, SUM([1_OLCUT]) AS 'BİRİNCİ ÖLÇÜT', SUM([2_OLCUT]) AS 'İKİNCİ ÖLÇÜT', SUM([3_OLCUT]) AS 'ÜÇÜNCÜ ÖLÇÜT', SUM([4_OLCUT]) AS 'DÖRDÜNCÜ ÖLÇÜT', SUM([5_OLCUT]) AS 'BEŞİNCİ ÖLÇÜT', SUM(TOPLAMPUAN) AS 'TOPLAM PUAN' FROM TBLNOTLAR  INNER JOIN  TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID=TBLNOTLAR.OGRENCININID WHERE SINIF=@P1 AND HAFTA=@P2 GROUP BY OGRENCIADISOYADI ORDER BY  SUM(TOPLAMPUAN) DESC", con);
                komutpuangetir5.Parameters.AddWithValue("@P1", sinif);
                komutpuangetir5.Parameters.AddWithValue("@P2", hafta);
                OleDbDataAdapter da5 = new OleDbDataAdapter(komutpuangetir5);
                DataTable dt5 = new DataTable();
                da5.Fill(dt5);
                dataGridView1.DataSource = dt5;
                dataGridView1.Columns[1].HeaderText = birinciolcut;
                dataGridView1.Columns[2].HeaderText = ikinciolcut;
                dataGridView1.Columns[3].HeaderText = ucuncuolcut;
                dataGridView1.Columns[4].HeaderText = dorduncuolcut;
                dataGridView1.Columns[5].HeaderText = besinciolcut;
            }
            else if (sec == 6)
            {
                
                OleDbConnection con = new OleDbConnection(conn.baglan);
                frmistatistikler frmistatistikler = new frmistatistikler();
                OleDbCommand komutpuangetir6 = new OleDbCommand("SELECT OGRENCIADISOYADI, SUM([1_OLCUT]) AS 'BİRİNCİ ÖLÇÜT', SUM([2_OLCUT]) AS 'İKİNCİ ÖLÇÜT', SUM([3_OLCUT]) AS 'ÜÇÜNCÜ ÖLÇÜT', SUM([4_OLCUT]) AS 'DÖRDÜNCÜ ÖLÇÜT', SUM([5_OLCUT]) AS 'BEŞİNCİ ÖLÇÜT', SUM(TOPLAMPUAN) AS 'TOPLAM PUAN' FROM TBLNOTLAR  INNER JOIN  TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID=TBLNOTLAR.OGRENCININID WHERE SINIF=@P1 AND BRANS=@P2 GROUP BY OGRENCIADISOYADI ORDER BY  SUM(TOPLAMPUAN) DESC", con);
                komutpuangetir6.Parameters.AddWithValue("@P1", sinif);
                komutpuangetir6.Parameters.AddWithValue("@P2", ders);
                OleDbDataAdapter da6 = new OleDbDataAdapter(komutpuangetir6);
                DataTable dt6 = new DataTable();
                da6.Fill(dt6);
                dataGridView1.DataSource = dt6;
                dataGridView1.Columns[1].HeaderText = birinciolcut;
                dataGridView1.Columns[2].HeaderText = ikinciolcut;
                dataGridView1.Columns[3].HeaderText = ucuncuolcut;
                dataGridView1.Columns[4].HeaderText = dorduncuolcut;
                dataGridView1.Columns[5].HeaderText = besinciolcut;
            }




        }
    }
}
