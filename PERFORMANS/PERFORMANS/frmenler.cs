using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PERFORMANS
{
    public partial class frmenler : Form
    {
        public frmenler()
        {
            InitializeComponent();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        baglantisinif conn=new baglantisinif();
        int ogrenciid;
        int ogrenciid2;
        int birhafolcutogrid;
        int birhafolcutogrid2;
        int ikihafolcutogrid;
        int ikihafolcutogrid2;
        int uchafolcutogrid;
        int uchafolcutogrid2;
        int dorthafolcutogrid;
        int dorthafolcutogrid2;
        int beshafolcutogrid;
        int beshafolcutogrid2;

        int topdonogryuk;
        int topdonogrdus;

        public int GetWeekNumber(DateTime dtPassed)
        {
            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dtPassed, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;

        }
        public int haftaal(DateTime dtPassed)
        {

            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dtPassed, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;

        }
        void haftanineniyileri()
        {
            int hafta = haftaal(DateTime.Now);
            groupBox1.Text = hafta + ". HAFTANIN ENLERİ ";
            OleDbConnection con = new OleDbConnection(conn.baglan);

            //HAFTANIN BİRİNCİ ÖĞRENCİSİ
            con.Open();
            OleDbCommand komuttopogrenciid = new OleDbCommand("select TOP 1 OGRENCININID,  SUM(TOPLAMPUAN) FROM TBLNOTLAR WHERE HAFTA=@P1  GROUP BY OGRENCININID ORDER BY SUM(TOPLAMPUAN) DESC", con);
            komuttopogrenciid.Parameters.AddWithValue("@P1", hafta);
            OleDbDataReader rd1 = komuttopogrenciid.ExecuteReader();
            while (rd1.Read())
            {
                ogrenciid = int.Parse(rd1[0].ToString());

            }
            con.Close();
            con.Open();
            OleDbCommand komuttopogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + ogrenciid, con);
            OleDbDataReader rd2 = komuttopogrenciadfoto.ExecuteReader();
            while (rd2.Read())
            {
                lbltophaf.Text = rd2[0].ToString() + " " + rd2[1].ToString();
                pckhaftopogren.ImageLocation = rd2[2].ToString();

            }
            con.Close();

            //haftanın birinci sinifi
            con.Open();
            OleDbCommand komuttopsinif = new OleDbCommand("select TOP 1 SINIFAD, SUM(TOPLAMPUAN) FROM TBLNOTLAR INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLNOTLAR.SINIF  WHERE HAFTA=@H1 GROUP BY SINIFAD ORDER BY SUM(TOPLAMPUAN) DESC", con);
            komuttopsinif.Parameters.AddWithValue("@H1", hafta);
            OleDbDataReader komuttopsinifrd = komuttopsinif.ExecuteReader();
            while (komuttopsinifrd.Read())
            {
                lbltophafsinif.Text = (komuttopsinifrd[0].ToString()).ToUpper();
                label20.Text = "TOPLAM PUAN";
                label21.Text = (komuttopsinifrd[1].ToString()).ToUpper();
            }

            con.Close();
            //haftanın ikinci öğrencisi ogrencisi
            con.Open();
            OleDbCommand komuthaf1olcut = new OleDbCommand("SELECT    TOP 2 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR WHERE    HAFTA = @H1 GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            komuthaf1olcut.Parameters.AddWithValue("@H1", hafta);
            OleDbDataReader komuthafbirolcutrd = komuthaf1olcut.ExecuteReader();
            while (komuthafbirolcutrd.Read())
            {
                birhafolcutogrid = int.Parse(komuthafbirolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutbirhafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + birhafolcutogrid, con);
            OleDbDataReader rd3 = komutbirhafogrenciadfoto.ExecuteReader();
            while (rd3.Read())
            {
                lblbirolcut.Text = rd3[0].ToString() + " " + rd3[1].ToString();
                pckolc1haf.ImageLocation = rd3[2].ToString();

            }
            con.Close();
            //haftanın üçüncü öğrencisi
            con.Open();
            OleDbCommand komuthaf2olcut = new OleDbCommand("SELECT    TOP 3 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR WHERE    HAFTA = @H1 GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            komuthaf2olcut.Parameters.AddWithValue("@H1", hafta);
            OleDbDataReader komuthafikiolcutrd = komuthaf2olcut.ExecuteReader();
            while (komuthafikiolcutrd.Read())
            {
                ikihafolcutogrid = int.Parse(komuthafikiolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutikihafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + ikihafolcutogrid, con);
            OleDbDataReader rd4 = komutikihafogrenciadfoto.ExecuteReader();
            while (rd4.Read())
            {
                lblikiolcut.Text = rd4[0].ToString() + " " + rd4[1].ToString();
                pckolc2haf.ImageLocation = rd4[2].ToString();

            }
            con.Close();

            //haftanın dördüncü  öğrencisi
            con.Open();
            OleDbCommand komuthaf3olcut = new OleDbCommand("SELECT    TOP 4 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR WHERE    HAFTA = @H1 GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            komuthaf3olcut.Parameters.AddWithValue("@H1", hafta);
            OleDbDataReader komuthafucolcutrd = komuthaf3olcut.ExecuteReader();
            while (komuthafucolcutrd.Read())
            {
                uchafolcutogrid = int.Parse(komuthafucolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutuchafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + uchafolcutogrid, con);
            OleDbDataReader rd5 = komutuchafogrenciadfoto.ExecuteReader();
            while (rd5.Read())
            {
                lblucolcut.Text = rd5[0].ToString() + " " + rd5[1].ToString();
                pckolc3haf.ImageLocation = rd5[2].ToString();

            }
            con.Close();

            //haftanın beşinci  öğrencisi
            con.Open();
            OleDbCommand komuthaf4olcut = new OleDbCommand("SELECT    TOP 5 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR WHERE    HAFTA = @H1 GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            komuthaf4olcut.Parameters.AddWithValue("@H1", hafta);
            OleDbDataReader komuthafdortolcutrd = komuthaf4olcut.ExecuteReader();
            while (komuthafdortolcutrd.Read())
            {
                dorthafolcutogrid = int.Parse(komuthafdortolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutdorthafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + dorthafolcutogrid, con);
            OleDbDataReader rd6 = komutdorthafogrenciadfoto.ExecuteReader();
            while (rd6.Read())
            {
                lbldortolcut.Text = rd6[0].ToString() + " " + rd6[1].ToString();
                pckolc4haf.ImageLocation = rd6[2].ToString();

            }
            con.Close();
            //haftanın 6. öğrencisi
            con.Open();
            OleDbCommand komuthaf5olcut = new OleDbCommand("SELECT    TOP 6 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR WHERE    HAFTA = @H1 GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            komuthaf5olcut.Parameters.AddWithValue("@H1", hafta);
            OleDbDataReader komuthafbesolcutrd = komuthaf5olcut.ExecuteReader();
            while (komuthafbesolcutrd.Read())
            {
                beshafolcutogrid = int.Parse(komuthafbesolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutbeshafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + beshafolcutogrid, con);
            OleDbDataReader rd7 = komutbeshafogrenciadfoto.ExecuteReader();
            while (rd7.Read())
            {
                lblbesolcut.Text = rd7[0].ToString() + " " + rd7[1].ToString();
                pckolc5haf.ImageLocation = rd7[2].ToString();

            }
            con.Close();
        }
        void donemineniyileri()
        {
            int hafta = haftaal(DateTime.Now);
            //groupBox2.Text = hafta + ". HAFTANIN EN DÜŞÜK PUANLI ÖĞRENCİLERİ";
            OleDbConnection con = new OleDbConnection(conn.baglan);

            //dönemin birinci öğrencisi
            con.Open();
            OleDbCommand komuttopogrenciid = new OleDbCommand("select TOP 1 OGRENCININID,  SUM(TOPLAMPUAN) FROM TBLNOTLAR   GROUP BY OGRENCININID ORDER BY SUM(TOPLAMPUAN) DESC", con);
            
            OleDbDataReader rd1 = komuttopogrenciid.ExecuteReader();
            while (rd1.Read())
            {
                ogrenciid2 = int.Parse(rd1[0].ToString());

            }
            con.Close();
            con.Open();
            OleDbCommand komuttopogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + ogrenciid2, con);
            OleDbDataReader rd2 = komuttopogrenciadfoto.ExecuteReader();
            while (rd2.Read())
            {
                lblhaftopogr.Text = rd2[0].ToString() + " " + rd2[1].ToString();
                pckhaftopogr.ImageLocation = rd2[2].ToString();

            }
            con.Close();

            //dönemin birinci sınıfı
            con.Open();
            OleDbCommand komuttopsinif = new OleDbCommand("select TOP 1 SINIFAD, SUM(TOPLAMPUAN) FROM TBLNOTLAR INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLNOTLAR.SINIF   GROUP BY SINIFAD ORDER BY SUM(TOPLAMPUAN) DESC", con);
            
            OleDbDataReader komuttopsinifrd = komuttopsinif.ExecuteReader();
            while (komuttopsinifrd.Read())
            {
                lblhaftopsin.Text = (komuttopsinifrd[0].ToString()).ToUpper();
                
                lbltopsinpuan.Text = (komuttopsinifrd[1].ToString()).ToUpper();
            }

            con.Close();
            //dönemin ikinci öğrencisi
            con.Open();
            OleDbCommand komuthaf1olcut = new OleDbCommand("SELECT    TOP 2 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR  GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            
            OleDbDataReader komuthafbirolcutrd = komuthaf1olcut.ExecuteReader();
            while (komuthafbirolcutrd.Read())
            {
                birhafolcutogrid2 = int.Parse(komuthafbirolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutbirhafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + birhafolcutogrid2, con);
            OleDbDataReader rd3 = komutbirhafogrenciadfoto.ExecuteReader();
            while (rd3.Read())
            {
                lblolcbir.Text = rd3[0].ToString() + " " + rd3[1].ToString();
                pckolcbir.ImageLocation = rd3[2].ToString();

            }
            con.Close();
            //dönemin üçüncü öğrencisi
            con.Open();
            OleDbCommand komuthaf2olcut = new OleDbCommand("SELECT    TOP 3 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            komuthaf2olcut.Parameters.AddWithValue("@H1", hafta);
            OleDbDataReader komuthafikiolcutrd = komuthaf2olcut.ExecuteReader();
            while (komuthafikiolcutrd.Read())
            {
                ikihafolcutogrid2 = int.Parse(komuthafikiolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutikihafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + ikihafolcutogrid2, con);
            OleDbDataReader rd4 = komutikihafogrenciadfoto.ExecuteReader();
            while (rd4.Read())
            {
                lblolciki.Text = rd4[0].ToString() + " " + rd4[1].ToString();
                pckolciki.ImageLocation = rd4[2].ToString();

            }
            con.Close();

            //dönemin dördüncü öğrencisi
            con.Open();
            OleDbCommand komuthaf3olcut = new OleDbCommand("SELECT    TOP 4 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR  GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
           
            OleDbDataReader komuthafucolcutrd = komuthaf3olcut.ExecuteReader();
            while (komuthafucolcutrd.Read())
            {
                uchafolcutogrid2 = int.Parse(komuthafucolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutuchafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + uchafolcutogrid2, con);
            OleDbDataReader rd5 = komutuchafogrenciadfoto.ExecuteReader();
            while (rd5.Read())
            {
                lblolcuc.Text = rd5[0].ToString() + " " + rd5[1].ToString();
                pckolcuc.ImageLocation = rd5[2].ToString();

            }
            con.Close();

            //dönemin beşinci öğrencisi
            con.Open();
            OleDbCommand komuthaf4olcut = new OleDbCommand("SELECT    TOP 5 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR  GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            
            OleDbDataReader komuthafdortolcutrd = komuthaf4olcut.ExecuteReader();
            while (komuthafdortolcutrd.Read())
            {
                dorthafolcutogrid2 = int.Parse(komuthafdortolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutdorthafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + dorthafolcutogrid2, con);
            OleDbDataReader rd6 = komutdorthafogrenciadfoto.ExecuteReader();
            while (rd6.Read())
            {
                lblolcdort.Text = rd6[0].ToString() + " " + rd6[1].ToString();
                pckolcdort.ImageLocation = rd6[2].ToString();

            }
            con.Close();
            //dönemin altıncı öğrencisi
            con.Open();
            OleDbCommand komuthaf5olcut = new OleDbCommand("SELECT    TOP 6 OGRENCININID,   SUM([TOPLAMPUAN]) FROM    TBLNOTLAR  GROUP BY    OGRENCININID ORDER BY    SUM([TOPLAMPUAN]) DESC", con);
            
            OleDbDataReader komuthafbesolcutrd = komuthaf5olcut.ExecuteReader();
            while (komuthafbesolcutrd.Read())
            {
                beshafolcutogrid2 = int.Parse(komuthafbesolcutrd[0].ToString());
            }
            con.Close();
            con.Open();
            OleDbCommand komutbeshafogrenciadfoto = new OleDbCommand("select OGRENCIADISOYADI, SINIFAD, FOTOGRAFYOL FROM TBLOGRENCILER INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLOGRENCILER.OGRENCISINIFI WHERE OGRENCIID=" + beshafolcutogrid2, con);
            OleDbDataReader rd7 = komutbeshafogrenciadfoto.ExecuteReader();
            while (rd7.Read())
            {
                lblolcbes.Text = rd7[0].ToString() + " " + rd7[1].ToString();
                pckolcbes.ImageLocation = rd7[2].ToString();

            }
            con.Close();

        }

       
        private void frmenler_Load(object sender, EventArgs e)
        {

            haftanineniyileri();
            donemineniyileri();
           

        }

        private void button1_Click(object sender, EventArgs e)
        {
            label8.Text=DateTime.Now.DayOfWeek.ToString();
        }
    }
}
