using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Xml.Linq;

namespace PERFORMANS
{
    public partial class frmogretmendurum : Form
    {
        public frmogretmendurum()
        {
            InitializeComponent();
        }
        baglantisinif conn=new baglantisinif();
        private void frmogretmendurum_Load(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand komutogretmendurum = new OleDbCommand("select ADISOYADI, count(*) AS 'PUANLANAN DERS SAYISI' FROM TBLDERSPROGRAMI INNER JOIN TBLOGRETMENLER ON TBLDERSPROGRAMI.OGRETMEN=TBLOGRETMENLER.OGRETMENID WHERE OLCDURUM<>false  GROUP BY ADISOYADI ORDER BY COUNT(*) DESC", con);
            OleDbDataAdapter da=new OleDbDataAdapter(komutogretmendurum);
            DataTable dt=new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();

            con.Open();
            OleDbCommand komutogretmendurum2 = new OleDbCommand("SELECT  DISTINCT ADISOYADI AS'PUANLAMA YAPMAYAN ÖĞRETMEN LİSTESİ', '0' AS 'PUANLANAN DERS SAYISI' FROM TBLDERSPROGRAMI INNER JOIN TBLOGRETMENLER ON TBLOGRETMENLER.OGRETMENID=TBLDERSPROGRAMI.OGRETMEN WHERE OGRETMEN NOT IN(SELECT OGRETMEN FROM TBLDERSPROGRAMI WHERE OLCDURUM<>false )", con);
            OleDbDataAdapter da2 = new OleDbDataAdapter(komutogretmendurum2);
            DataTable dt2 = new DataTable();
            da2.Fill(dt2);
            dataGridView2.DataSource = dt2;
            con.Close();

            con.Open();
            OleDbCommand komutogretmendurum3 = new OleDbCommand("SELECT DISTINCT ADISOYADI AS 'PUANLAMAYI TAMAMLAYAN ÖĞRETMEN LİSTESİ','BÜTÜN DERSLER PUANLANDI' AS 'PUANLAMA DURUMU' FROM TBLDERSPROGRAMI INNER JOIN TBLOGRETMENLER ON TBLDERSPROGRAMI.OGRETMEN=TBLOGRETMENLER.OGRETMENID WHERE OGRETMEN NOT IN(SELECT OGRETMEN FROM TBLDERSPROGRAMI WHERE OLCDURUM=0)", con);
            OleDbDataAdapter da3 = new OleDbDataAdapter(komutogretmendurum3);
            DataTable dt3 = new DataTable();
            da3.Fill(dt3);
            dataGridView3.DataSource = dt3;
            con.Close();
        }
        int ogretmenid;
        string ogretmenintc;
        public string sifrele(string s)
        {
            byte[] sdizi = ASCIIEncoding.ASCII.GetBytes(s);
            string sifreli = Convert.ToBase64String(sdizi);
            return sifreli;
        }

        public string sifrecoz(string s)
        {
            byte[] sdizi = Convert.FromBase64String(s);
            string sifresiz = ASCIIEncoding.ASCII.GetString(sdizi);
            return sifresiz;
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            try
            {
                con.Open();
                OleDbCommand komutogretmentcoku = new OleDbCommand("select OGRETMENID from TBLOGRETMENLER where ADISOYADI='" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "'", con);
                OleDbDataReader dr = komutogretmentcoku.ExecuteReader();
                if (dr.Read())
                {
                    ogretmenid = int.Parse(dr["OGRETMENID"].ToString());
                }
                con.Close();
                con.Open();
                OleDbCommand ogretmentcbul = new OleDbCommand("select TCKIMLIKNO FROM TBLOGRETMENLER where OGRETMENID=" + ogretmenid + "", con);
                OleDbDataReader dr2 = ogretmentcbul.ExecuteReader();
                if (dr2.Read())
                {
                    ogretmenintc = dr2[0].ToString();

                }
                frmdersprogrami drp = new frmdersprogrami();
                drp.ogretmentc = sifrecoz(ogretmenintc);
                drp.rolum = "yonetici";
                drp.Show();

            }
            catch (Exception hata)
            {

                MessageBox.Show("Hata oluştu. Lütfen program sağlayıcınıza başvurunuz." + hata.Message, "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            


        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            try
            {
                con.Open();
                OleDbCommand komutogretmentcoku = new OleDbCommand("select OGRETMENID from TBLOGRETMENLER where ADISOYADI='" + dataGridView2.CurrentRow.Cells[0].Value.ToString() + "'", con);
                OleDbDataReader dr = komutogretmentcoku.ExecuteReader();
                if (dr.Read())
                {
                    ogretmenid = int.Parse(dr["OGRETMENID"].ToString());
                }
                con.Close();
                con.Open();
                OleDbCommand ogretmentcbul = new OleDbCommand("select TCKIMLIKNO FROM TBLOGRETMENLER where OGRETMENID=" + ogretmenid + "", con);
                OleDbDataReader dr2 = ogretmentcbul.ExecuteReader();
                if (dr2.Read())
                {
                    ogretmenintc = dr2[0].ToString();

                }
                frmdersprogrami drp = new frmdersprogrami();
                drp.ogretmentc = sifrecoz(ogretmenintc);
                drp.rolum = "yonetici";
                drp.Show();

            }
            catch (Exception hata)
            {


                MessageBox.Show("Hata oluştu. Lütfen program sağlayıcınıza başvurunuz." + hata.Message, "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            try
            {
                con.Open();
                OleDbCommand komutogretmentcoku = new OleDbCommand("select OGRETMENID from TBLOGRETMENLER where ADISOYADI='" + dataGridView3.CurrentRow.Cells[0].Value.ToString() + "'", con);
                OleDbDataReader dr = komutogretmentcoku.ExecuteReader();
                if (dr.Read())
                {
                    ogretmenid = int.Parse(dr["OGRETMENID"].ToString());
                }
                con.Close();
                con.Open();
                OleDbCommand ogretmentcbul = new OleDbCommand("select TCKIMLIKNO FROM TBLOGRETMENLER where OGRETMENID=" + ogretmenid + "", con);
                OleDbDataReader dr2 = ogretmentcbul.ExecuteReader();
                if (dr2.Read())
                {
                    ogretmenintc = dr2[0].ToString();

                }
                frmdersprogrami drp = new frmdersprogrami();
                drp.ogretmentc = sifrecoz(ogretmenintc);
                drp.rolum = "yonetici";
                drp.Show();

            }
            catch (Exception hata)
            {


                MessageBox.Show("Hata oluştu. Lütfen program sağlayıcınıza başvurunuz." + hata.Message, "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}
