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
using System.Drawing.Imaging;

namespace PERFORMANS
{
    public partial class frmistatistikler : Form
    {
        public frmistatistikler()
        {
            InitializeComponent();
        }
        public int siniftoplamsayi;

        int sinif1;
        int sinif2;
        int sinif3;
        int sinif4;
        int sinif5;
        int sinif6;

        int ogr1;
        int ogr2;
        int ogr3;
        int ogr4;
        int ogr5;
        int ogr6;
        baglantisinif conn = new baglantisinif();
        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        public int ogretmennum;
        public int bransnum;
        private void pictureBox5_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        string birinciolcut;
        string ikinciolcut;
        string ucuncuolcut;
        string dorduncuolcut;
        string besinciolcut;
        string toplampuan;

        public DataTable dt=new DataTable();
        int sira = -1;
        void ogrencilistele()
        {
            lblolcutbır.ForeColor = Color.Black;
            lblolcutiki.ForeColor = Color.Black;
            lblolcutuc.ForeColor = Color.Black;
            lblolcutdort.ForeColor = Color.Black;
            lblolcutbes.ForeColor = Color.Black; 
            chart1.Series["OGRENCI"].Points.Clear();
            
            OleDbConnection con = new OleDbConnection(conn.baglan);
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter("select * from TBLOGRENCILER where OGRENCISINIFI=" + cmbsınıfsec.SelectedValue, con);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;

                DataGridViewCellStyle stil = new DataGridViewCellStyle();
                stil.BackColor = Color.Yellow;
                dataGridView1.Rows[sira].DefaultCellStyle = stil;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[4].Visible = false;
                dataGridView1.Columns[2].Visible = false;


                try
                {
                    pckogrencifoto.ImageLocation = dataGridView1.Rows[sira].Cells[4].Value.ToString();
                    lblogrenciadsoyad.Text = dataGridView1.Rows[sira].Cells[1].Value.ToString();
                    lblogrencinumara.Text = dataGridView1.Rows[sira].Cells[3].Value.ToString();
                    lblogrenciid.Text = dataGridView1.Rows[sira].Cells[0].Value.ToString();

                }
                catch (Exception hata)
                {

                    MessageBox.Show(hata.ToString());
                }
               

                dataGridView1.FirstDisplayedScrollingRowIndex = sira;
                //datagridin scroll görünür sırasını sira değişkeninden alır.

                con.Open();


               


                con.Close();

                con.Open();

                if (lblogrencinumara.Text != "")
                {
                    OleDbCommand komutderssayioku = new OleDbCommand("select count(BRANS) from TBLNOTLAR WHERE BRANS=@A1 and OGRENCINUMARASI=@A2", con);
                    komutderssayioku.Parameters.AddWithValue("@A1", bransnum);
                    komutderssayioku.Parameters.AddWithValue("@A2", int.Parse(lblogrencinumara.Text));
                    OleDbDataReader rd2 = komutderssayioku.ExecuteReader();
                    while (rd2.Read())
                    {

                        lbltoplamderssayisi.Text = rd2[0].ToString();
                    }

                }


                con.Close();

                lblolcutbır.Text = birinciolcut;
                lblolcutiki.Text = ikinciolcut;
                lblolcutuc.Text = ucuncuolcut;
                lblolcutdort.Text = dorduncuolcut;
                lblolcutbes.Text = besinciolcut;
                


            }
            catch (Exception hata )
            {
                MessageBox.Show(hata.ToString());
            }
            con.Open();
            OleDbCommand komutsirala = new OleDbCommand("SELECT OGRENCININID, SUM([1_OLCUT]) AS 'BİRİNCİ ÖLÇÜT', SUM([2_OLCUT]) AS 'İKİNCİ ÖLÇÜT', SUM([3_OLCUT]) AS 'ÜÇÜNCÜ ÖLÇÜT', SUM([4_OLCUT]) AS 'DÖRDÜNCÜ ÖLÇÜT', SUM([5_OLCUT]) AS 'BEŞİNCİ ÖLÇÜT', SUM(TOPLAMPUAN) AS 'TOPLAM PUAN' FROM TBLNOTLAR WHERE OGRENCININID=@P1 GROUP BY OGRENCININID ORDER BY  SUM(TOPLAMPUAN) DESC ", con);
            komutsirala.Parameters.AddWithValue("@P1", int.Parse(lblogrenciid.Text));
            OleDbDataReader rdsira= komutsirala.ExecuteReader();
            while (rdsira.Read())
            {
                lblbirpuan.Text=rdsira[1].ToString();
                lblikipuan.Text=rdsira[2].ToString();
                lblucpuan.Text=rdsira[3].ToString();
                lbldortpuan.Text = rdsira[4].ToString();
                lblbespuan.Text=rdsira[5].ToString();
                lbltoplam.Text=rdsira[6].ToString();
               
                chart1.Series["OGRENCI"].Points.AddXY("1", rdsira[1].ToString());
               

                chart1.Series["OGRENCI"].Points.AddXY("2", rdsira[2].ToString());
                
                chart1.Series["OGRENCI"].Points.AddXY("3", rdsira[3].ToString());
               
                chart1.Series["OGRENCI"].Points.AddXY("4", rdsira[4].ToString());
               
                chart1.Series["OGRENCI"].Points.AddXY("5", rdsira[5].ToString());
                chart1.Series["OGRENCI"].Points.AddXY("6", rdsira[6].ToString());
                ogr1 = int.Parse(rdsira[1].ToString());
                ogr2 = int.Parse(rdsira[2].ToString());
                ogr3 = int.Parse(rdsira[3].ToString());
                ogr4 = int.Parse(rdsira[4].ToString());
                ogr5 = int.Parse(rdsira[5].ToString());
                ogr6 = int.Parse(rdsira[6].ToString());


                if (ogr1 < sinif1)
                {
                    lblolcutbır.ForeColor = Color.Red;
                }
                if (ogr2 < sinif2)
                {
                    lblolcutiki.ForeColor = Color.Red;
                }
                if (ogr3 < sinif3)
                {
                    lblolcutuc.ForeColor = Color.Red;
                }
                if (ogr4 < sinif4)
                {
                    lblolcutdort.ForeColor = Color.Red;
                }
                if (ogr5 < sinif5)
                {
                    lblolcutbes.ForeColor = Color.Red;
                }
                if (ogr6 < sinif6)
                {
                    lbltpl.ForeColor = Color.Red;
                }


            }
            con.Close();
            con.Open();
            if (lblogrencinumara.Text != "")
            {
               
                OleDbCommand komutnotoku = new OleDbCommand("select UYARI, TARIH FROM TBLNOTLAR WHERE OGRENCININID=@K1 AND UYARIVARYOK=@K2", con);
                komutnotoku.Parameters.AddWithValue("@K1", int.Parse(lblogrenciid.Text));
               komutnotoku.Parameters.AddWithValue("@K2", true);
                OleDbDataAdapter dauyarioku = new OleDbDataAdapter();
                dauyarioku.SelectCommand = komutnotoku;
                DataTable dtuyarioku = new DataTable();
                dauyarioku.Fill(dtuyarioku);
                dataGridView2.DataSource = dtuyarioku;
                
            }
            con.Close();
            richTextBox1.Text = "";

            




        }
       
        void degistir()
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand siraoku = new OleDbCommand("select count(*) from TBLOGRENCILER WHERE OGRENCISINIFI=@P1", con);
            siraoku.Parameters.AddWithValue("@P1", cmbsınıfsec.SelectedValue);
            OleDbDataReader rd = siraoku.ExecuteReader();
            while (rd.Read())
            {

                if (sira <= int.Parse(rd[0].ToString()))
                {
                    sira = sira + 1;
                    if (sira == int.Parse(rd[0].ToString()))
                    {
                        
                        sira = 0;
                    }
                }
            }
            con.Close();
            ogrencilistele();
           

        }
        public string rol;
        private void frmistatistikler_Load(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            if (rol == "ogretmen")
            {
                //ÖLÇÜT OKU
                bransnum = 1;
                con.Open();
                OleDbCommand cmdolcutoku = new OleDbCommand("select OLCUTBIR, OLCUTIKI, OLCUTUC, OLCUTDORT, OLCUTBES FROM TBLOGRETMENLER WHERE OGRETMENID=" + ogretmennum, con);
                OleDbDataReader rd = cmdolcutoku.ExecuteReader();
                while (rd.Read())
                {
                    birinciolcut = rd[0].ToString();
                    ikinciolcut = rd[1].ToString();
                    ucuncuolcut = rd[2].ToString();
                    dorduncuolcut = rd[3].ToString();
                    besinciolcut = rd[4].ToString();





                }
                con.Close();


                //combobox sınıf seç
                con.Open();
                OleDbCommand komutsinifoku = new OleDbCommand("select DISTINCT(SINIF), SINIFAD FROM TBLDERSPROGRAMI INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE OGRETMEN=@P1", con);
                komutsinifoku.Parameters.AddWithValue("@P1", ogretmennum); 
                OleDbDataAdapter da4 = new OleDbDataAdapter(komutsinifoku);
                DataTable dt4 = new DataTable();
                da4.Fill(dt4);
                cmbsınıfsec.DisplayMember = "SINIFAD";
                cmbsınıfsec.ValueMember = "SINIF";
                cmbsınıfsec.DataSource = dt4;
                con.Close();
                //ÖLÇÜT DATAGRID VİEW

            }
            else if(rol == "yonetici")
            {
                //ÖLÇÜT OKU
               


                //combobox sınıf seç

                OleDbDataAdapter da4 = new OleDbDataAdapter("select DISTINCT(SINIF), SINIFAD FROM TBLDERSPROGRAMI INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF", con);
                DataTable dt4 = new DataTable();
                da4.Fill(dt4);
                cmbsınıfsec.DisplayMember = "SINIFAD";
                cmbsınıfsec.ValueMember = "SINIFID";
                cmbsınıfsec.DataSource = dt4;

                //ÖLÇÜT DATAGRID VİEW

            }




        }
        
        private void button5_Click(object sender, EventArgs e)
        {
            frmdatagrid2 frmdatagrid2 = new frmdatagrid2();
            frmdatagrid2.birinciolcut = birinciolcut;
            frmdatagrid2.ikinciolcut  = ikinciolcut;
            frmdatagrid2.ucuncuolcut=ucuncuolcut;
            frmdatagrid2.dorduncuolcut=dorduncuolcut;
            frmdatagrid2.besinciolcut=besinciolcut;
            frmdatagrid2.sec = 0;
            frmdatagrid2.sinif = int.Parse(cmbsınıfsec.SelectedValue.ToString());
            frmdatagrid2.Show();
        }

        private void btnogrencigetir_Click(object sender, EventArgs e)
        {
            degistir();
            sinifpuangetir();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            lblolcutbır.ForeColor = Color.Black;
            lblolcutiki.ForeColor = Color.Black;
            lblolcutuc.ForeColor = Color.Black;
            lblolcutdort.ForeColor = Color.Black;
            lblolcutbes.ForeColor = Color.Black;
            int secilen = dataGridView1.SelectedCells[0].RowIndex;
            
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.Columns[2].Visible = false;
            chart1.Series["OGRENCI"].Points.Clear();    
            OleDbConnection con = new OleDbConnection(conn.baglan);
            
            pckogrencifoto.ImageLocation = dataGridView1.Rows[secilen].Cells[4].Value.ToString();
            lblogrenciadsoyad.Text = dataGridView1.Rows[secilen].Cells[1].Value.ToString();
            lblogrencinumara.Text = dataGridView1.Rows[secilen].Cells[3].Value.ToString();
            lblogrenciid.Text = dataGridView1.Rows[secilen].Cells[0].Value.ToString();

           



            //datagridin scroll görünür sırasını sira değişkeninden alır.

            con.Open();


           


            con.Close();

            con.Open();

            if (lblogrencinumara.Text != "")
            {
                OleDbCommand komutderssayioku = new OleDbCommand("select count(BRANS) from TBLNOTLAR WHERE BRANS=@A1 and OGRENCINUMARASI=@A2", con);
                komutderssayioku.Parameters.AddWithValue("@A1", bransnum);
                komutderssayioku.Parameters.AddWithValue("@A2", int.Parse(lblogrencinumara.Text));
                OleDbDataReader rd2 = komutderssayioku.ExecuteReader();
                while (rd2.Read())
                {

                    lbltoplamderssayisi.Text = rd2[0].ToString();
                }

            }


            con.Close();

            lblolcutbır.Text = birinciolcut;
            lblolcutiki.Text = ikinciolcut;
            lblolcutuc.Text = ucuncuolcut;
            lblolcutdort.Text = dorduncuolcut;
            lblolcutbes.Text = besinciolcut;

            
            con.Open();
            OleDbCommand komutsirala = new OleDbCommand("SELECT OGRENCININID, SUM([1_OLCUT]) AS 'BİRİNCİ ÖLÇÜT', SUM([2_OLCUT]) AS 'İKİNCİ ÖLÇÜT', SUM([3_OLCUT]) AS 'ÜÇÜNCÜ ÖLÇÜT', SUM([4_OLCUT]) AS 'DÖRDÜNCÜ ÖLÇÜT', SUM([5_OLCUT]) AS 'BEŞİNCİ ÖLÇÜT', SUM(TOPLAMPUAN) AS 'TOPLAM PUAN' FROM TBLNOTLAR WHERE OGRENCININID=@P1 GROUP BY OGRENCININID ORDER BY  SUM(TOPLAMPUAN) DESC ", con);
            komutsirala.Parameters.AddWithValue("@P1", int.Parse(lblogrenciid.Text));
            OleDbDataReader rdsira = komutsirala.ExecuteReader();
            while (rdsira.Read())
            {
                lblbirpuan.Text = rdsira[1].ToString();
                lblikipuan.Text = rdsira[2].ToString();
                lblucpuan.Text = rdsira[3].ToString();
                lbldortpuan.Text = rdsira[4].ToString();
                lblbespuan.Text = rdsira[5].ToString();
                lbltoplam.Text = rdsira[6].ToString();

                chart1.Series["OGRENCI"].Points.AddXY(" 1", rdsira[1].ToString());
                chart1.Series["OGRENCI"].Points.AddXY("2", rdsira[2].ToString());
                chart1.Series["OGRENCI"].Points.AddXY("3", rdsira[3].ToString());
                chart1.Series["OGRENCI"].Points.AddXY("4", rdsira[4].ToString());
                chart1.Series["OGRENCI"].Points.AddXY("5", rdsira[5].ToString()); 
                chart1.Series["OGRENCI"].Points.AddXY(lbltpl.Text, rdsira[6].ToString());
                ogr1 = int.Parse(rdsira[1].ToString());
                ogr2 = int.Parse(rdsira[2].ToString());
                ogr3 = int.Parse(rdsira[3].ToString());
                ogr4 = int.Parse(rdsira[4].ToString());
                ogr5 = int.Parse(rdsira[5].ToString());
                ogr6 = int.Parse(rdsira[6].ToString());




            }
            con.Close();



            if (ogr1 < sinif1)
            {
                lblolcutbır.ForeColor = Color.Red;
            }
            if (ogr2 < sinif2)
            {
                lblolcutiki.ForeColor = Color.Red;
            }
            if (ogr3 < sinif3)
            {
                lblolcutuc.ForeColor = Color.Red;
            }
            if (ogr4 < sinif4)
            {
                lblolcutdort.ForeColor = Color.Red;
            }
            if (ogr5 < sinif5)
            {
                lblolcutbes.ForeColor = Color.Red;
            }
            if (ogr6 < sinif6)
            {
                lbltpl.ForeColor = Color.Red;
            }


        }
        void sinifpuangetir()
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            chart1.Series["SINIF"].Points.Clear();
            con.Open();
            OleDbCommand sinifogrencisayi = new OleDbCommand("select count(*) from TBLOGRENCILER WHERE OGRENCISINIFI=@P1", con);
            sinifogrencisayi.Parameters.AddWithValue("@P1", cmbsınıfsec.SelectedValue);
            OleDbDataReader rdsayi = sinifogrencisayi.ExecuteReader();
            while (rdsayi.Read())
            {
               
                    siniftoplamsayi = int.Parse(rdsayi[0].ToString());
               
            }
            con.Close();
            con.Open();
            OleDbCommand komutsinif = new OleDbCommand("SELECT SINIF, SUM([1_OLCUT]) AS 'BİRİNCİ ÖLÇÜT', SUM([2_OLCUT]) AS 'İKİNCİ ÖLÇÜT', SUM([3_OLCUT]) AS 'ÜÇÜNCÜ ÖLÇÜT', SUM([4_OLCUT]) AS 'DÖRDÜNCÜ ÖLÇÜT', SUM([5_OLCUT]) AS 'BEŞİNCİ ÖLÇÜT', SUM(TOPLAMPUAN) AS 'TOPLAM PUAN' FROM TBLNOTLAR WHERE SINIF=@P1 GROUP BY SINIF ORDER BY  SUM(TOPLAMPUAN) DESC ", con);
            komutsinif.Parameters.AddWithValue("@P1", cmbsınıfsec.SelectedValue);
            OleDbDataReader rdsinif2 = komutsinif.ExecuteReader();
            while (rdsinif2.Read())
            {


                chart1.Series["SINIF"].Points.AddXY("1", (int.Parse(rdsinif2[1].ToString()) / siniftoplamsayi).ToString());
                chart1.Series["SINIF"].Points.AddXY("2", (int.Parse(rdsinif2[2].ToString()) / siniftoplamsayi).ToString());
                chart1.Series["SINIF"].Points.AddXY("3", (int.Parse(rdsinif2[3].ToString()) / siniftoplamsayi).ToString());
                chart1.Series["SINIF"].Points.AddXY("4", (int.Parse(rdsinif2[4].ToString()) / siniftoplamsayi).ToString());
                chart1.Series["SINIF"].Points.AddXY("5", (int.Parse(rdsinif2[5].ToString()) / siniftoplamsayi).ToString());
                chart1.Series["SINIF"].Points.AddXY(lbltpl.Text, (int.Parse(rdsinif2[6].ToString()) / siniftoplamsayi).ToString());

                sinif1 = int.Parse(rdsinif2[1].ToString()) / siniftoplamsayi;
                sinif2 = int.Parse(rdsinif2[2].ToString()) / siniftoplamsayi;
                sinif3 = int.Parse(rdsinif2[3].ToString()) / siniftoplamsayi;
                sinif4 = int.Parse(rdsinif2[4].ToString()) / siniftoplamsayi;
                sinif5 = int.Parse(rdsinif2[5].ToString()) / siniftoplamsayi;
                sinif6 = int.Parse(rdsinif2[5].ToString()) / siniftoplamsayi;

            }
            con.Close();
        }

        private void cmbsınıfsec_SelectedIndexChanged(object sender, EventArgs e)
        {
            OleDbConnection con=new OleDbConnection(conn.baglan);
            if (rol == "ogretmen")
            {
                con.Open();
                OleDbCommand komutderssec = new OleDbCommand("select  DISTINCT(DERS), BRANSADI FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID=TBLDERSPROGRAMI.DERS WHERE SINIF=@D1 AND OGRETMEN=@D2", con);
                komutderssec.Parameters.AddWithValue("@D1", cmbsınıfsec.SelectedValue);
                komutderssec.Parameters.AddWithValue("@D2", ogretmennum);
                OleDbDataAdapter da2 = new OleDbDataAdapter(komutderssec);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                cmbderssec.DisplayMember = "BRANSADI";
                cmbderssec.ValueMember = "DERS";
                cmbderssec.DataSource = dt2;
                con.Close();
            }
            else if(rol == "yonetici")
            {
                con.Open();
                OleDbCommand komutderssec = new OleDbCommand("select  DISTINCT(DERS), BRANSADI FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID=TBLDERSPROGRAMI.DERS WHERE SINIF=@D1", con);
                komutderssec.Parameters.AddWithValue("@D1", cmbsınıfsec.SelectedValue);
                
                OleDbDataAdapter da2 = new OleDbDataAdapter(komutderssec);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                cmbderssec.DisplayMember = "BRANSADI";
                cmbderssec.ValueMember = "DERS";
                cmbderssec.DataSource = dt2;
                con.Close();
            }


        }

        private void button8_Click(object sender, EventArgs e)
        {
            frmdatagrid2 frmdatagrid2 = new frmdatagrid2();
            frmdatagrid2.birinciolcut = birinciolcut;
            frmdatagrid2.ikinciolcut = ikinciolcut;
            frmdatagrid2.ucuncuolcut = ucuncuolcut;
            frmdatagrid2.dorduncuolcut = dorduncuolcut;
            frmdatagrid2.besinciolcut = besinciolcut;
            frmdatagrid2.sec = 1;
            frmdatagrid2.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            frmenler frmenler = new frmenler();
            frmenler.Show();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            richTextBox1.Text = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString()+" tarih:" + dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            frmdatagrid2 frmdatagrid2 = new frmdatagrid2();
            frmdatagrid2.birinciolcut = birinciolcut;
            frmdatagrid2.ikinciolcut = ikinciolcut;
            frmdatagrid2.ucuncuolcut = ucuncuolcut;
            frmdatagrid2.dorduncuolcut = dorduncuolcut;
            frmdatagrid2.besinciolcut = besinciolcut;
            frmdatagrid2.sinif = int.Parse(cmbsınıfsec.SelectedValue.ToString());
            frmdatagrid2.sec = 5;
            frmdatagrid2.Show();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            frmdatagrid2 frmdatagrid2 = new frmdatagrid2();
            frmdatagrid2.birinciolcut = birinciolcut;
            frmdatagrid2.ikinciolcut = ikinciolcut;
            frmdatagrid2.ucuncuolcut = ucuncuolcut;
            frmdatagrid2.dorduncuolcut = dorduncuolcut;
            frmdatagrid2.besinciolcut = besinciolcut;
            frmdatagrid2.ders = int.Parse(cmbderssec.SelectedValue.ToString());
            frmdatagrid2.sinif = int.Parse(cmbsınıfsec.SelectedValue.ToString());
            frmdatagrid2.sec = 6;

            
            frmdatagrid2.Show();

        }

        private void btnsinifdersdonemnottable_Click(object sender, EventArgs e)
        {
            frmpuankayitlari frmpuankayitlari=new frmpuankayitlari();
            frmpuankayitlari.datasec = "puankayitlariders";
            frmpuankayitlari.brans = int.Parse(cmbderssec.SelectedValue.ToString());
            frmpuankayitlari.sinif = int.Parse(cmbsınıfsec.SelectedValue.ToString());
            frmpuankayitlari.Show();
        }
    }
}
