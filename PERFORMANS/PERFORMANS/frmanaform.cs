using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.OleDb;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Media;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PERFORMANS
{
    public partial class frmanaform : Form
    {
        public frmanaform()
        {
            InitializeComponent();
        }
        baglantisinif conn = new baglantisinif();
        public string rol;
        int sira = -1;
        public int gunlukhaftalik;
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
        public int haftaal(DateTime dtPassed)
        {

            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dtPassed, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;

        }
        void ogrencilistele()
        {
            int hafta=haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter("select * from TBLOGRENCILER where OGRENCISINIFI=" + sinif, con);
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
                    
                        pictureBox2.ImageLocation = dataGridView1.Rows[sira].Cells[4].Value.ToString();
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


                if (lblogrencinumara.Text != "")
                {
                    OleDbCommand komuttoplampuanoku = new OleDbCommand("select SUM(TOPLAMPUAN) from TBLNOTLAR WHERE HAFTA=@A1 and OGRENCINUMARASI=@A2", con);
                    komuttoplampuanoku.Parameters.AddWithValue("@A1", hafta );
                    komuttoplampuanoku.Parameters.AddWithValue("@A2", int.Parse(lblogrencinumara.Text));
                    OleDbDataReader rd2 = komuttoplampuanoku.ExecuteReader();
                    while (rd2.Read())
                    {

                        lblogrencitoplampuan.Text = rd2[0].ToString();
                    }

                }


                con.Close();

                con.Open();
                
                if (lblogrencinumara.Text != "")
                {
                    OleDbCommand komutderssayioku = new OleDbCommand("select count(BRANS) from TBLNOTLAR WHERE HAFTA=@A1 and OGRENCINUMARASI=@A2", con);
                    komutderssayioku.Parameters.AddWithValue("@A1", hafta);
                    komutderssayioku.Parameters.AddWithValue("@A2", int.Parse(lblogrencinumara.Text));
                    OleDbDataReader rd2 = komutderssayioku.ExecuteReader();
                    while (rd2.Read())
                    {

                        lbltoplamderssayisi.Text = rd2[0].ToString();
                    }

                }


                con.Close();




                rchnotdus.Text = "";

                ogrencipuan();
                if (chkhesapla.Checked == true)
                {
                    puansırala();

                }



            }
            catch (Exception)
            {

                MessageBox.Show("Değerlendirilicek sınıf yoktur", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        void puansırala()
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            if (chkhesapla.Checked == true)
            {
                con.Open();
                OleDbCommand komutpuansirala = new OleDbCommand("select OGRENCIADISOYADI, SUM(TOPLAMPUAN) AS 'TOPLAMPUAN'  FROM TBLNOTLAR   INNER JOIN TBLOGRENCILER ON TBLOGRENCILER.OGRENCIID=TBLNOTLAR.OGRENCININID WHERE HAFTA=@P1 AND SINIF=@P2 GROUP BY OGRENCIADISOYADI ORDER BY  SUM(TOPLAMPUAN) DESC ", con);
                komutpuansirala.Parameters.AddWithValue("@P1", hafta);
                komutpuansirala.Parameters.AddWithValue("@P2", sinif);

                OleDbDataAdapter da2 = new OleDbDataAdapter();
                da2.SelectCommand = komutpuansirala;
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                dataGridView2.DataSource = dt2;
                con.Close();
            }
            
            
        }
        public string tc;
        public string gun;
        public int bransid;
        public int ogretmenid;

        public int derssaati;
        public int sinif;

       
        private void frmanaform_Load(object sender, EventArgs e)
        {

            if (gunlukhaftalik == 1)
            {
                rdygunluk.Checked = true;
            }
            else if (gunlukhaftalik == 0)
            {
                rdyhaftalik.Checked = true;
            }
                int hafta = haftaal(DateTime.Now);
            label18.Text = rol.ToString();
            OleDbConnection con = new OleDbConnection(conn.baglan);
            ogrencigetir();

            if (label18.Text == "ogretmen")
            {
              
                con.Open();
                OleDbCommand adbransoku = new OleDbCommand("select OGRETMENID,ADISOYADI, BRANS, BRANSI FROM TBLOGRETMENLER where TCKIMLIKNO=@O1", con);
                adbransoku.Parameters.AddWithValue("@O1", sifrele(tc));
                OleDbDataReader rd = adbransoku.ExecuteReader();
                while (rd.Read())
                {

                    lbladbrans.Text = ("SAYIN " + rd[1] + " HOŞGELDİNİZ." + " BRANŞINIZ: " + rd[3]).ToUpper();
                    lblders.Text = rd[3].ToString();
                    
                    ogretmenid = int.Parse(rd[0].ToString());
                }
                con.Close();
                
                
                con.Open();
                OleDbCommand olcutoku = new OleDbCommand("select OLCUTBIR, OLCUTIKI, OLCUTUC,OLCUTDORT,OLCUTBES FROM TBLOGRETMENLER WHERE OGRETMENID=" + ogretmenid, con);
                OleDbDataReader rd2 = olcutoku.ExecuteReader();
                while (rd2.Read())
                {
                    lblbir.Text = rd2[0].ToString();
                    lbliki.Text = rd2[1].ToString();
                    lbluc.Text = rd2[2].ToString();
                    lbldort.Text = rd2[3].ToString();
                    lblbes.Text = rd2[4].ToString();
                }
                con.Close();
            }

            else if (label18.Text == "yonetici")
            {
                lbladbrans.Text = ("SAYIN YÖNETİCİ HOŞGELDİNİZ.").ToUpper();
                
            }

            
        }
        void ogrencipuan()
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            if (chkhesapla.Checked == true)
            {
                con.Open();
                OleDbCommand komuttoplampuan = new OleDbCommand("select sum(TOPLAMPUAN) FROM TBLNOTLAR WHERE  SINIF =@S1 AND  OGRENCINUMARASI=@P1 AND HAFTA=@P2 ", con);
                komuttoplampuan.Parameters.AddWithValue("@S1", sinif);
                komuttoplampuan.Parameters.AddWithValue("@P1", int.Parse(lblogrencinumara.Text));
                komuttoplampuan.Parameters.AddWithValue("@P2", hafta);

                OleDbDataReader rd = komuttoplampuan.ExecuteReader();
                while (rd.Read())
                {
                    lblogrencitoplampuan.Text = rd[0].ToString();
                }
                con.Close();
            }
           

        }
        Boolean olcdurum;
        void degistir()
        {
            int hafta= haftaal(DateTime.Now); 
            OleDbConnection con = new OleDbConnection(conn.baglan);
            if (chkhesapla.Checked == true)
            {
                puansırala();
                chkpuangor.Visible = true;

            }
            con.Open();
            OleDbCommand siraoku = new OleDbCommand("select count(*) from TBLOGRENCILER WHERE OGRENCISINIFI=@P1", con);
            siraoku.Parameters.AddWithValue("@P1", sinif);
            OleDbDataReader rd = siraoku.ExecuteReader();
            while (rd.Read())
            {
                
                if (sira>=-1 && sira < int.Parse(rd[0].ToString()))
                {
                    sira = sira + 1;
                    
                    rdygunluk.Enabled = false;
                    rdyhaftalik.Enabled = false;
                    button5.Enabled = false;
                    if (sira == int.Parse(rd[0].ToString()))
                    {
                       chkhesapla.Checked=true;
                        puansırala();
                        olcdurum = true;
                        MessageBox.Show("Bütün sınıfı puanladığınız için teşekkür ederiz. Puanlama yapmak için Ders Programına Git butonuna tıklayabilirsiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                      
                        rdygunluk.Enabled = true;
                        rdyhaftalik.Enabled = true;
                        button5.Enabled = true;
                        
                        sira = -1;
                        dataGridView1.DataSource = null;
                        pictureBox2.ImageLocation = null;
                        lblogrenciadsoyad.Text = null;
                        lblogrencinumara.Text = null;
                        lblogrenciid.Text = null;


                    }
                    else { olcdurum = false;
                    }
                }
            }
            con.Close();
            if (rdygunluk.Checked == true)
            {
                con.Open();
                OleDbCommand olcdurumguncelle = new OleDbCommand("update TBLDERSPROGRAMI SET OLCDURUM=@K1 WHERE TARIH=@P1 AND DERSSAATI=@P2 AND SINIF=@P3", con);
                olcdurumguncelle.Parameters.AddWithValue("@K1", true);
                olcdurumguncelle.Parameters.AddWithValue("@P1", gun);
                olcdurumguncelle.Parameters.AddWithValue("@P2", derssaati);
                olcdurumguncelle.Parameters.AddWithValue("@P3", sinif);

                if (olcdurum == true)
                {
                    olcdurumguncelle.ExecuteNonQuery();
                    puansırala();
                    
                }

                con.Close();
            }

            else if(rdyhaftalik.Checked == true)
            {
                con.Open();
                OleDbCommand olcdurumguncelle2= new OleDbCommand("update TBLDERSPROGRAMI SET OLCDURUM=@K1 WHERE SINIF=@P3  AND   DERS =@P4  ", con);
                olcdurumguncelle2.Parameters.AddWithValue("@K1", true);
                olcdurumguncelle2.Parameters.AddWithValue("@P3", sinif);
                olcdurumguncelle2.Parameters.AddWithValue("@P4", bransid);
                

                if (olcdurum == true)
                {
                    olcdurumguncelle2.ExecuteNonQuery();
                    puansırala();
                    
                }

                con.Close();
            }





                ogrencilistele();


            chkdolubir.Visible = true;
            chkdoluiki.Visible = true;
            chkdoluuc.Visible = true;
            chkdoludort.Visible = true;
            chkdolubes.Visible = true;
            chkolumluuyari.Visible = true;
            

            chkbosbir.Visible = false;
            chkbosiki.Visible = false;
            chkbosuc.Visible = false;
            chkbosdort.Visible = false;
            chkbosbes.Visible = false;
            chkolumsuzuyari.Visible = false;

            lblbir.ForeColor = Color.White;
            lbliki.ForeColor = Color.White;
            lbluc.ForeColor = Color.White;
            lbldort.ForeColor = Color.White;
            lblbes.ForeColor = Color.White;
            toplampuani = 5;

            bir = 1;
            iki = 1;
            uc = 1;
            dort = 1;
            bes = 1;
            
           
            
        }

        int ogretmenderssayi;
        int ogretmendersadi;

        void ogrencigetir()
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);

            if (derssaati != null)
            {

                if (rdygunluk.Checked != null && rdyhaftalik.Checked != null)
                {
                    chkhesapla.Checked = false;

                    con.Open();
                    OleDbCommand komutderssayioku = new OleDbCommand("SELECT    count(*)  FROM     TBLDERSPROGRAMI WHERE   DERS = @P3    AND OGRETMEN = @P1    AND SINIF = @P2 AND OLCDURUM=@P4", con);

                    komutderssayioku.Parameters.AddWithValue("@P3", bransid);
                    komutderssayioku.Parameters.AddWithValue("@P1", ogretmenid);
                    komutderssayioku.Parameters.AddWithValue("@P2", sinif);
                    komutderssayioku.Parameters.AddWithValue("@P4", false);
                    OleDbDataReader rd2 = komutderssayioku.ExecuteReader();
                    while (rd2.Read())
                    {
                        ogretmenderssayi = int.Parse(rd2[0].ToString());
                    }
                    con.Close();

                    if (sira < 0)
                    {
                       
                            
                            degistir();

                       
                    }
                    
                    //con.Open();
                    //OleDbCommand komutdersoku = new OleDbCommand("select DERS from TBLDERSPROGRAMI WHERE SINIF=@P1 AND DERSSAATI=@P2 AND TARIH=@P3", con);
                    //komutdersoku.Parameters.AddWithValue("@P1", sinif);
                    //komutdersoku.Parameters.AddWithValue("@P2", derssaati);
                    //komutdersoku.Parameters.AddWithValue("@P3", gun);
                    //OleDbDataReader komutdersokurd = komutdersoku.ExecuteReader();
                    //while (komutdersokurd.Read())
                    //{
                    //    ogretmendersadi = int.Parse(komutdersokurd[0].ToString());
                    //}
                    //con.Close();
                   

                }
                

            }

            else
            {
                MessageBox.Show("Size tanımlanmış ders yoktur");
            }



        }
        private void button5_Click(object sender, EventArgs e)
        {
            frmdersprogrami frmdersprogrami = new frmdersprogrami();
            frmdersprogrami.rol = "ogretmen";
            frmdersprogrami.ogretmentc= tc;
            frmdersprogrami.ogretmenid = ogretmenid;
            this.Close();
            frmdersprogrami.Show();





        }
        
        private void button4_Click(object sender, EventArgs e)
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
           
            con.Open();



            if (lblogrenciadsoyad.Text != "" )
            {
                if (rdygunluk.Checked == true)
                {
                    OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                    komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                    komutpuankaydet.Parameters.AddWithValue("@D1", derssaati);
                    komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                    komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                    komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                    komutpuankaydet.Parameters.AddWithValue("@P4", 0);
                    komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                    komutpuankaydet.ExecuteNonQuery();

                }
                else if (rdyhaftalik.Checked == true)
                {

                    for(int i=1; i<=ogretmenderssayi; i++)
                    {
                        OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                        komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                        komutpuankaydet.Parameters.AddWithValue("@D1", i);
                        komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                        komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                        komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                        komutpuankaydet.Parameters.AddWithValue("@P4", 0);
                        komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                        komutpuankaydet.ExecuteNonQuery();
                    }
                }
            }
            else
            {
                MessageBox.Show("Lütfen ders programınızdan ders seçiniz.", "Bilgi");
            }



                con.Close();
                if (lblogrenciadsoyad.Text != "")
                { degistir(); }
            
                
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();

            if (lblogrenciadsoyad.Text != "")
            {
                if (rdygunluk.Checked == true)
                {
                    OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN, HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                    komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                    komutpuankaydet.Parameters.AddWithValue("@D1", derssaati);
                    komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                    komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                    komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                    komutpuankaydet.Parameters.AddWithValue("@P4", 2);
                    komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                    komutpuankaydet.ExecuteNonQuery();
                }
                else  if (rdyhaftalik.Checked == true)
                {
                    for (int i = 1; i <= ogretmenderssayi; i++)
                    {
                        OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                        komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                        komutpuankaydet.Parameters.AddWithValue("@D1", i);
                        komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                        komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                        komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                        komutpuankaydet.Parameters.AddWithValue("@P4", 2);
                        komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                        komutpuankaydet.ExecuteNonQuery();
                    }
                }
              

            }
            else { MessageBox.Show("Lütfen ders programınızdan ders seçiniz.", "Bilgi"); }
            con.Close();
            if (lblogrenciadsoyad.Text != "")
            { degistir(); }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();

            if (lblogrenciadsoyad.Text != "")
            {
                if (rdygunluk.Checked == true)
                {
                    OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                    komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                    komutpuankaydet.Parameters.AddWithValue("@D1", derssaati);
                    komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                    komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                    komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                    komutpuankaydet.Parameters.AddWithValue("@P4", 1);
                    komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                    komutpuankaydet.ExecuteNonQuery();
                }

                else if (rdyhaftalik.Checked == true)
                {
                    for (int i = 1; i <= ogretmenderssayi; i++)
                    {
                        OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                        komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                        komutpuankaydet.Parameters.AddWithValue("@D1", i);
                        komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                        komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                        komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                        komutpuankaydet.Parameters.AddWithValue("@P4", 1);
                        komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                        komutpuankaydet.ExecuteNonQuery();
                    }
                }

            }

           
            else { MessageBox.Show("Lütfen ders programınızdan ders seçiniz.", "Bilgi"); }
            con.Close();
            if (lblogrenciadsoyad.Text != "")
            { degistir(); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();

            if (lblogrenciadsoyad.Text != "" )
            {
                if (rdygunluk.Checked == true)
                {
                    OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                    komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                    komutpuankaydet.Parameters.AddWithValue("@D1", derssaati);
                    komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                    komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                    komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                    komutpuankaydet.Parameters.AddWithValue("@P4", 3);
                    komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                    komutpuankaydet.ExecuteNonQuery();
                }

                else if(rdyhaftalik.Checked == true)
                {
                    for (int i = 1; i <= ogretmenderssayi; i++)
                    {
                        OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                        komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                        komutpuankaydet.Parameters.AddWithValue("@D1", i);
                        komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                        komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                        komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                        komutpuankaydet.Parameters.AddWithValue("@P4", 3);
                        komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                        komutpuankaydet.ExecuteNonQuery();
                    }
                }


            }
            else { MessageBox.Show("Lütfen ders programınızdan ders seçiniz.", "Bilgi"); }
            con.Close();
            if (lblogrenciadsoyad.Text != "")
            { degistir(); }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();

            if (lblogrenciadsoyad.Text != "" )
            {
                if (rdygunluk.Checked == true)
                {
                    OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                    komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                    komutpuankaydet.Parameters.AddWithValue("@D1", derssaati);
                    komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                    komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                    komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                    komutpuankaydet.Parameters.AddWithValue("@P4", 4);
                    komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                    komutpuankaydet.ExecuteNonQuery();
                }
                else if (rdyhaftalik.Checked == true)
                {
                    for (int i = 1; i <= ogretmenderssayi; i++)
                    {
                        OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5)", con);
                        komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                        komutpuankaydet.Parameters.AddWithValue("@D1", i);
                        komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                        komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P2", lblogrencinumara.Text);
                        komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                        komutpuankaydet.Parameters.AddWithValue("@P4", 4);
                        komutpuankaydet.Parameters.AddWithValue("@P5", hafta);
                        komutpuankaydet.ExecuteNonQuery();
                    }
                }


            }
            else { MessageBox.Show("Lütfen ders programınızdan ders seçiniz.", "Bilgi"); }
            con.Close();
            if (lblogrenciadsoyad.Text != "")
            { degistir(); }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();


            if (lblogrenciadsoyad.Text != "")
            {

                if (rdygunluk.Checked == true)
                {
                    OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, 1_OLCUT, 2_OLCUT, 3_OLCUT, 4_OLCUT, 5_OLCUT, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5,@P6,@P7,@P8,@P9,@P10)", con);
                    komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                    komutpuankaydet.Parameters.AddWithValue("@D1", derssaati);
                    komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                    komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P2", int.Parse(lblogrencinumara.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                    komutpuankaydet.Parameters.AddWithValue("@P4", 1);
                    komutpuankaydet.Parameters.AddWithValue("@P5", 1);
                    komutpuankaydet.Parameters.AddWithValue("@P6", 1);
                    komutpuankaydet.Parameters.AddWithValue("@P7", 1);
                    komutpuankaydet.Parameters.AddWithValue("@P8", 1);
                    komutpuankaydet.Parameters.AddWithValue("@P9", 5);
                    komutpuankaydet.Parameters.AddWithValue("@P10", hafta);
                    komutpuankaydet.ExecuteNonQuery();
                }
                else if (rdyhaftalik.Checked == true)
                {
                                       for (int i = 1; i <= ogretmenderssayi; i++)
                    {
                        OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, 1_OLCUT, 2_OLCUT, 3_OLCUT, 4_OLCUT, 5_OLCUT, TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5,@P6,@P7,@P8,@P9,@P10)", con);
                        komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                        komutpuankaydet.Parameters.AddWithValue("@D1", i);
                        komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                        komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P2", int.Parse(lblogrencinumara.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                        komutpuankaydet.Parameters.AddWithValue("@P4", 1);
                        komutpuankaydet.Parameters.AddWithValue("@P5", 1);
                        komutpuankaydet.Parameters.AddWithValue("@P6", 1);
                        komutpuankaydet.Parameters.AddWithValue("@P7", 1);
                        komutpuankaydet.Parameters.AddWithValue("@P8", 1);
                        komutpuankaydet.Parameters.AddWithValue("@P9", 5);
                        komutpuankaydet.Parameters.AddWithValue("@P10", hafta);
                        komutpuankaydet.ExecuteNonQuery();
                    }
                }

            }
            else { MessageBox.Show("Lütfen ders programınızdan ders seçiniz.", "Bilgi"); }
            con.Close();
            if (lblogrenciadsoyad.Text != "")
            { degistir(); }
        }

        

        int toplampuani = 5;
        int bir = 1;
        int iki = 1;
        int uc = 1;
        int dort = 1;
        int bes = 1;
        private void button8_Click(object sender, EventArgs e)
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();

            if (lblogrenciadsoyad.Text != "")
            {
                if (rdygunluk.Checked == true)
                {
                    OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, 1_OLCUT, 2_OLCUT, 3_OLCUT, 4_OLCUT, 5_OLCUT,UYARITIPI,UYARIVARYOK,UYARI,TOPLAMPUAN, HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5,@P6,@P7,@P8,@UT,@P9,@P10,@P11,@P12)", con);
                    komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                    komutpuankaydet.Parameters.AddWithValue("@D1", derssaati);
                    komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                    komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P2", int.Parse(lblogrencinumara.Text));
                    komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                    komutpuankaydet.Parameters.AddWithValue("@P4", bir);
                    komutpuankaydet.Parameters.AddWithValue("@P5", iki);
                    komutpuankaydet.Parameters.AddWithValue("@P6", uc);
                    komutpuankaydet.Parameters.AddWithValue("@P7", dort);
                    komutpuankaydet.Parameters.AddWithValue("@P8", bes);
                    if (rchnotdus.Text != "")
                    {
                        if (chkolumluuyari.Visible == true)
                        {
                            komutpuankaydet.Parameters.AddWithValue("@UT", true);
                            komutpuankaydet.Parameters.AddWithValue("@P9", true);
                            komutpuankaydet.Parameters.AddWithValue("@P10", rchnotdus.Text);
                        }
                        else if (chkolumsuzuyari.Visible == true)
                        {
                            komutpuankaydet.Parameters.AddWithValue("@UT", false);
                            komutpuankaydet.Parameters.AddWithValue("@P9", true);
                            komutpuankaydet.Parameters.AddWithValue("@P10", rchnotdus.Text);
                        }

                    }
                    else if (rchnotdus.Text == "")
                    {
                        komutpuankaydet.Parameters.AddWithValue("@UT", false);
                        komutpuankaydet.Parameters.AddWithValue("@P9", false);
                        komutpuankaydet.Parameters.AddWithValue("@P10", "");
                    }
                    komutpuankaydet.Parameters.AddWithValue("@P11", toplampuani);
                    komutpuankaydet.Parameters.AddWithValue("@P12", hafta);
                    komutpuankaydet.ExecuteNonQuery();
                }

                else if (rdyhaftalik.Checked == true)
                {
                    for (int i = 1; i <= ogretmenderssayi; i++)
                    {
                        OleDbCommand komutpuankaydet = new OleDbCommand("insert into TBLNOTLAR (BRANS,DERSSAATI,SINIF,OGRENCININID,OGRENCINUMARASI,TARIH, 1_OLCUT, 2_OLCUT, 3_OLCUT, 4_OLCUT, 5_OLCUT,UYARITIPI,UYARIVARYOK,UYARI,TOPLAMPUAN,HAFTA) VALUES(@P1,@D1,@S1,@O1,@P2,@P3,@P4,@P5,@P6,@P7,@P8,@UT,@P9,@P10,@P11,@P12)", con);
                        komutpuankaydet.Parameters.AddWithValue("@P1", bransid);
                        komutpuankaydet.Parameters.AddWithValue("@D1", i);
                        komutpuankaydet.Parameters.AddWithValue("@S1", sinif);
                        komutpuankaydet.Parameters.AddWithValue("@O1", int.Parse(lblogrenciid.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P2", int.Parse(lblogrencinumara.Text));
                        komutpuankaydet.Parameters.AddWithValue("@P3", dateTimePicker1.Value.ToString());
                        komutpuankaydet.Parameters.AddWithValue("@P4", bir);
                        komutpuankaydet.Parameters.AddWithValue("@P5", iki);
                        komutpuankaydet.Parameters.AddWithValue("@P6", uc);
                        komutpuankaydet.Parameters.AddWithValue("@P7", dort);
                        komutpuankaydet.Parameters.AddWithValue("@P8", bes);
                        if (rchnotdus.Text != "")
                        {
                            if (chkolumluuyari.Visible == true)
                            {
                                komutpuankaydet.Parameters.AddWithValue("@UT", true);
                                komutpuankaydet.Parameters.AddWithValue("@P9", true);
                                komutpuankaydet.Parameters.AddWithValue("@P10", rchnotdus.Text);
                            }
                            else if (chkolumsuzuyari.Visible == true)
                            {
                                komutpuankaydet.Parameters.AddWithValue("@UT", false);
                                komutpuankaydet.Parameters.AddWithValue("@P9", true);
                                komutpuankaydet.Parameters.AddWithValue("@P10", rchnotdus.Text);
                            }

                        }
                        else if (rchnotdus.Text == "")
                        {
                            komutpuankaydet.Parameters.AddWithValue("@UT", false);
                            komutpuankaydet.Parameters.AddWithValue("@P9", false);
                            komutpuankaydet.Parameters.AddWithValue("@P10", "");
                        }
                        komutpuankaydet.Parameters.AddWithValue("@P11", toplampuani);
                        komutpuankaydet.Parameters.AddWithValue("@P12", hafta);
                        komutpuankaydet.ExecuteNonQuery();
                    }


                }
               
            }
            else { MessageBox.Show("Lütfen ders programınızdan ders seçiniz.", "Bilgi"); }
            con.Close();
            if (lblogrenciadsoyad.Text != "")
            { degistir(); }
        }
        

        private void pictureBox6_Click_1(object sender, EventArgs e)
        {
            if (sira == -1)
            {
                if (label18.Text == "ogretmen")
                {
                    this.WindowState = FormWindowState.Minimized;
                    
                }
                else if (label18.Text == "yonetici")
                {
                    this.WindowState = FormWindowState.Minimized;


                }

            }
            else
            {
                MessageBox.Show("Değerlendirmeniz devam ediyor. Lütfen değerlendirme bitmeden sayfadan çıkmayınız.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
        }

        private void pictureBox5_Click_1(object sender, EventArgs e)
        {
            if (sira == -1)
            {
                if (label18.Text == "ogretmen")
                {
                    Application.Exit();
                    
                }
                else if (label18.Text =="yonetici")
                {
                    Application.Exit();
                    

                }
                
            }
            else
            {
                MessageBox.Show("Değerlendirmeniz devam ediyor. Lütfen değerlendirme bitmeden programdan çıkış yapmayınız.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            frmayarlar frmayarlar = new frmayarlar();
            frmayarlar.ogretmenayar = ogretmenid;
            frmayarlar.Show();

        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            if (rol == "ogretmen")
            {
                frmistatistikler frmistatistikler = new frmistatistikler();
                frmistatistikler.rol = "ogretmen";
                frmistatistikler.ogretmennum = ogretmenid;
                frmistatistikler.bransnum = bransid;
                frmistatistikler.Show();
            }
            else if(rol == "yonetici")
            {
                frmistatistikler frmistatistikler = new frmistatistikler();
                frmistatistikler.rol = "yonetici";
                frmistatistikler.Show();
            }


        }

        private void chkdolubir_Click(object sender, EventArgs e)
        {
            if (chkdolubir.Visible == true)
            {


                lblbir.ForeColor = Color.Red;
                toplampuani -= 1;
                bir -= 1;
                chkdolubir.Visible = false;
                chkbosbir.Visible = true;

            }

            label12.Text = toplampuani.ToString();
        }

        private void chkbosbir_Click(object sender, EventArgs e)
        {
            if (chkbosbir.Visible == true)
            {
                lblbir.ForeColor = Color.White;
                toplampuani += 1;
                bir += 1;
                chkbosbir.Visible = false;
                chkdolubir.Visible = true;

            }
            label12.Text = toplampuani.ToString();
        }

        private void chkdoluiki_Click(object sender, EventArgs e)
        {
            if (chkdoluiki.Visible == true)
            {


                lbliki.ForeColor = Color.Red;
                toplampuani -= 1;
                iki -= 1;
                chkdoluiki.Visible = false;
                chkbosiki.Visible = true;

            }

            label12.Text = toplampuani.ToString();
        }

        private void chkbosiki_Click(object sender, EventArgs e)
        {
            if (chkbosiki.Visible == true)
            {
                lbliki.ForeColor = Color.White;
                toplampuani += 1;
                iki += 1;
                chkbosiki.Visible = false;
                chkdoluiki.Visible = true;


            }
            label12.Text = toplampuani.ToString();
        }

        private void chkdoluuc_Click(object sender, EventArgs e)
        {
            if (chkdoluuc.Visible == true)
            {


                lbluc.ForeColor = Color.Red;
                toplampuani -= 1;
                uc -= 1;
                chkdoluuc.Visible = false;
                chkbosuc.Visible = true;
            }

            label12.Text = toplampuani.ToString();
        }

        private void chkbosuc_Click(object sender, EventArgs e)
        {
            if (chkbosuc.Visible == true)
            {
                lbluc.ForeColor = Color.White;
                toplampuani += 1;
                uc += 1;
                chkbosuc.Visible = false;
                chkdoluuc.Visible = true;

            }
            label12.Text = toplampuani.ToString();
        }

        private void chkdoludort_Click(object sender, EventArgs e)
        {

            if (chkdoludort.Visible == true)
            {


                lbldort.ForeColor = Color.Red;
                toplampuani -= 1;
                dort -= 1;
                chkdoludort.Visible = false;
                chkbosdort.Visible = true;
            }
            label12.Text = toplampuani.ToString();
        }

        private void chkbosdort_Click(object sender, EventArgs e)
        {
            if (chkbosdort.Visible == true)
            {
                lbldort.ForeColor = Color.White;
                toplampuani += 1;
                dort += 1;
                chkbosdort.Visible = false;
                chkdoludort.Visible = true;

            }
            label12.Text = toplampuani.ToString();
        }

        private void chkdolubes_Click(object sender, EventArgs e)
        {
            if (chkdolubes.Visible == true)
            {


                lblbes.ForeColor = Color.Red;
                toplampuani -= 1;
                bes -= 1;
                chkdolubes.Visible = false;
                chkbosbes.Visible = true;
            }
            label12.Text = toplampuani.ToString();
        }

        private void chkbosbes_Click(object sender, EventArgs e)
        {
            if (chkbosbes.Visible == true)
            {
                lblbes.ForeColor = Color.White;
                toplampuani += 1;
                bes += 1;
                chkbosbes.Visible = false;
                chkdolubes.Visible = true;

            }
            label12.Text = toplampuani.ToString();
        }
        private void chkolumluuyari_Click_1(object sender, EventArgs e)
        {
            if (chkolumluuyari.Visible == true)
            {
                chkolumluuyari.Visible = false;
                chkolumsuzuyari.Visible = true;
            }
        }

        private void chkolumsuzuyari_Click_1(object sender, EventArgs e)
        {
            if (chkolumsuzuyari.Visible == true)
            {
                chkolumsuzuyari.Visible = false;
                chkolumluuyari.Visible = true;
            }
        }
        

        private void chkpuangor_CheckedChanged(object sender, EventArgs e)
        {
            if (chkpuangor.Checked == true)
            {
                groupBox7.Visible = true;
                chkpuangor.Text = "PUAN HESAPLAMASINI KAPAT";
                chkpuangor.ForeColor = Color.Black;
            }
            else
            {
                chkpuangor.Text = "PUAN HESAPLAMASINI AÇ";
                groupBox7.Visible = false;
                chkpuangor.ForeColor = Color.Red;

            }
        }

        private void chkhesapla_CheckedChanged(object sender, EventArgs e)
        {
            if (chkhesapla.Checked == true)
            {
                chkhesapla.ForeColor = Color.Black;
                chkpuangor.Visible = true;
            }
            else
            {
                chkhesapla.ForeColor = Color.Red;
                chkpuangor.Visible = false;
            }
        }




        private void cmbsınıfsec_SelectedIndexChanged(object sender, EventArgs e)
        {
            
          

            
            OleDbConnection con = new OleDbConnection(conn.baglan);
            if (label18.Text == "ogretmen")
            {

                con.Open();
                OleDbCommand derssaatigetir = new OleDbCommand("select  DERSSAATI, SAATAD FROM TBLDERSPROGRAMI INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID=TBLDERSPROGRAMI.DERSSAATI WHERE TARIH=@P1 AND OGRETMEN=@P2 AND SINIF=@P3 AND OLCDURUM=@P4", con);
                derssaatigetir.Parameters.AddWithValue("@P1", gun);
                derssaatigetir.Parameters.AddWithValue("@P2", ogretmenid);
                derssaatigetir.Parameters.AddWithValue("@P3", sinif);
                derssaatigetir.Parameters.AddWithValue("@P4", false);

                OleDbDataAdapter da5 = new OleDbDataAdapter();
                da5.SelectCommand = derssaatigetir;
                DataTable dt5 = new DataTable();
                da5.Fill(dt5);
                
                con.Close();




            }
            else if (label18.Text == "yonetici")
            {

                con.Open();
                OleDbCommand derssaatigetir2 = new OleDbCommand("select  DERSSAATI, SAATAD FROM TBLDERSPROGRAMI INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID=TBLDERSPROGRAMI.DERSSAATI WHERE TARIH=@k1  AND SINIF=@k3 AND OLCDURUM=@k4", con);
                derssaatigetir2.Parameters.AddWithValue("@k1", gun);
                derssaatigetir2.Parameters.AddWithValue("@k3", sinif);
                derssaatigetir2.Parameters.AddWithValue("@k4", false);

                OleDbDataAdapter da5 = new OleDbDataAdapter();
                da5.SelectCommand = derssaatigetir2;
                DataTable dt5 = new DataTable();
                da5.Fill(dt5);
               
                con.Close();
            }
        }







       

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }
    }
}
