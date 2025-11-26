using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PERFORMANS
{
    public partial class frmyoneticinanaform : Form
    {
        public frmyoneticinanaform()
        {
            InitializeComponent();
        }
        baglantisinif conn = new baglantisinif();
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
        int a;


        public int haftaal(DateTime dtPassed)
        {

            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dtPassed, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;

        }
        void ogretmenkayitlistele()
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbDataAdapter ogretmenlistele=new OleDbDataAdapter("select OGRETMENID , TCKIMLIKNO , ADISOYADI ,SIFRE, BRANSADI  from TBLOGRETMENLER INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID=TBLOGRETMENLER.BRANS ORDER BY OGRETMENID ASC",con);
            DataSet ds = new DataSet();
            ogretmenlistele.Fill(ds);
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                row["OGRETMENID"] = row["OGRETMENID"] ;
                row["TCKIMLIKNO"] = sifrecoz(row["TCKIMLIKNO"] as string);
                row["ADISOYADI"] = row["ADISOYADI"];
                row["SIFRE"] = sifrecoz(row["SIFRE"] as string);
                row["BRANSADI"] = row["BRANSADI"];
            }
           
            dataGridView2.DataSource = ds.Tables[0];
            dataGridView2.Columns[0].HeaderText = "ÖĞRETMEN İD";
            dataGridView2.Columns[1].HeaderText = "T.C. KİMLİK NUMARA";
            dataGridView2.Columns[2].HeaderText = "ADI SOYADI";
            dataGridView2.Columns[3].HeaderText = "SIFRE";
            dataGridView2.Columns[4].HeaderText = "BRANS ADI";

        }
        int dersprogramisiraoku;
        void dersprogramilist()
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            
            OleDbDataAdapter dersprogramilistele = new OleDbDataAdapter("SELECT KAYITID AS 'SIRA NUMARASI', GUNLER AS 'GÜN', SINIFAD AS 'SINIF', SAATAD AS 'DERS SAATİ', BRANSADI AS 'DERS ADI', ADISOYADI AS 'ÖĞRETMEN', OLCDURUM AS 'ÖLÇÜM DURUMU' FROM ((   ( (   TBLDERSPROGRAMI INNER JOIN TBLGUN ON TBLGUN.HAFTANINGUNU = TBLDERSPROGRAMI.TARIH   ) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID = TBLDERSPROGRAMI.SINIF) INNER JOIN TBLSAATLER ON TBLSAATLER.SAATID = TBLDERSPROGRAMI.DERSSAATI) INNER JOIN TBLBRANSLAR ON TBLBRANSLAR.BRANSID = TBLDERSPROGRAMI.DERS) INNER JOIN TBLOGRETMENLER ON TBLOGRETMENLER.OGRETMENID = TBLDERSPROGRAMI.OGRETMEN ORDER BY KAYITID ASC", con);
            DataTable dtprogramgor = new DataTable();
            dersprogramilistele.Fill(dtprogramgor);
            dataGridView1.DataSource = dtprogramgor;
            DataGridViewCellStyle style = new DataGridViewCellStyle();
            for(int i= 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (Convert.ToBoolean(dataGridView1.Rows[i].Cells[6].Value) == true)
                {
                    style.BackColor = Color.LightGreen;
                    dataGridView1.Rows[i].DefaultCellStyle = style;
                }
                else
                {
                    style.BackColor = Color.Red;
                    dataGridView1.Rows[i].DefaultCellStyle = style;
                }
            }

            con.Open();
            OleDbCommand siraoku = new OleDbCommand("select  count(*) from TBLDERSPROGRAMI", con);
            OleDbDataReader siraokurd=siraoku.ExecuteReader();
            while(siraokurd.Read())
            {
                dersprogramisiraoku = int.Parse(siraokurd[0].ToString());
                label2.Text=siraokurd[0].ToString();
            }
            con.Close();

            dataGridView1.FirstDisplayedScrollingRowIndex = dersprogramisiraoku;
           
        }

        private void frmyoneticinanaform_Load(object sender, EventArgs e)
        {
            hafta = haftaal(DateTime.Now);
            label1.Text=hafta.ToString();
            ogretmenkayitlistele(); 
            dersprogramilist();
            OleDbConnection con = new OleDbConnection(conn.baglan);
            //GÜNLERİ GETİR
            con.Open();
            OleDbDataAdapter gunoku = new OleDbDataAdapter("select HAFTANINGUNU, GUNLER FROM TBLGUN", con);
            DataTable dt = new DataTable();
            gunoku.Fill(dt);
            cmbgun.DisplayMember = "GUNLER";
            cmbgun.ValueMember = "HAFTANINGUNU";
            cmbgun.DataSource = dt;
            
            con.Close();
            //DERSLERİ GETİR
            con.Open();
            OleDbDataAdapter sinifoku = new OleDbDataAdapter("select SINIFID, SINIFAD FROM TBLSINIFLAR", con);
            DataTable dt2 = new DataTable();
            sinifoku.Fill(dt2);
            cmbsinif.DisplayMember = "SINIFAD";
            cmbsinif.ValueMember = "SINIFID";
            cmbsinif.DataSource = dt2;
            con.Close();
            //SAATLERİ GETİR
            con.Open();
            OleDbDataAdapter saatoku = new OleDbDataAdapter("select SAATID, SAATAD FROM TBLSAATLER", con);
            DataTable dt3 = new DataTable();
            saatoku.Fill(dt3);
            cmbderssaati.DisplayMember = "SAATAD";
            cmbderssaati.ValueMember = "SAATID";
            cmbderssaati.DataSource = dt3;
            con.Close();

            //DERSLERİ GETİR
            con.Open();
            OleDbDataAdapter dersoku = new OleDbDataAdapter("select BRANSID, BRANSADI FROM TBLBRANSLAR", con);
            DataTable dt4 = new DataTable();
            dersoku.Fill(dt4);
            cmbders.DisplayMember = "BRANSADI";
            cmbders.ValueMember = "BRANSID";
            cmbbrans.DisplayMember = "BRANSADI";
            cmbbrans.ValueMember = "BRANSID";
            cmbders.DataSource = dt4;
            cmbbrans.DataSource = dt4;
            con.Close();

            //ÖĞRETMENLERİ GETİR
            con.Open();
            OleDbDataAdapter ogretmenoku = new OleDbDataAdapter("select OGRETMENID, ADISOYADI FROM TBLOGRETMENLER", con);
            DataTable dt5 = new DataTable();
            ogretmenoku.Fill(dt5);
            cmbogretmen.DisplayMember = "ADISOYADI";
            cmbogretmen.ValueMember = "OGRETMENID";
            cmbogretmen.DataSource = dt5;
            con.Close();

            //datagrid1 ders programı listele

           
        }
        string ogretmenders;
        string kayitoku;
        int hafta;
        private void button9_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            hafta = haftaal(DateTime.Now);
            con.Open();
            OleDbCommand derskayitoku = new OleDbCommand("select TARIH, SINIF, DERSSAATI FROM TBLDERSPROGRAMI WHERE TARIH=@P1 AND SINIF=@P2 AND DERSSAATI=@P3", con);
            derskayitoku.Parameters.AddWithValue("@P1", cmbgun.SelectedValue);
            derskayitoku.Parameters.AddWithValue("@P2", cmbsinif.SelectedValue);
            derskayitoku.Parameters.AddWithValue("@P3", cmbderssaati.SelectedValue);

            OleDbDataReader derskayitokurd = derskayitoku.ExecuteReader();
            if (derskayitokurd.Read()==true)
            {
                kayitoku ="true" ;
               MessageBox.Show("Bu ders önceden kaydedildi.");
               

            }
            else
            {
                kayitoku ="false" ;
            }

            con.Close();
            con.Open();
            OleDbCommand ogretmendersoku = new OleDbCommand("select TARIH, DERSSAATI, OGRETMEN FROM TBLDERSPROGRAMI WHERE TARIH=@k1 AND  DERSSAATI=@k3 AND OGRETMEN=@k2", con);
            ogretmendersoku.Parameters.AddWithValue("@k1", cmbgun.SelectedValue);
            ogretmendersoku.Parameters.AddWithValue("@k3", cmbderssaati.SelectedValue);
            ogretmendersoku.Parameters.AddWithValue("@k2", cmbogretmen.SelectedValue);
            

            OleDbDataReader ogretmendersokurd = ogretmendersoku.ExecuteReader();
            if (ogretmendersokurd.Read() == true)
            {
                ogretmenders = "true";
                MessageBox.Show("Bu ogretmene aynı gun ve saatte ders verilmiş.");


            }
            else
            {
                ogretmenders = "false";
            }

            con.Close();


            if (kayitoku == "false"&&ogretmenders=="false")
            {

                con.Open();
                OleDbCommand derskaydet=new OleDbCommand("insert into TBLDERSPROGRAMI (TARIH, SINIF, DERSSAATI, DERS,OGRETMEN, OLCDURUM, HAFTA)  VALUES(@K1,@K2,@K3,@K4, @K5, @K6, @H1)",con) ;
                derskaydet.Parameters.AddWithValue("@K1", cmbgun.SelectedValue);
                derskaydet.Parameters.AddWithValue("@K2", cmbsinif.SelectedValue);
                derskaydet.Parameters.AddWithValue("@K3", cmbderssaati.SelectedValue);
                derskaydet.Parameters.AddWithValue("@K4", cmbders.SelectedValue);
                derskaydet.Parameters.AddWithValue("@K5", cmbogretmen.SelectedValue);
                derskaydet.Parameters.AddWithValue("@K6", false);
                derskaydet.Parameters.AddWithValue("@H1", hafta);
                derskaydet.ExecuteNonQuery() ;
                
                

            }
            
            con.Close();
            dersprogramilist();
        }

       

        private void button12_Click(object sender, EventArgs e)
        {
            dersprogramilist();
        }
        int kayıtid;
        private void button10_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            DialogResult result = MessageBox.Show(kayıtid + " numaralı kayıt güncellenecek onaylıyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            con.Open();
            OleDbCommand komutguncelle=new OleDbCommand("update TBLDERSPROGRAMI SET TARIH=@P1, SINIF=@P2, DERSSAATI=@P3, DERS=@P4,OGRETMEN=@P  WHERE KAYITID=@K1", con);
            komutguncelle.Parameters.AddWithValue("@P1", cmbgun.SelectedValue.ToString());
            komutguncelle.Parameters.AddWithValue("@P2", cmbsinif.SelectedValue);
            komutguncelle.Parameters.AddWithValue("@P3", cmbderssaati.SelectedValue);
            komutguncelle.Parameters.AddWithValue("@P4", cmbders.SelectedValue);
            komutguncelle.Parameters.AddWithValue("@P5", cmbogretmen.SelectedValue);
            komutguncelle.Parameters.AddWithValue("@K1", kayıtid);
            if(kayıtid != null)
            {
                if (result == DialogResult.Yes)
                {
                    komutguncelle.ExecuteNonQuery();
                    MessageBox.Show("Ders güncellendi.");
                }


              
            }
            con.Close();
            dersprogramilist();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                kayıtid = int.Parse(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                cmbgun.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                cmbsinif.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                cmbderssaati.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                cmbders.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                cmbogretmen.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            }
            catch (Exception)
            {

                MessageBox.Show("Lütfen listeye tıklayın");
            }
           
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            DialogResult result = MessageBox.Show(kayıtid + " numaralı kayıt silinecek onaylıyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            con.Open();
            OleDbCommand komutguncelle = new OleDbCommand("DELETE FROM TBLDERSPROGRAMI WHERE KAYITID=@K1", con);
           
            komutguncelle.Parameters.AddWithValue("@K1", kayıtid);
            if (result == DialogResult.Yes&&kayıtid!=null)
            {
                komutguncelle.ExecuteNonQuery();
                MessageBox.Show("Ders silindi.");
            }


            con.Close();
            dersprogramilist();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dersprogramilist();
            groupBox1.Visible = true;
            groupBox4.Visible=false;
           
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            groupBox1.Visible=false;
            groupBox4.Visible=true;
          
        }
        string ogretmen;
        private void btnkaydet_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand komutkayitoku = new OleDbCommand("select * from TBLOGRETMENLER WHERE TCKIMLIKNO=@P1", con);
            komutkayitoku.Parameters.AddWithValue("@P1", sifrele(mskogrttc.Text));
            OleDbDataReader komutkayitokurd=komutkayitoku.ExecuteReader();
            if (komutkayitokurd.Read()==true)
            {
                ogretmen = "true";
                MessageBox.Show(mskogrttc.Text + " T.C. kimlik numaralı öğretmen zaten kayıtlı");

            }
            else
            {
                ogretmen="false";
            }

                con.Close();

            
          
            if (ogretmen == "false"  )
            {
                if (mskogrttc.Text != "" && txtogrtsif.Text != "" && txtogrtadsoyad.Text != "")
                {
                    DialogResult result = MessageBox.Show(mskogrttc.Text + " T.C. kimlik numaralı öğretmen kaydedilecek. Onaylıyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        con.Open();
                        OleDbCommand komutkaydet = new OleDbCommand("insert into TBLOGRETMENLER (TCKIMLIKNO, SIFRE, ADISOYADI,BRANS) VALUES(@K1, @K2, @K3, @K4) ", con);
                        komutkaydet.Parameters.AddWithValue("@K1", sifrele(mskogrttc.Text));
                        komutkaydet.Parameters.AddWithValue("@K2", sifrele(txtogrtsif.Text));
                        komutkaydet.Parameters.AddWithValue("@K3", txtogrtadsoyad.Text);
                        komutkaydet.Parameters.AddWithValue("@K4", cmbbrans.SelectedValue);
                        komutkaydet.ExecuteNonQuery();
                        MessageBox.Show(mskogrttc.Text + "  T.C. kimlik numaralı öğretmen kaydedildi.");
                        con.Close();
                    }
                }
                else { MessageBox.Show("Lütfen ilgili yerleri doldurunuz."); }
               
               
            }
            
                
            ogretmenkayitlistele();
        }

        private void mskogrttc_MouseClick(object sender, MouseEventArgs e)
        {
            mskogrttc.SelectionStart = mskogrttc.Text.Length;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (mskogrttc.UseSystemPasswordChar == false)
            {
                mskogrttc.UseSystemPasswordChar = true;
            }
            else if (mskogrttc.UseSystemPasswordChar == true)
            { mskogrttc.UseSystemPasswordChar = false; }
           
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (txtogrtsif.UseSystemPasswordChar == false)
            {
                txtogrtsif.UseSystemPasswordChar = true;
            }
            else if (txtogrtsif.UseSystemPasswordChar == true)
            { txtogrtsif.UseSystemPasswordChar = false; }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmistatistikler frmistatistikler = new frmistatistikler();
            frmistatistikler.rol = "yonetici";
            frmistatistikler.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            frmenler frmenler = new frmenler();
            frmenler.Show();
        }

        

        private void button3_Click(object sender, EventArgs e)
        {
            frmanaform frmanaform = new frmanaform();
            frmanaform.rol = "yonetici";
            frmanaform.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmpuankayitlari frmpuankayitlari = new frmpuankayitlari();
            frmpuankayitlari.datasec = "puankayitlari";
            frmpuankayitlari.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmpuankayitlari frmpuankayitlari = new frmpuankayitlari();
            frmpuankayitlari.datasec = "puanlanmamisdersler";
            frmpuankayitlari.Show();
        }
        public string haftaacoku;
        private void btnyenihaftaac_Click(object sender, EventArgs e)
        {
            hafta=haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand komuthaftaacoku = new OleDbCommand("select *  from TBLDERSPROGRAMI WHERE HAFTA=@H", con);
            komuthaftaacoku.Parameters.AddWithValue("@H", hafta);  
            OleDbDataReader komuthaftaacokurd=komuthaftaacoku.ExecuteReader();
            if(komuthaftaacokurd.Read()==false)
            {
                haftaacoku = "musait";
            }
            else
            {
                haftaacoku = "namusait";
            }
            con.Close();

            DialogResult result = MessageBox.Show("Hafta açılacak onaylıyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (haftaacoku == "musait")
            {
                if(result == DialogResult.Yes)
                {
                    con.Open();
                    OleDbCommand haftaacguncelle = new OleDbCommand("update TBLDERSPROGRAMI SET OLCDURUM='0', HAFTA="+hafta, con);
                    haftaacguncelle.ExecuteNonQuery();
                    MessageBox.Show("Hafta açıldı.");
                    con.Close();
                }
                
            }
            else
            {
                MessageBox.Show("Hafta daha önce açılmış", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
           
        }

        private void button7_Click(object sender, EventArgs e)
        {

           
                cmbuyari.Visible = true;
                if (cmbuyari.SelectedItem == "BU HAFTA")
                {
                    frmdatagrid2 frmdatagrid2 = new frmdatagrid2();
                    frmdatagrid2.sec = 3;
               cmbuyari.SelectedItem=null;
                cmbuyari.Visible = false;
                    frmdatagrid2.Show();
                }
                else if (cmbuyari.SelectedItem == "HEPSİ")
                {
                    frmdatagrid2 frmdatagrid2 = new frmdatagrid2();
                    frmdatagrid2.sec = 4;
                   cmbuyari.SelectedItem=null;
                    cmbuyari.Visible = false;
                    frmdatagrid2.Show();
                }
           
            
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txtogrtid.Text = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                mskogrttc.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                txtogrtsif.Text =dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                txtogrtadsoyad.Text = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                cmbbrans.Text = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
            }
            catch (Exception)
            {

                MessageBox.Show("Lütfen tablo hücrelerini tıklayınız.");
            }
           
        }
        int ogrtid;
        private void button4_Click_1(object sender, EventArgs e)
        {
            ogrtid=byte.Parse(txtogrtid.Text);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            DialogResult result = MessageBox.Show(txtogrtadsoyad.Text + " adlı öğretmen güncellenecek. Onaylıyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                con.Open();
                OleDbCommand ogretmenguncelle = new OleDbCommand("update TBLOGRETMENLER SET TCKIMLIKNO=@P1, SIFRE=@P2, ADISOYADI=@P3, BRANS=@P4 WHERE OGRETMENID=@K1", con);
                
                ogretmenguncelle.Parameters.AddWithValue("@P1", sifrele(mskogrttc.Text));
                ogretmenguncelle.Parameters.AddWithValue("@P2", sifrele(txtogrtsif.Text));
                ogretmenguncelle.Parameters.AddWithValue("@P3", txtogrtadsoyad.Text);
                ogretmenguncelle.Parameters.AddWithValue("@P4", cmbbrans.SelectedValue);
            ogretmenguncelle.Parameters.AddWithValue("@K1", ogrtid);
            if (txtogrtid.Text != "" && result == DialogResult.Yes)
                {
                    ogretmenguncelle.ExecuteNonQuery();
                    MessageBox.Show(txtogrtadsoyad.Text + " adlı öğretmen bilgileri güncellendi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Lütfen öğretmen tablosundan öğretmen seçiniz.");
                }

                con.Close();
                ogretmenkayitlistele();
            
        }

        private void btnogretmendurum_Click(object sender, EventArgs e)
        {
            frmogretmendurum frmogretmendurum = new frmogretmendurum();
            frmogretmendurum.Show();

        }
    }
}
