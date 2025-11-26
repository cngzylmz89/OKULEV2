using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PERFORMANS
{
    public partial class frmyoneticikayit : Form
    {
        public frmyoneticikayit()
        {
            InitializeComponent();
        }

        baglantisinif conn=new baglantisinif();
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

        private void btngiris_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);

            con.Open();
            OleDbCommand komutgiris = new OleDbCommand("select TCKIMLIKNO FROM TBLYONETICI WHERE TCKIMLIKNO=@P1", con);
            komutgiris.Parameters.AddWithValue("@P1", sifrele(msktc.Text));
            OleDbDataReader rd = komutgiris.ExecuteReader();
            while (rd.Read())
            {
                a = 1;
            }
            con.Close();


            if (a != 1)
            {
                con.Open();
                OleDbCommand cmd = new OleDbCommand("insert into TBLYONETICI (ADISOYADI,TCKIMLIKNO, SIFRE) VALUES(@P1,@P2,@P3)", con);
                cmd.Parameters.AddWithValue("@P1", txtadısoyadı.Text);
                cmd.Parameters.AddWithValue("@P2", sifrele(msktc.Text));
                cmd.Parameters.AddWithValue("@P3", sifrele(txtsifre.Text));
               


                if (msktc.Text != "" && txtsifre.Text != "" && txtadısoyadı.Text != "")
                {

                    cmd.ExecuteNonQuery();
                    MessageBox.Show(msktc.Text + " T.C. kimlik numaralı yönetici kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {

                    MessageBox.Show("Lütfen ilgili yerleri doldurunuz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                con.Close();
            }
            else
            {
                DialogResult result1 = MessageBox.Show(msktc.Text + " T.C. KİMLİK NUMARALI YÖNETİCİ ZATEN KAYITLI. ŞİFREYİ YENİLEMEK İSTİYOR MUSUNUZ?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result1 == DialogResult.Yes)
                {
                    con.Open();
                    OleDbCommand yoneticisifreguncelle = new OleDbCommand("update TBLYONETICI SET SIFRE=@P1", con);
                    yoneticisifreguncelle.Parameters.AddWithValue("@P1", sifrele(txtsifre.Text));
                    yoneticisifreguncelle.ExecuteNonQuery();
                    MessageBox.Show(msktc.Text+" T.C. kimlik numaralı yöneticinin şifresi güncellendi.", "Bilgi", MessageBoxButtons.OK,  MessageBoxIcon.Information );
                    con.Close() ;

                }
                


            }


        }

        private void msktc_MouseClick(object sender, MouseEventArgs e)
        {
            msktc.SelectionStart = msktc.Text.Length;
        }
    }
}
