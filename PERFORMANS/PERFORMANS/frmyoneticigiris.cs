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
using System.Diagnostics.Eventing.Reader;

namespace PERFORMANS
{
    public partial class frmyoneticigiris : Form
    {
        public frmyoneticigiris()
        {
            InitializeComponent();
        }
        public string sifrele(string s)
        {
            byte[] sdizi = ASCIIEncoding.ASCII.GetBytes(s);
            string sifreli = Convert.ToBase64String(sdizi);
            return sifreli;
        }
        baglantisinif conn=new baglantisinif();
        public string sifrecoz(string s)
        {
            byte[] sdizi = Convert.FromBase64String(s);
            string sifresiz = ASCIIEncoding.ASCII.GetString(sdizi);
            return sifresiz;
        }
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\CENGİZ\Desktop\PERFORMANSISTEMI.accdb
        private void btngiris_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand komutgiris = new OleDbCommand("select  TCKIMLIKNO, SIFRE FROM TBLYONETICI WHERE TCKIMLIKNO=@P1 AND SIFRE=@P2", con);
            komutgiris.Parameters.AddWithValue("@P1", sifrele(msktc.Text));
            komutgiris.Parameters.AddWithValue("@P2",sifrele(txtsifre.Text));
            OleDbDataReader rd=komutgiris.ExecuteReader();
            if (rd.Read())
            {
                frmyoneticinanaform frmyoneticinanaform=new frmyoneticinanaform();
                frmyoneticinanaform.Show();
               
            }
            else 
            { 
                MessageBox.Show("Kullanıcı bulunamadı. Lütfen program sağlayıcınıza başvurunuz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            con.Close();
        }

        private void btntcac_Click(object sender, EventArgs e)
        {
            if(msktc.UseSystemPasswordChar==false)
            {msktc.UseSystemPasswordChar= true;}
            else if(msktc.UseSystemPasswordChar==true)
            { msktc.UseSystemPasswordChar= false;}
        }

        private void btnsifreac_Click(object sender, EventArgs e)
        {
            if (txtsifre.UseSystemPasswordChar == false)
            { txtsifre.UseSystemPasswordChar = true; }
            else if (txtsifre.UseSystemPasswordChar == true)
            { txtsifre.UseSystemPasswordChar = false; }
        }

        

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnkayitol_Click(object sender, EventArgs e)
        {
            
            groupBox3.Visible = true;
            btnkayitol.Visible = false;
           frmyoneticigiris frmgiris = new frmyoneticigiris();
            
        }


        private void varolankulkayit_Click(object sender, EventArgs e)
        {
            varolankulkayit.Visible = false ;
            adminkayit.Visible= true ;
            lblkullanici.Text = "T.C. KİMLİK NUMARANIZ:";
            button4.Visible = true;
            button5.Visible = true;
        }
        private void adminkayit_Click(object sender, EventArgs e)
        {
            adminkayit.Visible= false ;
            varolankulkayit.Visible=true ;
            lblkullanici.Text = "KULLANICI ADMİN:";
            button4.Visible = false;
            button5.Visible = false;


        }
        int yanlisat = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            if (adminkayit.Visible == true)
            {
                yanlisat += 1;
                //kullanıcı kayıt
                con.Open();
                OleDbCommand komutgiris = new OleDbCommand("select TCKIMLIKNO, SIFRE FROM TBLYONETICI WHERE TCKIMLIKNO=@P1 AND SIFRE=@P2", con);
                komutgiris.Parameters.AddWithValue("@P1", sifrele(txtkul.Text));
                komutgiris.Parameters.AddWithValue("@P2", sifrele(txtsif.Text));
                OleDbDataReader rd = komutgiris.ExecuteReader();
                if (rd.Read())
                {

                    frmyoneticikayit frmyoneticikayit = new frmyoneticikayit();
                    frmyoneticikayit.Show();
                }
                else
                {
                    if (yanlisat == 3)
                    {
                        MessageBox.Show("Üç kere yanlış bilgi girdiniz.");
                        Application.Exit();
                    }
                    else
                    {
                        MessageBox.Show("Girmiş olduğunuz bilgilere ait hesap bulunamamıştır.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                        


                }
                con.Close();
            }

            if (varolankulkayit.Visible == true)
            {
                yanlisat += 1;
                //admin kayıt
                txtkul.UseSystemPasswordChar = true;
                txtsifre.UseSystemPasswordChar = true;
                
                    string kayıtad = "S22e6L15i12M16";
                string sifre = "y28I11l15m16A1z29";
                if (txtkul.Text == kayıtad && txtsif.Text == sifre)
                {
                    frmyoneticikayit frmyoneticikayit = new frmyoneticikayit();
                    frmyoneticikayit.Show();
                }
                else
                {
                    
                    if (yanlisat == 3)
                    {
                        MessageBox.Show("Üç kere yanlış bilgi girdiniz.");
                        Application.Exit();
                    }
                    else
                    {
                        MessageBox.Show("Kullanıcı bulunamadı.Lütfen program sağlayıcınızın size verdiği kullanıcı adı ve şifreyle kayıt formuna gidiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (txtkul.UseSystemPasswordChar == false)
            { txtkul.UseSystemPasswordChar = true; }
            else if (txtkul.UseSystemPasswordChar == true)
            { txtkul.UseSystemPasswordChar = false; }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (txtsif.UseSystemPasswordChar == false)
            { txtsif.UseSystemPasswordChar = true; }
            else if (txtsif.UseSystemPasswordChar == true)
            { txtsif.UseSystemPasswordChar = false; }
        }

        private void msktc_MouseClick(object sender, MouseEventArgs e)
        {
            msktc.SelectionStart=msktc.Text.Length;
        }
    }
}
