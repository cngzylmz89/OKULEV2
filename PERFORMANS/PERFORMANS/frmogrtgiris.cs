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


namespace PERFORMANS
{
    public partial class frmogrtgiris : Form
    {

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
        public frmogrtgiris()
        {
            InitializeComponent();
        }

        private void btngiris_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand komutgiris = new OleDbCommand("select TCKIMLIKNO, SIFRE FROM TBLOGRETMENLER WHERE TCKIMLIKNO=@P1 AND SIFRE=@P2", con);
            komutgiris.Parameters.AddWithValue("@P1", sifrele(msktc.Text));
            komutgiris.Parameters.AddWithValue("@P2", sifrele(txtsifre.Text));
            OleDbDataReader rd = komutgiris.ExecuteReader();
            if (rd.Read())
            {
                //frmanaform frmanaform = new frmanaform();
                //frmanaform.tc = msktc.Text;
                //frmanaform.rol = "ogretmen";
                //frmanaform.Show();

                frmdersprogrami frmdersprogrami = new frmdersprogrami();
                frmdersprogrami.ogretmentc = msktc.Text;
                frmdersprogrami.Show();
                this.Close();

            }
            else
            {
                MessageBox.Show("Kullanıcı bulunamadı. Lütfen sistem yöneticinize başvurunuz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            con.Close();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.WindowState= FormWindowState.Minimized;
        }

        private void btntcac_Click(object sender, EventArgs e)
        {
            if (msktc.UseSystemPasswordChar == true)
            {
                msktc.UseSystemPasswordChar = false;
            }
            else if(msktc.UseSystemPasswordChar == false)
            {
                msktc.UseSystemPasswordChar = true;
            }
        }

        private void btnsifreac_Click(object sender, EventArgs e)
        {
            if (txtsifre.UseSystemPasswordChar == true)
            {
                txtsifre.UseSystemPasswordChar= false;
            }
            else if(txtsifre.UseSystemPasswordChar== false)
            {
                txtsifre.UseSystemPasswordChar= true;
            }
        }

        private void msktc_MouseClick(object sender, MouseEventArgs e)
        {
            msktc.SelectionStart=msktc.Text.Length;
        }
    }
}
