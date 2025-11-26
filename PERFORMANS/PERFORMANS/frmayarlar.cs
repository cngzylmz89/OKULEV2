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
    public partial class frmayarlar : Form
    {
        public frmayarlar()
        {
            InitializeComponent();
        }
        baglantisinif conn= new baglantisinif();

        
        public int ogretmenayar;
        
        private void frmayarlar_Load(object sender, EventArgs e)
        {
            OleDbConnection con=new OleDbConnection(conn.baglan);
            con.Open(); 
            OleDbCommand komutolcutoku=new OleDbCommand("select OLCUTBIR, OLCUTIKI, OLCUTUC, OLCUTDORT,OLCUTBES FROM TBLOGRETMENLER WHERE OGRETMENID="+ogretmenayar,con);
            OleDbDataReader rd=komutolcutoku.ExecuteReader();
            while (rd.Read())
            {
                rcholcutbir.Text = rd[0].ToString();
                rcholcutiki.Text = rd[1].ToString();
                rcholcutuc.Text = rd[2].ToString();
                rcholcutdort.Text = rd[3].ToString();
                rcholcutbes.Text = rd[4].ToString();
            }
            con.Close();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(conn.baglan);
            DialogResult result=MessageBox.Show("Olcutler güncellenecek. Onaylıyor musunuz?","Soru",MessageBoxButtons.YesNo, MessageBoxIcon.Question); 
            con.Open();
            OleDbCommand komutolcutguncelle=new OleDbCommand("update TBLOGRETMENLER SET OLCUTBIR=@P1, OLCUTIKI=@P2, OLCUTUC=@P3, OLCUTDORT=@P4, OLCUTBES=@P5 WHERE OGRETMENID="+ogretmenayar,con);
            komutolcutguncelle.Parameters.AddWithValue("@P1", rcholcutbir.Text);
            komutolcutguncelle.Parameters.AddWithValue("@P2", rcholcutiki.Text);
            komutolcutguncelle.Parameters.AddWithValue("@P3", rcholcutuc.Text);
            komutolcutguncelle.Parameters.AddWithValue("@P4", rcholcutdort.Text);
            komutolcutguncelle.Parameters.AddWithValue("@P5", rcholcutbes.Text);
            if (rcholcutbir.Text != "" && rcholcutiki.Text != "" && rcholcutuc.Text != "" && rcholcutdort.Text != "" && rcholcutbes.Text != "") 
            {
                if (result == DialogResult.Yes)
                {
                    komutolcutguncelle.ExecuteNonQuery();
                    MessageBox.Show("Olcutler güncellendi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
            }
            else
            {
                MessageBox.Show("Lütfen kutucukları doldurunuz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            con.Close() ;
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.WindowState= FormWindowState.Minimized;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            this.Close();
            
        }
    }
}
