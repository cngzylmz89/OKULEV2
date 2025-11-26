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
using System.Xml.Linq;

namespace PERFORMANS
{
    public partial class frmdersprogrami : Form
    {
        public frmdersprogrami()
        {
            InitializeComponent();
        }
        public string rol;
        public string ogretmentc;
        baglantisinif conn=new baglantisinif();
        public int ogretmenid;
        int kacincigun;


        int monday1dersid;
        int monday1sinifid;
        int monday1kayitid;

        int monday2dersid;
        int monday2sinifid;
        int monday2kayitid;

        int monday3dersid;
        int monday3sinifid;
        int monday3kayitid;

        int monday4dersid;
        int monday4sinifid;
        int monday4kayitid;

        int monday5dersid;
        int monday5sinifid;
        int monday5kayitid;

        int monday6dersid;
        int monday6sinifid;
        int monday6kayitid;

        int monday7dersid;
        int monday7sinifid;
        int monday7kayitid;

        int tuesday1dersid;
        int tuesday1sinifid;
        int tuesday1kayitid;

        int tuesday2dersid;
        int tuesday2sinifid;
        int tuesday2kayitid;

        int tuesday3dersid;
        int tuesday3sinifid;
        int tuesday3kayitid;

        int tuesday4dersid;
        int tuesday4sinifid;
        int tuesday4kayitid;

        int tuesday5dersid;
        int tuesday5sinifid;
        int tuesday5kayitid;

        int tuesday6dersid;
        int tuesday6sinifid;
        int tuesday6kayitid;

        int tuesday7dersid;
        int tuesday7sinifid;
        int tuesday7kayitid;

        int wednesday1dersid;
        int wednesday1sinifid;
        int wednesday1kayitid;

        int wednesday2dersid;
        int wednesday2sinifid;
        int wednesday2kayitid;

        int wednesday3dersid;
        int wednesday3sinifid;
        int wednesday3kayitid;

        int wednesday4dersid;
        int wednesday4sinifid;
        int wednesday4kayitid;

        int wednesday5dersid;
        int wednesday5sinifid;
        int wednesday5kayitid;

        int wednesday6dersid;
        int wednesday6sinifid;
        int wednesday6kayitid;

        int wednesday7dersid;
        int wednesday7sinifid;
        int wednesday7kayitid;

        int thursday1dersid;
        int thursday1sinifid;
        int thursday1kayitid;

        int thursday2dersid;
        int thursday2sinifid;
        int thursday2kayitid;

        int thursday3dersid;
        int thursday3sinifid;
        int thursday3kayitid;

        int thursday4dersid;
        int thursday4sinifid;
        int thursday4kayitid;

        int thursday5dersid;
        int thursday5sinifid;
        int thursday5kayitid;

        int thursday6dersid;
        int thursday6sinifid;
        int thursday6kayitid;

        int thursday7dersid;
        int thursday7sinifid;
        int thursday7kayitid;

        int friday1dersid;
        int friday1sinifid;
        int friday1kayitid;

        int friday2dersid;
        int friday2sinifid;
        int friday2kayitid;

        int friday3dersid;
        int friday3sinifid;
        int friday3kayitid;

        int friday4dersid;
        int friday4sinifid;
        int friday4kayitid;

        int friday5dersid;
        int friday5sinifid;
        int friday5kayitid;

        int friday6dersid;
        int friday6sinifid;
        int friday6kayitid;

        int friday7dersid;
        int friday7sinifid;
        int friday7kayitid;

        int toplamderssayisi;
        int puanlanmisderssayisi;
        decimal yuzdelik;
        

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

        void dersprogramigetir()
        {
            kacincigun = (int)DateTime.Now.DayOfWeek;
            switch (kacincigun)
            {
                case 1:label1.ForeColor = Color.White; 
                   label1.Font = new Font(label1.Font, FontStyle.Bold);
                    label1.BackColor=Color.RoyalBlue;
                    break;
                case 2: label2.ForeColor = Color.White;
                    label2.Font = new Font(label2.Font, FontStyle.Bold);
                    label2.BackColor=Color.RoyalBlue;
                    break;
                case 3: label3.ForeColor = Color.White; 
                    label3.Font = new Font(label3.Font, FontStyle.Bold);
                    label3.BackColor=Color.RoyalBlue;
                    break;
                case 4: label4.ForeColor = Color.White; 
                    label4.Font = new Font(label4.Font, FontStyle.Bold);
                    label4.BackColor=Color.RoyalBlue;
                    break;
                case 5: label5.ForeColor = Color.White; 
                    label5.Font = new Font(label5.Font, FontStyle.Bold); 
                    label5.BackColor=Color.RoyalBlue;
                    break;
            }
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand komutogrtid = new OleDbCommand("select OGRETMENID FROM TBLOGRETMENLER WHERE TCKIMLIKNO=@P1", con);
            komutogrtid.Parameters.AddWithValue("@P1", sifrele(ogretmentc));
            OleDbDataReader drid = komutogrtid.ExecuteReader();
            while (drid.Read())
            {
                ogretmenid = int.Parse(drid[0].ToString());
            }
            con.Close();
            //pazartesi 1.saat
            con.Open();
            OleDbCommand monday1 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday1.Parameters.AddWithValue("@P1", "Monday");
            monday1.Parameters.AddWithValue("@P2", ogretmenid);
            monday1.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader dr1 = monday1.ExecuteReader();
            while (dr1.Read())
            {
                btnmonday1.Text = dr1[0].ToString()+" " + dr1[3].ToString();
                monday1dersid = int.Parse(dr1[1].ToString());
                monday1sinifid = int.Parse(dr1[2].ToString());
                monday1kayitid = int.Parse(dr1[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand monday1olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday1olc.Parameters.AddWithValue("@P1", "Monday");
            monday1olc.Parameters.AddWithValue("@P2", ogretmenid);
            monday1olc.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader dr1olc = monday1olc.ExecuteReader();
            while (dr1olc.Read())
            {
                if (dr1olc[0].ToString() == "True")
                {
                    btnmonday1.BackColor = Color.Green;
                }
                else
                {
                    btnmonday1.BackColor = Color.Orange;
                }
            }
            con.Close();
           
            //pazartesi 2.saat



            con.Open();
            OleDbCommand monday2 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday2.Parameters.AddWithValue("@P1", "Monday");
            monday2.Parameters.AddWithValue("@P2", ogretmenid);
            monday2.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader dr2 = monday2.ExecuteReader();
            while (dr2.Read())
            {
                btnmonday2.Text = dr2[0].ToString()+ " " + dr2[3].ToString();
                monday2dersid = int.Parse(dr2[1].ToString());
                monday2sinifid = int.Parse(dr2[2].ToString());
                monday2kayitid = int.Parse(dr2[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand monday2olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday2olc.Parameters.AddWithValue("@P1", "Monday");
            monday2olc.Parameters.AddWithValue("@P2", ogretmenid);
            monday2olc.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader dr2olc = monday2olc.ExecuteReader();
            while (dr2olc.Read())
            {
                if (dr2olc[0].ToString() == "True")
                {
                    btnmonday2.BackColor = Color.Green;
                }
                else
                {
                    btnmonday2.BackColor = Color.Orange;
                }
            }
            con.Close();
           

            //pazartesi 3.saat



            con.Open();
            OleDbCommand monday3 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday3.Parameters.AddWithValue("@P1", "Monday");
            monday3.Parameters.AddWithValue("@P2", ogretmenid);
            monday3.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader dr3 = monday3.ExecuteReader();
            while (dr3.Read())
            {
                btnmonday3.Text = dr3[0].ToString()+ " " + dr3[3].ToString();
                monday3dersid = int.Parse(dr3[1].ToString());
                monday3sinifid = int.Parse(dr3[2].ToString());
                monday3kayitid = int.Parse(dr3[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand monday3olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday3olc.Parameters.AddWithValue("@P1", "Monday");
            monday3olc.Parameters.AddWithValue("@P2", ogretmenid);
            monday3olc.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader dr3olc = monday3olc.ExecuteReader();
            while (dr3olc.Read())
            {
                if (dr3olc[0].ToString() == "True")
                {
                    btnmonday3.BackColor = Color.Green;
                }
                else
                {
                    btnmonday3.BackColor = Color.Orange;
                }
            }
            con.Close();
            


            //pazartesi 4.saat



            con.Open();
            OleDbCommand monday4 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday4.Parameters.AddWithValue("@P1", "Monday");
            monday4.Parameters.AddWithValue("@P2", ogretmenid);
            monday4.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader dr4 = monday4.ExecuteReader();
            while (dr4.Read())
            {
                btnmonday4.Text = dr4[0].ToString()+ " " + dr4[3].ToString();
                monday4dersid = int.Parse(dr4[1].ToString());
                monday4sinifid = int.Parse(dr4[2].ToString());
                monday4kayitid = int.Parse(dr4[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand monday4olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday4olc.Parameters.AddWithValue("@P1", "Monday");
            monday4olc.Parameters.AddWithValue("@P2", ogretmenid);
            monday4olc.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader dr4olc = monday4olc.ExecuteReader();
            while (dr4olc.Read())
            {
                if (dr4olc[0].ToString() == "True")
                {
                    btnmonday4.BackColor = Color.Green;
                }
                else
                {
                    btnmonday4.BackColor = Color.Orange;
                }
            }
            con.Close();

            //pazartesi 5.saat



            con.Open();
            OleDbCommand monday5 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday5.Parameters.AddWithValue("@P1", "Monday");
            monday5.Parameters.AddWithValue("@P2", ogretmenid);
            monday5.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader dr5 = monday5.ExecuteReader();
            while (dr5.Read())
            {
                btnmonday5.Text = dr5[0].ToString() + " " + dr5[3].ToString();
                monday5dersid = int.Parse(dr5[1].ToString());
                monday5sinifid = int.Parse(dr5[2].ToString());
                monday5kayitid = int.Parse(dr5[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand monday5olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday5olc.Parameters.AddWithValue("@P1", "Monday");
            monday5olc.Parameters.AddWithValue("@P2", ogretmenid);
            monday5olc.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader dr5olc = monday5olc.ExecuteReader();
            while (dr5olc.Read())
            {
                if (dr5olc[0].ToString() == "True")
                {
                    btnmonday5.BackColor = Color.Green;
                }
                else
                {
                    btnmonday5.BackColor = Color.Orange;
                }
            }
            con.Close();


            //pazartesi 6.saat



            con.Open();
            OleDbCommand monday6 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday6.Parameters.AddWithValue("@P1", "Monday");
            monday6.Parameters.AddWithValue("@P2", ogretmenid);
            monday6.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader dr6 = monday6.ExecuteReader();
            while (dr6.Read())
            {
                btnmonday6.Text = dr6[0].ToString() + " " + dr6[3].ToString();
                monday6dersid = int.Parse(dr6[1].ToString());
                monday6sinifid = int.Parse(dr6[2].ToString());
                monday6kayitid = int.Parse(dr6[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand monday6olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday6olc.Parameters.AddWithValue("@P1", "Monday");
            monday6olc.Parameters.AddWithValue("@P2", ogretmenid);
            monday6olc.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader dr6olc = monday6olc.ExecuteReader();
            while (dr6olc.Read())
            {
                if (dr6olc[0].ToString() == "True")
                {
                    btnmonday6.BackColor = Color.Green;
                }
                else
                {
                    btnmonday6.BackColor = Color.Orange;
                }
            }
            con.Close();

            //pazartesi 7.saat



            con.Open();
            OleDbCommand monday7 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday7.Parameters.AddWithValue("@P1", "Monday");
            monday7.Parameters.AddWithValue("@P2", ogretmenid);
            monday7.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader dr7 = monday7.ExecuteReader();
            while (dr7.Read())
            {
                btnmonday7.Text = dr7[0].ToString() + " " + dr7[3].ToString();
                monday7dersid = int.Parse(dr7[1].ToString());
                monday7sinifid = int.Parse(dr7[2].ToString());
                monday7kayitid = int.Parse(dr7[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand monday7olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            monday7olc.Parameters.AddWithValue("@P1", "Monday");
            monday7olc.Parameters.AddWithValue("@P2", ogretmenid);
            monday7olc.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader dr7olc = monday7olc.ExecuteReader();
            while (dr7olc.Read())
            {
                if (dr7olc[0].ToString() == "True")
                {
                    btnmonday7.BackColor = Color.Green;
                }
                else
                {
                    btnmonday7.BackColor = Color.Orange;
                }
            }
            con.Close();

            //salı 1.saat
            con.Open();
            OleDbCommand tuesday1 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday1.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday1.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday1.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader drtuesday1 = tuesday1.ExecuteReader();
            while (drtuesday1.Read())
            {
                btntuesday1.Text = drtuesday1[0].ToString() + " " + drtuesday1[3].ToString();
                tuesday1dersid = int.Parse(drtuesday1[1].ToString());
                tuesday1sinifid = int.Parse(drtuesday1[2].ToString());
                tuesday1kayitid = int.Parse(drtuesday1[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand tuesday1olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday1olc.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday1olc.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday1olc.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader drtuesday1olc = tuesday1olc.ExecuteReader();
            while (drtuesday1olc.Read())
            {
                if (drtuesday1olc[0].ToString() == "True")
                {
                    btntuesday1.BackColor = Color.Green;
                }
                else
                {
                    btntuesday1.BackColor = Color.Orange;
                }
            }
            con.Close();

            //salı 2.saat
            con.Open();
            OleDbCommand tuesday2 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday2.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday2.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday2.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader drtuesday2 = tuesday2.ExecuteReader();
            while (drtuesday2.Read())
            {
                btntuesday2.Text = drtuesday2[0].ToString() + " " + drtuesday2[3].ToString();
                tuesday2dersid = int.Parse(drtuesday2[1].ToString());
                tuesday2sinifid = int.Parse(drtuesday2[2].ToString());
                tuesday2kayitid = int.Parse(drtuesday2[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand tuesday2olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday2olc.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday2olc.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday2olc.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader drtuesday2olc = tuesday2olc.ExecuteReader();
            while (drtuesday2olc.Read())
            {
                if (drtuesday2olc[0].ToString() == "True")
                {
                    btntuesday2.BackColor = Color.Green;
                }
                else
                {
                    btntuesday2.BackColor = Color.Orange;
                }
            }
            con.Close();

            //salı 3.saat
            con.Open();
            OleDbCommand tuesday3 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday3.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday3.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday3.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader drtuesday3 = tuesday3.ExecuteReader();
            while (drtuesday3.Read())
            {
                btntuesday3.Text = drtuesday3[0].ToString() + " " + drtuesday3[3].ToString();
                tuesday3dersid = int.Parse(drtuesday3[1].ToString());
                tuesday3sinifid = int.Parse(drtuesday3[2].ToString());
                tuesday3kayitid = int.Parse(drtuesday3[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand tuesday3olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday3olc.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday3olc.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday3olc.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader drtuesday3olc = tuesday3olc.ExecuteReader();
            while (drtuesday3olc.Read())
            {
                if (drtuesday3olc[0].ToString() == "True")
                {
                    btntuesday3.BackColor = Color.Green;
                }
                else
                {
                    btntuesday3.BackColor = Color.Orange;
                }
            }
            con.Close();

            //salı 4.saat
            con.Open();
            OleDbCommand tuesday4 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday4.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday4.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday4.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader drtuesday4 = tuesday4.ExecuteReader();
            while (drtuesday4.Read())
            {
                btntuesday4.Text = drtuesday4[0].ToString() + " " + drtuesday4[3].ToString();
                tuesday4dersid = int.Parse(drtuesday4[1].ToString());
                tuesday4sinifid = int.Parse(drtuesday4[2].ToString());
                tuesday4kayitid = int.Parse(drtuesday4[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand tuesday4olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday4olc.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday4olc.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday4olc.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader drtuesday4olc = tuesday4olc.ExecuteReader();
            while (drtuesday4olc.Read())
            {
                if (drtuesday4olc[0].ToString() == "True")
                {
                    btntuesday4.BackColor = Color.Green;
                }
                else
                {
                    btntuesday4.BackColor = Color.Orange;
                }
            }
            con.Close();

            //salı 5.saat
            con.Open();
            OleDbCommand tuesday5 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday5.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday5.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday5.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader drtuesday5 = tuesday5.ExecuteReader();
            while (drtuesday5.Read())
            {
                btntuesday5.Text = drtuesday5[0].ToString() + " " + drtuesday5[3].ToString();
                tuesday5dersid = int.Parse(drtuesday5[1].ToString());
                tuesday5sinifid = int.Parse(drtuesday5[2].ToString());
                tuesday5kayitid = int.Parse(drtuesday5[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand tuesday5olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday5olc.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday5olc.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday5olc.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader drtuesday5olc = tuesday5olc.ExecuteReader();
            while (drtuesday5olc.Read())
            {
                if (drtuesday5olc[0].ToString() == "True")
                {
                    btntuesday5.BackColor = Color.Green;
                }
                else
                {
                    btntuesday5.BackColor = Color.Orange;
                }
            }
            con.Close();

            //salı 6.saat
            con.Open();
            OleDbCommand tuesday6 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday6.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday6.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday6.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader drtuesday6 = tuesday6.ExecuteReader();
            while (drtuesday6.Read())
            {
                btntuesday6.Text = drtuesday6[0].ToString() + " " + drtuesday6[3].ToString();
                tuesday6dersid = int.Parse(drtuesday6[1].ToString());
                tuesday6sinifid = int.Parse(drtuesday6[2].ToString());
                tuesday6kayitid = int.Parse(drtuesday6[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand tuesday6olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday6olc.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday6olc.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday6olc.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader drtuesday6olc = tuesday6olc.ExecuteReader();
            while (drtuesday6olc.Read())
            {
                if (drtuesday6olc[0].ToString() == "True")
                {
                    btntuesday6.BackColor = Color.Green;
                }
                else
                {
                    btntuesday6.BackColor = Color.Orange;
                }
            }
            con.Close();

            //salı 7.saat
            con.Open();
            OleDbCommand tuesday7 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday7.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday7.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday7.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader drtuesday7 = tuesday7.ExecuteReader();
            while (drtuesday7.Read())
            {
                btntuesday7.Text = drtuesday7[0].ToString() + " " + drtuesday7[3].ToString();
                tuesday7dersid = int.Parse(drtuesday7[1].ToString());
                tuesday7sinifid = int.Parse(drtuesday7[2].ToString());
                tuesday7kayitid = int.Parse(drtuesday7[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand tuesday7olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            tuesday7olc.Parameters.AddWithValue("@P1", "Tuesday");
            tuesday7olc.Parameters.AddWithValue("@P2", ogretmenid);
            tuesday7olc.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader drtuesday7olc = tuesday7olc.ExecuteReader();
            while (drtuesday7olc.Read())
            {
                if (drtuesday7olc[0].ToString() == "True")
                {
                    btntuesday7.BackColor = Color.Green;
                }
                else
                {
                    btntuesday7.BackColor = Color.Orange;
                }
            }
            con.Close();

            //çarşamba 1.saat
            con.Open();
            OleDbCommand wednesday1 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday1.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday1.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday1.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader drwednesday1 = wednesday1.ExecuteReader();
            while (drwednesday1.Read())
            {
                btnwednesday1.Text = drwednesday1[0].ToString() + " " + drwednesday1[3].ToString();
                wednesday1dersid = int.Parse(drwednesday1[1].ToString());
                wednesday1sinifid = int.Parse(drwednesday1[2].ToString());
                wednesday1kayitid = int.Parse(drwednesday1[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand wednesday1olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday1olc.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday1olc.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday1olc.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader drwednesday1olc = wednesday1olc.ExecuteReader();
            while (drwednesday1olc.Read())
            {
                if (drwednesday1olc[0].ToString() == "True")
                {
                    btnwednesday1.BackColor = Color.Green;
                }
                else
                {
                    btnwednesday1.BackColor = Color.Orange;
                }
            }
            con.Close();

            //çarşamba 2.saat
            con.Open();
            OleDbCommand wednesday2 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday2.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday2.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday2.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader drwednesday2 = wednesday2.ExecuteReader();
            while (drwednesday2.Read())
            {
                btnwednesday2.Text = drwednesday2[0].ToString() + " " + drwednesday2[3].ToString();
                wednesday2dersid = int.Parse(drwednesday2[1].ToString());
                wednesday2sinifid = int.Parse(drwednesday2[2].ToString());
                wednesday2kayitid = int.Parse(drwednesday2[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand wednesday2olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday2olc.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday2olc.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday2olc.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader drwednesday2olc = wednesday2olc.ExecuteReader();
            while (drwednesday2olc.Read())
            {
                if (drwednesday2olc[0].ToString() == "True")
                {
                    btnwednesday2.BackColor = Color.Green;
                }
                else
                {
                    btnwednesday2.BackColor = Color.Orange;
                }
            }
            con.Close();

            //çarşamba 3.saat
            con.Open();
            OleDbCommand wednesday3 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday3.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday3.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday3.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader drwednesday3 = wednesday3.ExecuteReader();
            while (drwednesday3.Read())
            {
                btnwednesday3.Text = drwednesday3[0].ToString() + " " + drwednesday3[3].ToString();
                wednesday3dersid = int.Parse(drwednesday3[1].ToString());
                wednesday3sinifid = int.Parse(drwednesday3[2].ToString());
                wednesday3kayitid = int.Parse(drwednesday3[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand wednesday3olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday3olc.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday3olc.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday3olc.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader drwednesday3olc = wednesday3olc.ExecuteReader();
            while (drwednesday3olc.Read())
            {
                if (drwednesday3olc[0].ToString() == "True")
                {
                    btnwednesday3.BackColor = Color.Green;
                }
                else
                {
                    btnwednesday3.BackColor = Color.Orange;
                }
            }
            con.Close();

            //çarşamba 4.saat
            con.Open();
            OleDbCommand wednesday4 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday4.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday4.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday4.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader drwednesday4 = wednesday4.ExecuteReader();
            while (drwednesday4.Read())
            {
                btnwednesday4.Text = drwednesday4[0].ToString() + " " + drwednesday4[3].ToString();
                wednesday4dersid = int.Parse(drwednesday4[1].ToString());
                wednesday4sinifid = int.Parse(drwednesday4[2].ToString());
                wednesday4kayitid = int.Parse(drwednesday4[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand wednesday4olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday4olc.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday4olc.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday4olc.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader drwednesday4olc = wednesday4olc.ExecuteReader();
            while (drwednesday4olc.Read())
            {
                if (drwednesday4olc[0].ToString() == "True")
                {
                    btnwednesday4.BackColor = Color.Green;
                }
                else
                {
                    btnwednesday4.BackColor = Color.Orange;
                }
            }
            con.Close();

            //çarşamba 5.saat
            con.Open();
            OleDbCommand wednesday5 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday5.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday5.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday5.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader drwednesday5 = wednesday5.ExecuteReader();
            while (drwednesday5.Read())
            {
                btnwednesday5.Text = drwednesday5[0].ToString() + " " + drwednesday5[3].ToString();
                wednesday5dersid = int.Parse(drwednesday5[1].ToString());
                wednesday5sinifid = int.Parse(drwednesday5[2].ToString());
                wednesday5kayitid = int.Parse(drwednesday5[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand wednesday5olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday5olc.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday5olc.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday5olc.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader drwednesday5olc = wednesday5olc.ExecuteReader();
            while (drwednesday5olc.Read())
            {
                if (drwednesday5olc[0].ToString() == "True")
                {
                    btnwednesday5.BackColor = Color.Green;
                }
                else
                {
                    btnwednesday5.BackColor = Color.Orange;
                }
            }
            con.Close();

            //çarşamba 6.saat
            con.Open();
            OleDbCommand wednesday6 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday6.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday6.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday6.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader drwednesday6 = wednesday6.ExecuteReader();
            while (drwednesday6.Read())
            {
                btnwednesday6.Text = drwednesday6[0].ToString() + " " + drwednesday6[3].ToString();
                wednesday6dersid = int.Parse(drwednesday6[1].ToString());
                wednesday6sinifid = int.Parse(drwednesday6[2].ToString());
                wednesday6kayitid = int.Parse(drwednesday6[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand wednesday6olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday6olc.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday6olc.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday6olc.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader drwednesday6olc = wednesday6olc.ExecuteReader();
            while (drwednesday6olc.Read())
            {
                if (drwednesday6olc[0].ToString() == "True")
                {
                    btnwednesday6.BackColor = Color.Green;
                }
                else
                {
                    btnwednesday6.BackColor = Color.Orange;
                }
            }
            con.Close();

            //çarşamba 7.saat
            con.Open();
            OleDbCommand wednesday7 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday7.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday7.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday7.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader drwednesday7 = wednesday7.ExecuteReader();
            while (drwednesday7.Read())
            {
                btnwednesday7.Text = drwednesday7[0].ToString() + " " + drwednesday7[3].ToString();
                wednesday7dersid = int.Parse(drwednesday7[1].ToString());
                wednesday7sinifid = int.Parse(drwednesday7[2].ToString());
                wednesday7kayitid = int.Parse(drwednesday7[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand wednesday7olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            wednesday7olc.Parameters.AddWithValue("@P1", "Wednesday");
            wednesday7olc.Parameters.AddWithValue("@P2", ogretmenid);
            wednesday7olc.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader drwednesday7olc = wednesday7olc.ExecuteReader();
            while (drwednesday7olc.Read())
            {
                if (drwednesday7olc[0].ToString() == "True")
                {
                    btnwednesday7.BackColor = Color.Green;
                }
                else
                {
                    btnwednesday7.BackColor = Color.Orange;
                }
            }
            con.Close();

            //perşembe 1.saat
            con.Open();
            OleDbCommand thursday1 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday1.Parameters.AddWithValue("@P1", "Thursday");
            thursday1.Parameters.AddWithValue("@P2", ogretmenid);
            thursday1.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader drthursday1 = thursday1.ExecuteReader();
            while (drthursday1.Read())
            {
                btnthursday1.Text = drthursday1[0].ToString() + " " + drthursday1[3].ToString();
                thursday1dersid = int.Parse(drthursday1[1].ToString());
                thursday1sinifid = int.Parse(drthursday1[2].ToString());
                thursday1kayitid = int.Parse(drthursday1[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand thursday1olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday1olc.Parameters.AddWithValue("@P1", "Thursday");
            thursday1olc.Parameters.AddWithValue("@P2", ogretmenid);
            thursday1olc.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader drthursday1olc = thursday1olc.ExecuteReader();
            while (drthursday1olc.Read())
            {
                if (drthursday1olc[0].ToString() == "True")
                {
                    btnthursday1.BackColor = Color.Green;
                }
                else
                {
                    btnthursday1.BackColor = Color.Orange;
                }
            }
            con.Close();


            //perşembe 2.saat
            con.Open();
            OleDbCommand thursday2 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday2.Parameters.AddWithValue("@P1", "Thursday");
            thursday2.Parameters.AddWithValue("@P2", ogretmenid);
            thursday2.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader drthursday2 = thursday2.ExecuteReader();
            while (drthursday2.Read())
            {
                btnthursday2.Text = drthursday2[0].ToString() + " " + drthursday2[3].ToString();
                thursday2dersid = int.Parse(drthursday2[1].ToString());
                thursday2sinifid = int.Parse(drthursday2[2].ToString());
                thursday2kayitid = int.Parse(drthursday2[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand thursday2olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday2olc.Parameters.AddWithValue("@P1", "Thursday");
            thursday2olc.Parameters.AddWithValue("@P2", ogretmenid);
            thursday2olc.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader drthursday2olc = thursday2olc.ExecuteReader();
            while (drthursday2olc.Read())
            {
                if (drthursday2olc[0].ToString() == "True")
                {
                    btnthursday2.BackColor = Color.Green;
                }
                else
                {
                    btnthursday2.BackColor = Color.Orange;
                }
            }
            con.Close();


            //perşembe 3.saat
            con.Open();
            OleDbCommand thursday3 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday3.Parameters.AddWithValue("@P1", "Thursday");
            thursday3.Parameters.AddWithValue("@P2", ogretmenid);
            thursday3.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader drthursday3 = thursday3.ExecuteReader();
            while (drthursday3.Read())
            {
                btnthursday3.Text = drthursday3[0].ToString() + " " + drthursday3[3].ToString();
                thursday3dersid = int.Parse(drthursday3[1].ToString());
                thursday3sinifid = int.Parse(drthursday3[2].ToString());
                thursday3kayitid = int.Parse(drthursday3[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand thursday3olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday3olc.Parameters.AddWithValue("@P1", "Thursday");
            thursday3olc.Parameters.AddWithValue("@P2", ogretmenid);
            thursday3olc.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader drthursday3olc = thursday3olc.ExecuteReader();
            while (drthursday3olc.Read())
            {
                if (drthursday3olc[0].ToString() == "True")
                {
                    btnthursday3.BackColor = Color.Green;
                }
                else
                {
                    btnthursday3.BackColor = Color.Orange;
                }
            }
            con.Close();

            //perşembe 4.saat
            con.Open();
            OleDbCommand thursday4 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday4.Parameters.AddWithValue("@P1", "Thursday");
            thursday4.Parameters.AddWithValue("@P2", ogretmenid);
            thursday4.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader drthursday4 = thursday4.ExecuteReader();
            while (drthursday4.Read())
            {
                btnthursday4.Text = drthursday4[0].ToString() + " " + drthursday4[3].ToString();
                thursday4dersid = int.Parse(drthursday4[1].ToString());
                thursday4sinifid = int.Parse(drthursday4[2].ToString());
                thursday4kayitid = int.Parse(drthursday4[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand thursday4olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday4olc.Parameters.AddWithValue("@P1", "Thursday");
            thursday4olc.Parameters.AddWithValue("@P2", ogretmenid);
            thursday4olc.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader drthursday4olc = thursday4olc.ExecuteReader();
            while (drthursday4olc.Read())
            {
                if (drthursday4olc[0].ToString() == "True")
                {
                    btnthursday4.BackColor = Color.Green;
                }
                else
                {
                    btnthursday4.BackColor = Color.Orange;
                }
            }
            con.Close();


            //perşembe 5.saat
            con.Open();
            OleDbCommand thursday5 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday5.Parameters.AddWithValue("@P1", "Thursday");
            thursday5.Parameters.AddWithValue("@P2", ogretmenid);
            thursday5.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader drthursday5 = thursday5.ExecuteReader();
            while (drthursday5.Read())
            {
                btnthursday5.Text = drthursday5[0].ToString() + " " + drthursday5[3].ToString();
                thursday5dersid = int.Parse(drthursday5[1].ToString());
                thursday5sinifid = int.Parse(drthursday5[2].ToString());
                thursday5kayitid = int.Parse(drthursday5[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand thursday5olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday5olc.Parameters.AddWithValue("@P1", "Thursday");
            thursday5olc.Parameters.AddWithValue("@P2", ogretmenid);
            thursday5olc.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader drthursday5olc = thursday5olc.ExecuteReader();
            while (drthursday5olc.Read())
            {
                if (drthursday5olc[0].ToString() == "True")
                {
                    btnthursday5.BackColor = Color.Green;
                }
                else
                {
                    btnthursday5.BackColor = Color.Orange;
                }
            }
            con.Close();

            //perşembe 6.saat
            con.Open();
            OleDbCommand thursday6 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday6.Parameters.AddWithValue("@P1", "Thursday");
            thursday6.Parameters.AddWithValue("@P2", ogretmenid);
            thursday6.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader drthursday6 = thursday6.ExecuteReader();
            while (drthursday6.Read())
            {
                btnthursday6.Text = drthursday6[0].ToString() + " " + drthursday6[3].ToString();
                thursday6dersid = int.Parse(drthursday6[1].ToString());
                thursday6sinifid = int.Parse(drthursday6[2].ToString());
                thursday6kayitid = int.Parse(drthursday6[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand thursday6olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday6olc.Parameters.AddWithValue("@P1", "Thursday");
            thursday6olc.Parameters.AddWithValue("@P2", ogretmenid);
            thursday6olc.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader drthursday6olc = thursday6olc.ExecuteReader();
            while (drthursday6olc.Read())
            {
                if (drthursday6olc[0].ToString() == "True")
                {
                    btnthursday6.BackColor = Color.Green;
                }
                else
                {
                    btnthursday6.BackColor = Color.Orange;
                }
            }
            con.Close();

            //perşembe 7.saat
            con.Open();
            OleDbCommand thursday7 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday7.Parameters.AddWithValue("@P1", "Thursday");
            thursday7.Parameters.AddWithValue("@P2", ogretmenid);
            thursday7.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader drthursday7 = thursday7.ExecuteReader();
            while (drthursday7.Read())
            {
                btnthursday7.Text = drthursday7[0].ToString() + " " + drthursday7[3].ToString();
                thursday7dersid = int.Parse(drthursday7[1].ToString());
                thursday7sinifid = int.Parse(drthursday7[2].ToString());
                thursday7kayitid = int.Parse(drthursday7[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand thursday7olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            thursday7olc.Parameters.AddWithValue("@P1", "Thursday");
            thursday7olc.Parameters.AddWithValue("@P2", ogretmenid);
            thursday7olc.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader drthursday7olc = thursday7olc.ExecuteReader();
            while (drthursday7olc.Read())
            {
                if (drthursday7olc[0].ToString() == "True")
                {
                    btnthursday7.BackColor = Color.Green;
                }
                else
                {
                    btnthursday7.BackColor = Color.Orange;
                }
            }
            con.Close();

            //cuma 1.saat
            con.Open();
            OleDbCommand friday1 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday1.Parameters.AddWithValue("@P1", "Friday");
            friday1.Parameters.AddWithValue("@P2", ogretmenid);
            friday1.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader drfriday1 = friday1.ExecuteReader();
            while (drfriday1.Read())
            {
                btnfriday1.Text = drfriday1[0].ToString() + " " + drfriday1[3].ToString();
                friday1dersid = int.Parse(drfriday1[1].ToString());
                friday1sinifid = int.Parse(drfriday1[2].ToString());
                friday1kayitid = int.Parse(drfriday1[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand friday1olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday1olc.Parameters.AddWithValue("@P1", "Friday");
            friday1olc.Parameters.AddWithValue("@P2", ogretmenid);
            friday1olc.Parameters.AddWithValue("@P3", "1");
            OleDbDataReader drfriday1olc = friday1olc.ExecuteReader();
            while (drfriday1olc.Read())
            {
                if (drfriday1olc[0].ToString() == "True")
                {
                    btnfriday1.BackColor = Color.Green;
                }
                else
                {
                    btnfriday1.BackColor = Color.Orange;
                }
            }
            con.Close();

            //cuma 2.saat
            con.Open();
            OleDbCommand friday2 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday2.Parameters.AddWithValue("@P1", "Friday");
            friday2.Parameters.AddWithValue("@P2", ogretmenid);
            friday2.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader drfriday2 = friday2.ExecuteReader();
            while (drfriday2.Read())
            {
                btnfriday2.Text = drfriday2[0].ToString() + " " + drfriday2[3].ToString();
                friday2dersid = int.Parse(drfriday2[1].ToString());
                friday2sinifid = int.Parse(drfriday2[2].ToString());
                friday2kayitid = int.Parse(drfriday2[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand friday2olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday2olc.Parameters.AddWithValue("@P1", "Friday");
            friday2olc.Parameters.AddWithValue("@P2", ogretmenid);
            friday2olc.Parameters.AddWithValue("@P3", "2");
            OleDbDataReader drfriday2olc = friday2olc.ExecuteReader();
            while (drfriday2olc.Read())
            {
                if (drfriday2olc[0].ToString() == "True")
                {
                    btnfriday2.BackColor = Color.Green;
                }
                else
                {
                    btnfriday2.BackColor = Color.Orange;
                }
            }
            con.Close();

            //cuma 3.saat
            con.Open();
            OleDbCommand friday3 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday3.Parameters.AddWithValue("@P1", "Friday");
            friday3.Parameters.AddWithValue("@P2", ogretmenid);
            friday3.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader drfriday3 = friday3.ExecuteReader();
            while (drfriday3.Read())
            {
                btnfriday3.Text = drfriday3[0].ToString() + " " + drfriday3[3].ToString();
                friday3dersid = int.Parse(drfriday3[1].ToString());
                friday3sinifid = int.Parse(drfriday3[2].ToString());
                friday3kayitid = int.Parse(drfriday3[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand friday3olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday3olc.Parameters.AddWithValue("@P1", "Friday");
            friday3olc.Parameters.AddWithValue("@P2", ogretmenid);
            friday3olc.Parameters.AddWithValue("@P3", "3");
            OleDbDataReader drfriday3olc = friday3olc.ExecuteReader();
            while (drfriday3olc.Read())
            {
                if (drfriday3olc[0].ToString() == "True")
                {
                    btnfriday3.BackColor = Color.Green;
                }
                else
                {
                    btnfriday3.BackColor = Color.Orange;
                }
            }
            con.Close();

            //cuma 4.saat
            con.Open();
            OleDbCommand friday4 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday4.Parameters.AddWithValue("@P1", "Friday");
            friday4.Parameters.AddWithValue("@P2", ogretmenid);
            friday4.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader drfriday4 = friday4.ExecuteReader();
            while (drfriday4.Read())
            {
                btnfriday4.Text = drfriday4[0].ToString() + " " + drfriday4[3].ToString();
                friday4dersid = int.Parse(drfriday4[1].ToString());
                friday4sinifid = int.Parse(drfriday4[2].ToString());
                friday4kayitid = int.Parse(drfriday4[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand friday4olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday4olc.Parameters.AddWithValue("@P1", "Friday");
            friday4olc.Parameters.AddWithValue("@P2", ogretmenid);
            friday4olc.Parameters.AddWithValue("@P3", "4");
            OleDbDataReader drfriday4olc = friday4olc.ExecuteReader();
            while (drfriday4olc.Read())
            {
                if (drfriday4olc[0].ToString() == "True")
                {
                    btnfriday4.BackColor = Color.Green;
                }
                else
                {
                    btnfriday4.BackColor = Color.Orange;
                }
            }
            con.Close();

            //cuma 5.saat
            con.Open();
            OleDbCommand friday5 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday5.Parameters.AddWithValue("@P1", "Friday");
            friday5.Parameters.AddWithValue("@P2", ogretmenid);
            friday5.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader drfriday5 = friday5.ExecuteReader();
            while (drfriday5.Read())
            {
                btnfriday5.Text = drfriday5[0].ToString() + " " + drfriday5[3].ToString();
                friday5dersid = int.Parse(drfriday5[1].ToString());
                friday5sinifid = int.Parse(drfriday5[2].ToString());
                friday5kayitid = int.Parse(drfriday5[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand friday5olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday5olc.Parameters.AddWithValue("@P1", "Friday");
            friday5olc.Parameters.AddWithValue("@P2", ogretmenid);
            friday5olc.Parameters.AddWithValue("@P3", "5");
            OleDbDataReader drfriday5olc = friday5olc.ExecuteReader();
            while (drfriday5olc.Read())
            {
                if (drfriday5olc[0].ToString() == "True")
                {
                    btnfriday5.BackColor = Color.Green;
                }
                else
                {
                    btnfriday5.BackColor = Color.Orange;
                }
            }
            con.Close();

            //cuma 6.saat
            con.Open();
            OleDbCommand friday6 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday6.Parameters.AddWithValue("@P1", "Friday");
            friday6.Parameters.AddWithValue("@P2", ogretmenid);
            friday6.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader drfriday6 = friday6.ExecuteReader();
            while (drfriday6.Read())
            {
                btnfriday6.Text = drfriday6[0].ToString() + " " + drfriday6[3].ToString();
                friday6dersid = int.Parse(drfriday6[1].ToString());
                friday6sinifid = int.Parse(drfriday6[2].ToString());
                friday6kayitid = int.Parse(drfriday6[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand friday6olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday6olc.Parameters.AddWithValue("@P1", "Friday");
            friday6olc.Parameters.AddWithValue("@P2", ogretmenid);
            friday6olc.Parameters.AddWithValue("@P3", "6");
            OleDbDataReader drfriday6olc = friday6olc.ExecuteReader();
            while (drfriday6olc.Read())
            {
                if (drfriday6olc[0].ToString() == "True")
                {
                    btnfriday6.BackColor = Color.Green;
                }
                else
                {
                    btnfriday6.BackColor = Color.Orange;
                }
            }
            con.Close();

            //cuma 7.saat
            con.Open();
            OleDbCommand friday7 = new OleDbCommand("select BRANSADI, DERS, SINIF,SINIFAD, KAYITID FROM (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday7.Parameters.AddWithValue("@P1", "Friday");
            friday7.Parameters.AddWithValue("@P2", ogretmenid);
            friday7.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader drfriday7 = friday7.ExecuteReader();
            while (drfriday7.Read())
            {
                btnfriday7.Text = drfriday7[0].ToString() + " " + drfriday7[3].ToString();
                friday7dersid = int.Parse(drfriday7[1].ToString());
                friday7sinifid = int.Parse(drfriday7[2].ToString());
                friday7kayitid = int.Parse(drfriday7[4].ToString());
            }

            con.Close();
            con.Open();
            OleDbCommand friday7olc = new OleDbCommand("select OLCDURUM FROM TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID WHERE  TARIH=@P1 AND OGRETMEN=@P2 AND DERSSAATI=@P3", con);
            friday7olc.Parameters.AddWithValue("@P1", "Friday");
            friday7olc.Parameters.AddWithValue("@P2", ogretmenid);
            friday7olc.Parameters.AddWithValue("@P3", "7");
            OleDbDataReader drfriday7olc = friday7olc.ExecuteReader();
            while (drfriday7olc.Read())
            {
                if (drfriday7olc[0].ToString() == "True")
                {
                    btnfriday7.BackColor = Color.Green;
                }
                else
                {
                    btnfriday7.BackColor = Color.Orange;
                }
            }
            con.Close();

            con.Open();
            OleDbCommand adbransoku = new OleDbCommand("select OGRETMENID,ADISOYADI, BRANS, BRANSI FROM TBLOGRETMENLER where TCKIMLIKNO=@O1", con);
            adbransoku.Parameters.AddWithValue("@O1", sifrele(ogretmentc));
            OleDbDataReader rd = adbransoku.ExecuteReader();
            while (rd.Read())
            {

                lbladbrans.Text = ("SAYIN " + rd[1] + " HOŞGELDİNİZ.").ToUpper();
                

            }
            con.Close();
            lblbilgilendirme.Text="-->Puanlama şeklini seçtikten sonra puanlama yapmak istediğiniz derse tıklayınız.";
            lblbilgi2.Text = "-->Yeşil renkli butonlar puanlanmış, turuncu renkli butonlar puanlanmamış dersleri gösterir.";
            lblbilgi3.Text= "-->Günlük puanlama sadece seçilen dersi, haftalık puanlama hafta içinde sınıfın seçilen dersteki puanlanmamış bütün derslerini puanlar.";
            timer1.Start();
            //toplam ders sayısı

            con.Open();
            OleDbCommand toplamders=new OleDbCommand("select count(*) from TBLDERSPROGRAMI where OGRETMEN=@P1",con);
            toplamders.Parameters.AddWithValue("@P1", ogretmenid);
            OleDbDataReader drtoplamders=toplamders.ExecuteReader();
            while(drtoplamders.Read())
            {
                progressBar1.Maximum = int.Parse(drtoplamders[0].ToString());
                toplamderssayisi = int.Parse(drtoplamders[0].ToString());


            }
            con.Close();

            //puanlanmış ders sayısı
            con.Open();
            OleDbCommand puanlanmisders = new OleDbCommand("select count(*) from TBLDERSPROGRAMI WHERE OLCDURUM=@P1 AND OGRETMEN=@P2", con);
            puanlanmisders.Parameters.AddWithValue("@P1", true);
            puanlanmisders.Parameters.AddWithValue("@P2", ogretmenid);
            OleDbDataReader drpuanlanmisders = puanlanmisders.ExecuteReader();
            while (drpuanlanmisders.Read())
            {
                progressBar1.Value = int.Parse(drpuanlanmisders[0].ToString());
                puanlanmisderssayisi = int.Parse(drpuanlanmisders[0].ToString());
            }
            con.Close();
            if (toplamderssayisi != 0)
            {
                yuzdelik = (100 / toplamderssayisi) * puanlanmisderssayisi;
                lblpuanlamayuzde.Text = ("%" + yuzdelik.ToString() + " oranında puanlama yapılmıştır.");
                lblpuanlanmisderssy.Text= (puanlanmisderssayisi.ToString() + " ders puanlanmıştır.");
            }

            //puanlanmamış ders listesi
            OleDbDataAdapter da = new OleDbDataAdapter("select BRANSADI AS 'DERS ADI', SINIFAD AS 'SINIF ADI', DERSSAATI AS 'DERS SAATİ' from (TBLDERSPROGRAMI INNER JOIN TBLBRANSLAR ON TBLDERSPROGRAMI.DERS=TBLBRANSLAR.BRANSID) INNER JOIN TBLSINIFLAR ON TBLSINIFLAR.SINIFID=TBLDERSPROGRAMI.SINIF WHERE OLCDURUM=@P1 AND OGRETMEN=@P2", con);
            da.SelectCommand.Parameters.AddWithValue("@P1", false);
            da.SelectCommand.Parameters.AddWithValue("@P2", ogretmenid);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].HeaderText="PUANLANMAMIŞ DERSLER";


        }
        int haftaici;
        public int haftaal(DateTime dtPassed)
        {

            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dtPassed, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;

        }
        private void frmdersprogrami_Load(object sender, EventArgs e)
        {
            int hafta = haftaal(DateTime.Now);
            OleDbConnection con = new OleDbConnection(conn.baglan);
            con.Open();
            OleDbCommand komuthaftaoku = new OleDbCommand("select HAFTA FROM TBLDERSPROGRAMI", con);
            OleDbDataReader komuthaftaokulrd = komuthaftaoku.ExecuteReader();
            while (komuthaftaokulrd.Read())
            {
                haftaici = int.Parse(komuthaftaokulrd[0].ToString());
            }
            con.Close();

            con.Open();
            if (haftaici != hafta)
            {
                OleDbCommand komuthaftayidegistir = new OleDbCommand("update TBLDERSPROGRAMI SET HAFTA=@P1 , OLCDURUM=@P2", con);
                komuthaftayidegistir.Parameters.AddWithValue("@P1", hafta);
                komuthaftayidegistir.Parameters.AddWithValue("@P2", 0);
                komuthaftayidegistir.ExecuteNonQuery();
            }
            con.Close();
            dersprogramigetir();

        }
        public string rolum;
        private void pictureBox5_Click(object sender, EventArgs e)
        {
            if (rolum == "yonetici")
            {
                this.Close();

            }
            else
            {

                Application.Exit();
            }
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            pckhand.Left += 1;
            if (pckhand.Left > 10)
            {
                pckhand.Left = 0;
            }
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            frmistatistikler frmistatistikler = new frmistatistikler();
            frmistatistikler.ogretmennum=ogretmenid;
            frmistatistikler.rol="ogretmen";
            frmistatistikler.Show();

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            frmayarlar frmayarlar = new frmayarlar();
            frmayarlar.ogretmenayar = ogretmenid;
            frmayarlar.Show(Owner);

        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.WindowState= FormWindowState.Minimized;
        }

        private void btnmonday1_Click(object sender, EventArgs e)
        {
            
            
            if(btnmonday1.BackColor==Color.Orange)
            {
                if(rdygunluk.Checked==true||rdyhaftalik.Checked==true)
                {
                    
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Pazartesi gününün 1. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.bransid = monday1dersid;
                            frmanaform.sinif = monday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if(rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnmonday1.BackColor == Color.Green) 
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void btnmonday2_Click(object sender, EventArgs e)
        {
            
           
            if (btnmonday2.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Pazartesi gününün 2. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnmonday2.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void btnmonday3_Click(object sender, EventArgs e)
        {
            if (btnmonday3.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Pazartesi gününün 3. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnmonday3.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        private void btnmonday4_Click(object sender, EventArgs e)
        {
            if (btnmonday4.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Pazartesi gününün 4. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday4dersid; 
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnmonday4.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnmonday5_Click(object sender, EventArgs e)
        {
            if (btnmonday5.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Pazartesi gününün 5. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnmonday5.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnmonday6_Click(object sender, EventArgs e)
        {
            if(btnmonday6.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Pazartesi gününün 6. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnmonday6.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnmonday7_Click(object sender, EventArgs e)
        {
            if (btnmonday7.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Pazartesi gününün 7. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = monday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = monday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Monday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnmonday7.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btntuesday1_Click(object sender, EventArgs e)
        {
            if (btntuesday1.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Salı gününün 1. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btntuesday1.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btntuesday2_Click(object sender, EventArgs e)
        {
            if (btntuesday2.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Salı gününün 2. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btntuesday2.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btntuesday3_Click(object sender, EventArgs e)
        {
            if (btntuesday3.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Salı gününün 3. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btntuesday3.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btntuesday4_Click(object sender, EventArgs e)
        {
            if (btntuesday4.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Salı gününün 4. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btntuesday4.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btntuesday5_Click(object sender, EventArgs e)
        {
            if (btntuesday5.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Salı gününün 5. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btntuesday5.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void btntuesday6_Click(object sender, EventArgs e)
        {
            if (btntuesday6.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Salı gününün 6. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btntuesday6.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btntuesday7_Click(object sender, EventArgs e)
        {
            if (btntuesday7.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Salı gününün 7. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = tuesday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = tuesday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Tuesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btntuesday7.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnwednesday1_Click(object sender, EventArgs e)
        {
            if (btnwednesday1.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Çarşamba gününün 1. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnwednesday1.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnwednesday2_Click(object sender, EventArgs e)
        {
            if (btnwednesday2.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Çarşamba gününün 2. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnwednesday2.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnwednesday3_Click(object sender, EventArgs e)
        {
            if (btnwednesday3.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Çarşamba gününün 3. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnwednesday3.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnwednesday4_Click(object sender, EventArgs e)
        {
            if (btnwednesday4.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Çarşamba gününün 4. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnwednesday4.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnwednesday5_Click(object sender, EventArgs e)
        {
            if (btnwednesday5.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Çarşamba gününün 5. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnwednesday5.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnwednesday6_Click(object sender, EventArgs e)
        {
            if (btnwednesday6.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Çarşamba gününün 6. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnwednesday6.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnwednesday7_Click(object sender, EventArgs e)
        {
            if (btnwednesday7.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Çarşamba gününün 7. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = wednesday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = wednesday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Wednesday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnwednesday7.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnthursday1_Click(object sender, EventArgs e)
        {
            if (btnthursday1.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Perşembe gününün 1. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnthursday1.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnthursday2_Click(object sender, EventArgs e)
        {
            if (btnthursday2.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Perşembe gününün 2. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnthursday2.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnthursday3_Click(object sender, EventArgs e)
        {
            if (btnthursday3.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Perşembe gününün 3. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnthursday3.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnthursday4_Click(object sender, EventArgs e)
        {
            if (btnthursday4.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Perşembe gününün 4. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnthursday4.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnthursday5_Click(object sender, EventArgs e)
        {
            if (btnthursday5.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Perşembe gününün 5. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnthursday5.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnthursday6_Click(object sender, EventArgs e)
        {

            if (btnthursday6.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Perşembe gününün 6. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnthursday6.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnthursday7_Click(object sender, EventArgs e)
        {
            if (btnthursday7.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Perşembe gününün 7. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = thursday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = thursday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Thursday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnthursday7.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnfriday1_Click(object sender, EventArgs e)
        {
            if (btnfriday1.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Cuma gününün 1. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday1dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday1sinifid;
                            frmanaform.derssaati = 1;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnfriday1.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnfriday2_Click(object sender, EventArgs e)
        {
            if (btnfriday2.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Cuma gününün 2. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday2dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday2sinifid;
                            frmanaform.derssaati = 2;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnfriday2.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnfriday3_Click(object sender, EventArgs e)
        {
            if (btnfriday3.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Cuma gününün 3. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday3dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday3sinifid;
                            frmanaform.derssaati = 3;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnfriday3.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnfriday4_Click(object sender, EventArgs e)
        {
            if (btnfriday4.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Cuma gününün 4. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday4dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday4sinifid;
                            frmanaform.derssaati = 4;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnfriday4.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnfriday5_Click(object sender, EventArgs e)
        {
            if (btnfriday5.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Cuma gününün 5. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday5dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday5sinifid;
                            frmanaform.derssaati = 5;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnfriday5.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnfriday6_Click(object sender, EventArgs e)
        {
            if (btnfriday6.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Cuma gününün 6. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday6dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday6sinifid;
                            frmanaform.derssaati = 6;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnfriday6.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnfriday7_Click(object sender, EventArgs e)
        {
            if (btnfriday7.BackColor == Color.Orange)
            {
                if (rdygunluk.Checked == true || rdyhaftalik.Checked == true)
                {
                    if (rdygunluk.Checked == true)
                    {
                        DialogResult result = MessageBox.Show("Cuma gününün 7. ders saatini puanlamak için puanlama sayfasına gidilecek? Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();

                        }
                    }
                    else if (rdyhaftalik.Checked == true)
                    {
                        DialogResult resulthaftalik = MessageBox.Show("Seçilen sınıfın seçilen dersteki puanlanmamış bütün ders saatlerini puanlamak için puanlama sayfasına gidilecek. Onaylıyor musunuz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resulthaftalik == DialogResult.Yes)
                        {
                            frmanaform frmanaform = new frmanaform();
                            frmanaform.ogretmenid = ogretmenid;
                            frmanaform.bransid = friday7dersid;
                            frmanaform.rol = "ogretmen";
                            frmanaform.tc = ogretmentc;
                            frmanaform.sinif = friday7sinifid;
                            frmanaform.derssaati = 7;
                            frmanaform.gun = "Friday";
                            if (rdygunluk.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 1;
                            }
                            else if (rdyhaftalik.Checked == true)
                            {
                                frmanaform.gunlukhaftalik = 0;
                            }
                            frmanaform.Show();
                            this.Close();
                        }
                    }


                }
                else
                {
                    MessageBox.Show("Lütfen puanlama şeklini seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (btnfriday7.BackColor == Color.Green)
            {
                MessageBox.Show("Bu ders zaten puanlanmış.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

       
    }
}
