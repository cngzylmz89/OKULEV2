using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace DILIMIZINZENGINLIKLERI
{
    public partial class Form1 : Form
    {
        // Kullanacağımız timer
        private Timer myTimer = new Timer();
        // Foreground kontrolü için yardımcı timer
        private Timer checkTimer = new Timer();
        // Timer durdurulmuş mu bilgisi
        private bool timerPaused = false;
        public Timer yazTimer = new Timer();
        private int index = 0;

       public  string sozcuk, sozcukanlam, sozcukcumle, deyim, deyimanlam, deyimcumle, atasozu, atasozuyazar, sozcukcumlekoyu, deyimcumlekoyu;
        string timer;
        
        
       baglantisinif con=new baglantisinif();
        public Form1()
        {
            InitializeComponent();    
        }

       
     

        //OleDbConnection excel = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Combine(Application.StartupPath, "KELIMELER.xlsx") + ";Extended Properties='Excel 12.0 Xml;HDR=YES;'");
      
      public  DateTime tarih = DateTime.Now;
       public  int sayi;
        int saniye = 300;
        SoundPlayer daktilo = new SoundPlayer(Application.StartupPath + @"\ses.wav");
       
        SoundPlayer giris = new SoundPlayer(Application.StartupPath + @"\giris.wav");
      


        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Bu uygulama DİLİMİZİN ZENGİNLİKLERİ PROJESİ kapsamında Sultangazi 125. Yıl Ortaokulu Türkçe Öğretmeni Cengiz YILMAZ tarafından hazırlanmıştır. Bilgi için muallimiturki@gmail.com adresine ileti gönderebilirsiniz.", "Bilgi");
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
           OleDbConnection excel=new OleDbConnection(con.baglan);
            saniye--;
            if (saniye < 60)
            {
                if (sayi == 0)
                {
                    saniyeler.Text = "Kapatılmasına " + saniye.ToString() + " saniye kaldı.";
                    pictureBox3.Visible = true;
                   
                   
                }
               
               
            }
            else
            {
                int dakika = saniye / 60;
                int saniye2 = saniye % 60;
                if (sayi==0)
                {
                    saniyeler.Text = "Kapatılmasına " + dakika.ToString() + " dakika " + saniye2.ToString() + " saniye kaldı.";
                }
               

            }

            if (sayi==0&&saniye < 2)
            {
                try
                {
                    excel.Open();
                    OleDbCommand cmd2 = new OleDbCommand("update [KELIMELER$] set GOSTER=@p1 where TARIH=@p2", excel);
                    cmd2.Parameters.AddWithValue("@p1", 1);
                    cmd2.Parameters.AddWithValue("@p2", tarih.ToString("dd.MM.yyyy"));
                    cmd2.ExecuteNonQuery();

                    excel.Close();
                }
                catch (Exception hata)
                {

                    MessageBox.Show(hata.Message, "Hata");
                }
                
            }
            
            if (saniye <= 0)
            {
                Application.Exit();
            }

        }

      

       
        


        private void SetColoredText(RichTextBox rtb, string text, string coloredPart)
        {
            rtb.Clear();
            int start = text.IndexOf(coloredPart);

            rtb.AppendText(text);

            if (start >= 0)
            {
                rtb.Select(start, coloredPart.Length);
                rtb.SelectionColor = Color.Red; // istediğin renk
                rtb.SelectionLength = 0;
            }
        }
        

        private void timer2_Tick(object sender, EventArgs e)
        {
            
            int indexsozcuk = lblsozcuk.Text.IndexOf(sozcuk);
            if (index < sozcuk.Length)
            {
                

                lblsozcuk.Text += sozcuk[index];



                //if (sayi == 0)
                //{
                //    daktilo.Play();
                //}
               
                
                index++;
            }
            else
            {
                timer2.Stop();
              
                index = 0;
                timer = "sozcukanlam";
                tmrsozcukanlam.Start();
               
            }

            
        }
      
        
        

        
        private void Form1_Shown(object sender, EventArgs e)
        {
            if (sayi == 0)
            {
                giris.Play();
            }
           
            
            

            label1.Text = DateTime.Now.ToString("dd.MM.yyyy");


            if (sayi == 0)
            {
                timer1.Start();
                
                timer2.Start(); // form açılınca yazı başlasın
            }
            else if (sayi == 1)
            {
                timer2.Interval = 100;
                tmrsozcukanlam.Interval = 100;
                tmrsozcukcumle.Interval = 100;
                tmrdeyim.Interval =100;
                tmrdeyimanlam.Interval = 100;
                tmrdeyimcumle.Interval =100;
                tmratasozu.Interval = 100;
                tmryazar.Interval = 100;

                timer1.Start();
               
                timer2.Start();
            }
          
        }
        private void tmrsozcukanlam_Tick(object sender, EventArgs e)
        {
            if (index < sozcukanlam.Length)
            {
              
                lblsozcukanlam.Text += sozcukanlam[index];
              
                index++;
            }
            else
            {
                tmrsozcukanlam.Stop();
                index = 0;
                timer = "sozcukcumle";
                
                tmrsozcukcumle.Start();
                
            }
        }

        private void tmrsozcukcumle_Tick(object sender, EventArgs e)
        {
            if (index < sozcukcumle.Length)
            {
                lblsozcukcumle.Text += sozcukcumle[index];

             

                index++;
            }
            else
            {
                tmrsozcukcumle.Stop();
                index = 0;
                timer = "deyim";
                SetColoredText(lblsozcukcumle, sozcukcumle, sozcukcumlekoyu);
               
                tmrdeyim.Start();


            }
        }

        private void tmrdeyim_Tick(object sender, EventArgs e)
        {
            if (index < deyim.Length)
            {
                lbldeyim.Text += deyim[index];

               

                index++;
            }
            else
            {
                tmrdeyim.Stop();
                index = 0;
                timer = "deyimanlam";
                tmrdeyimanlam.Start();

            }
        }

        private void tmrdeyimanlam_Tick(object sender, EventArgs e)
        {
            if (index < deyimanlam.Length)
            {
                lbldeyimanlam.Text += deyimanlam[index];

               

                index++;
            }
            else
            {
                tmrdeyimanlam.Stop();
                index = 0;
                timer = "deyimcumle";
                tmrdeyimcumle.Start();

            }
        }

        private void tmrdeyimcumle_Tick(object sender, EventArgs e)
        {
            if (index < deyimcumle.Length)
            {
                lbldeyimcumle.Text += deyimcumle[index];

              

                index++;
            }
            else
            {
                tmrdeyimcumle.Stop();
                index = 0;
                timer = "atasozu";
                SetColoredText(lbldeyimcumle, deyimcumle, deyimcumlekoyu);
                tmratasozu.Start();

            }
        }

        private void tmratasozu_Tick(object sender, EventArgs e)
        {
            if (index < atasozu.Length)
            {
                lblatasozsoz.Text += atasozu[index];

               

                index++;
            }
            else
            {
                tmratasozu.Stop();
                index = 0;
                timer = "atasozuyazar";
                tmryazar.Start();


            }
        }

        private void tmryazar_Tick(object sender, EventArgs e)
        {
            if (index < atasozuyazar.Length)
            {
                lblatasozuyazar.Text+= atasozuyazar[index];

               

                index++;
               
            }
            else
            {
               
                tmryazar.Stop();
                index = 0;
                timer = "bitti";
                


                           if (sayi == 1&&timer=="bitti")
                {
                    Application.Exit();
                }
               


            }
        }
    }
}
