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
<<<<<<< HEAD
using System.Media;
using Microsoft.Win32;
using System.Speech.Synthesis;
=======
>>>>>>> 1e3164c1f27f86f5c311cb1d2252f9e69f862450


namespace DILIMIZINZENGINLIKLERI
{
    public partial class DilimizinZenginlikleri : Form
    {
        SpeechSynthesizer ses = new SpeechSynthesizer();
        // Kullanacağımız timer
        private Timer myTimer = new Timer();
        // Foreground kontrolü için yardımcı timer
        private Timer checkTimer = new Timer();
        // Timer durdurulmuş mu bilgisi
        private bool timerPaused = false;
        public Timer yazTimer = new Timer();
        private int index = 0;

        string sozcuk, sozcukanlam, sozcukcumle, deyim, deyimanlam, deyimcumle, atasozu, atasozuyazar;
        string timer;

        private void AddToStartup()
        {
            string exePath = Application.ExecutablePath;
            RegistryKey key = Registry.CurrentUser.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
            key.SetValue("BenimProgramim", exePath);
        }

        public DilimizinZenginlikleri()
        {

            InitializeComponent();
<<<<<<< HEAD
            // Kenarlıkları kaldır
            this.FormBorderStyle = FormBorderStyle.None;
            // Normal moda al (Maximized yerine)
            this.WindowState = FormWindowState.Normal;
            // Ekranın tamamını kapla (görev çubuğu dahil)
            this.Bounds = Screen.PrimaryScreen.Bounds;
            // Üstte kalmasını istiyorsan
            this.TopMost = true;

           

           


            var currentScreen = Screen.FromControl(this);
            this.Bounds = currentScreen.Bounds;

            this.AutoScaleMode = AutoScaleMode.Dpi;
=======
            AddToStartup();
>>>>>>> 1e3164c1f27f86f5c311cb1d2252f9e69f862450
            checkTimer.Interval = 300; // 0.3 saniyede bir kontrol et
            checkTimer.Tick += CheckTimer_Tick;
          
        }

<<<<<<< HEAD
       
=======
        private void AddToStartup()
        {
            try
            {
                string appName = "Dilimizinzenginlikleri"; // Başlangıçta görünecek ad
                string exePath = Application.ExecutablePath;

                RegistryKey key = Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);

                // Zaten ekli değilse ekle
                if (key.GetValue(appName) == null)
                {
                    key.SetValue(appName, exePath);
                   
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

>>>>>>> 1e3164c1f27f86f5c311cb1d2252f9e69f862450
        private void CheckTimer_Tick(object sender, EventArgs e)
        {
            IntPtr fg = GetForegroundWindow();

            if (fg != this.Handle)
                PauseTimer();
            else
                ResumeTimer();
        }

        OleDbConnection excel = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Combine(Application.StartupPath, "KELIMELER.xlsx") + ";Extended Properties='Excel 12.0 Xml;HDR=YES;'");
        DateTime tarih = DateTime.Now;
        int sayi;
        int saniye = 300;
        SoundPlayer daktilo = new SoundPlayer(Application.StartupPath + @"\ses.wav");
<<<<<<< HEAD
        
        private void Form1_Load(object sender, EventArgs e)
        {


            AddToStartup();
            timer = "sozcuk";
=======
        async Task getir()
        {
            this.SuspendLayout();
            
            this.ResumeLayout();
            this.WindowState = FormWindowState.Maximized;
>>>>>>> 1e3164c1f27f86f5c311cb1d2252f9e69f862450
            label1.Text = DateTime.Now.ToString("dd.MM.yyyy");
            await Task.Run(() =>
            {
                System.Threading.Thread.Sleep(1000);
            
            excel.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [KELIMELER$] WHERE TARIH=@p1", excel);
            cmd.Parameters.AddWithValue("@p1", tarih.ToString("dd.MM.yyyy"));
            OleDbDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                sayi = int.Parse(dr[10].ToString());
                if (sayi == 0)
                {

                        this.Invoke(new Action(() => {
                            sozcuk = dr[1].ToString();
                            sozcukanlam = dr[2].ToString();
                            sozcukcumle = dr[3].ToString();
                            deyim = dr[4].ToString();
                            deyimanlam = dr[5].ToString();
                            deyimcumle = dr[6].ToString();
                            atasozu = dr[7].ToString();
                            atasozuyazar = dr[8].ToString();
                            
                            
                        }));
                       

                    

                }
                else if (sayi == 1)
                {
                        this.Invoke(new Action(() => {
                           sozcuk = dr[1].ToString();
                    sozcukanlam = dr[2].ToString() ;
                    sozcukcumle = dr[3].ToString();
                    deyim = dr[4].ToString();
                    deyimanlam = dr[5].ToString();
                    deyimcumle = dr[6].ToString();
                    atasozu = dr[7].ToString();
                    atasozuyazar = dr[8].ToString();
                        }));
                       
                    


                }

            }
            excel.Close();
            });

            if (sayi == 0)
            {
                timer1.Start();
                checkTimer.Start();
                timer2.Start(); // form açılınca yazı başlasın
            }
            else if (sayi == 1)
            {
                timer2.Interval = 150;
                tmrsozcukanlam.Interval = 150;
                tmrsozcukcumle.Interval = 150;
                tmrdeyim.Interval = 150;
                tmrdeyimanlam.Interval = 150;
                tmrdeyimcumle.Interval = 150;
                tmratasozu.Interval = 150;
                tmryazar.Interval = 150;

                timer1.Start();
                checkTimer.Start();
                timer2.Start();
            }
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.CenterScreen;
        }
        

        

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Bu uygulama DİLİMİZİN ZENGİNLİKLERİ PROJESİ kapsamında Sultangazi 125. Yıl Ortaokulu Türkçe Öğretmeni Cengiz YILMAZ tarafından hazırlanmıştır. Bilgi için muallimiturki@gmail.com adresine ileti gönderebilirsiniz.", "Bilgi");
        }

       

        private void btncikis_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Application.Exit();
        }


       
<<<<<<< HEAD
=======

>>>>>>> 1e3164c1f27f86f5c311cb1d2252f9e69f862450
        private void timer1_Tick(object sender, EventArgs e)
        {
            IntPtr
                ıntPtr1 = GetForegroundWindow();
            IntPtr ıntPtr = ıntPtr1;
            IntPtr fg = ıntPtr;

            if (fg != this.Handle)
                PauseTimer();
            else
                ResumeTimer();
            saniye--;
            if (saniye < 60)
            {
                if (sayi == 0)
                {
                    saniyeler.Text = "Kapatılmasına " + saniye.ToString() + " saniye kaldı.";
                }
               
                pictureBox3.Visible = true;
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
            if (saniye == 0)
            {
                excel.Open();
                OleDbCommand cmd2 = new OleDbCommand("update [KELIMELER$] set GOSTER=@p1 where TARIH=@p2", excel);
                cmd2.Parameters.AddWithValue("@p1", 1);
                cmd2.Parameters.AddWithValue("@p2", tarih.ToString("dd.MM.yyyy"));
                cmd2.ExecuteNonQuery();
                excel.Close();
                Application.Exit();
            }
        }

      

        private void PauseTimer()
        {
            if (!timerPaused)
            {
                timer1.Stop();
                if (timer=="sozcuk")
                {
                    timer2.Stop();
                }
                if (timer=="sozcukanlam")
                {
                    tmrsozcukanlam.Stop();
                }
                if (timer=="sozcukcumle")
                {
                    
                    tmrsozcukcumle.Stop();
                }
                if (timer == "deyim")
                {
                    tmrdeyim.Stop();
                }
                if (timer == "deyimanlam")
                {
                    tmrdeyimanlam.Stop();
                }
                if (timer == "deyimcumle")
                {
                    tmrdeyimcumle.Stop();
                }
                if (timer == "atasozu")
                {
                    tmratasozu.Stop();
                }

                if (timer == "atasozuyazar")
                {
                    tmryazar.Stop();
                }

                timerPaused = true;
               
            }
        }

       

        private void ResumeTimer()
        {
            if (timerPaused)
            {
                timer1.Start();
                if (timer=="sozcuk")
                {
                    timer2.Start();
                }
                if (timer == "sozcukanlam")
                {
                    tmrsozcukanlam.Start();
                }
                if (timer == "sozcukcumle")
                {
                    tmrsozcukcumle.Start();
                }
                if (timer == "deyim")
                {
                    tmrdeyim.Start();
                }
                if (timer == "deyimanlam")
                {
                    tmrdeyimanlam.Start();
                }
                if (timer == "deyimcumle")
                {
                    tmrdeyimcumle.Start();
                }
                if (timer == "atasozu")
                {
                    tmratasozu.Start();
                }
                if (timer == "atasozuyazar")
                {
                    tmryazar.Start();
                }

                timerPaused = false;
                
            }
        }

        

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        

        private void timer2_Tick(object sender, EventArgs e)
        {
            int indexsozcuk = lblsozcuk.Text.IndexOf(sozcuk);
            if (index < sozcuk.Length)
            {
                

                lblsozcuk.Text += sozcuk[index];

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
            this.WindowState = FormWindowState.Maximized;
            
            label1.Text = DateTime.Now.ToString("dd.MM.yyyy");
            getir();
        }

       

        private void tmrsozcukanlam_Tick(object sender, EventArgs e)
        {
            if (index < sozcukanlam.Length)
            {
<<<<<<< HEAD
               lblsozcukanlam.Text += sozcukanlam[index];
             
=======
              
                lblsozcukanlam.Text += sozcukanlam[index];
                if (sayi == 0)
                {
                    daktilo.Play();
                }
>>>>>>> 1e3164c1f27f86f5c311cb1d2252f9e69f862450
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
<<<<<<< HEAD
                
=======
               
>>>>>>> 1e3164c1f27f86f5c311cb1d2252f9e69f862450
                tmryazar.Stop();
                index = 0;
                timer = "bitti";

               
;                if (sayi == 1&&timer=="bitti")
                {
                    Application.Exit();
                }
               


            }
        }
    }
}
