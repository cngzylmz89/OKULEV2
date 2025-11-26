namespace PERFORMANS
{
    partial class frmogrtgiris
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmogrtgiris));
            this.btnsifreac = new System.Windows.Forms.Button();
            this.btntcac = new System.Windows.Forms.Button();
            this.btngiris = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.msktc = new System.Windows.Forms.MaskedTextBox();
            this.txtsifre = new System.Windows.Forms.TextBox();
            this.txtkullaniciadi = new System.Windows.Forms.TextBox();
            this.btnkayit = new System.Windows.Forms.Button();
            this.txtsifree = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox5 = new System.Windows.Forms.PictureBox();
            this.pictureBox6 = new System.Windows.Forms.PictureBox();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnsifreac
            // 
            this.btnsifreac.BackgroundImage = global::PERFORMANS.Properties.Resources.icons8_eye_30px;
            this.btnsifreac.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnsifreac.Location = new System.Drawing.Point(679, 94);
            this.btnsifreac.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnsifreac.Name = "btnsifreac";
            this.btnsifreac.Size = new System.Drawing.Size(53, 33);
            this.btnsifreac.TabIndex = 6;
            this.btnsifreac.UseVisualStyleBackColor = true;
            this.btnsifreac.Click += new System.EventHandler(this.btnsifreac_Click);
            // 
            // btntcac
            // 
            this.btntcac.BackgroundImage = global::PERFORMANS.Properties.Resources.icons8_eye_30px;
            this.btntcac.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btntcac.Location = new System.Drawing.Point(679, 34);
            this.btntcac.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btntcac.Name = "btntcac";
            this.btntcac.Size = new System.Drawing.Size(53, 33);
            this.btntcac.TabIndex = 5;
            this.btntcac.UseVisualStyleBackColor = true;
            this.btntcac.Click += new System.EventHandler(this.btntcac_Click);
            // 
            // btngiris
            // 
            this.btngiris.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(161)))), ((int)(((byte)(224)))));
            this.btngiris.Location = new System.Drawing.Point(452, 153);
            this.btngiris.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btngiris.Name = "btngiris";
            this.btngiris.Size = new System.Drawing.Size(220, 46);
            this.btngiris.TabIndex = 3;
            this.btngiris.Text = "GİRİŞ";
            this.btngiris.UseVisualStyleBackColor = false;
            this.btngiris.Click += new System.EventHandler(this.btngiris_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.MistyRose;
            this.label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label2.Location = new System.Drawing.Point(232, 94);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(127, 27);
            this.label2.TabIndex = 2;
            this.label2.Text = "ŞİFRENİZ:";
            // 
            // msktc
            // 
            this.msktc.BackColor = System.Drawing.SystemColors.HotTrack;
            this.msktc.ForeColor = System.Drawing.SystemColors.Window;
            this.msktc.Location = new System.Drawing.Point(452, 34);
            this.msktc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.msktc.Mask = "00000000000";
            this.msktc.Name = "msktc";
            this.msktc.Size = new System.Drawing.Size(220, 33);
            this.msktc.TabIndex = 1;
            this.msktc.UseSystemPasswordChar = true;
            this.msktc.ValidatingType = typeof(int);
            this.msktc.MouseClick += new System.Windows.Forms.MouseEventHandler(this.msktc_MouseClick);
            // 
            // txtsifre
            // 
            this.txtsifre.BackColor = System.Drawing.SystemColors.HotTrack;
            this.txtsifre.ForeColor = System.Drawing.SystemColors.Window;
            this.txtsifre.Location = new System.Drawing.Point(453, 94);
            this.txtsifre.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtsifre.Name = "txtsifre";
            this.txtsifre.Size = new System.Drawing.Size(219, 33);
            this.txtsifre.TabIndex = 2;
            this.txtsifre.UseSystemPasswordChar = true;
            // 
            // txtkullaniciadi
            // 
            this.txtkullaniciadi.BackColor = System.Drawing.SystemColors.HotTrack;
            this.txtkullaniciadi.ForeColor = System.Drawing.SystemColors.Window;
            this.txtkullaniciadi.Location = new System.Drawing.Point(391, 38);
            this.txtkullaniciadi.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtkullaniciadi.Name = "txtkullaniciadi";
            this.txtkullaniciadi.Size = new System.Drawing.Size(219, 33);
            this.txtkullaniciadi.TabIndex = 9;
            this.txtkullaniciadi.UseSystemPasswordChar = true;
            // 
            // btnkayit
            // 
            this.btnkayit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(8)))), ((int)(((byte)(203)))), ((int)(((byte)(0)))));
            this.btnkayit.Location = new System.Drawing.Point(392, 156);
            this.btnkayit.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnkayit.Name = "btnkayit";
            this.btnkayit.Size = new System.Drawing.Size(220, 46);
            this.btnkayit.TabIndex = 8;
            this.btnkayit.Text = "KAYIT OL";
            this.btnkayit.UseVisualStyleBackColor = false;
            // 
            // txtsifree
            // 
            this.txtsifree.BackColor = System.Drawing.SystemColors.HotTrack;
            this.txtsifree.ForeColor = System.Drawing.SystemColors.Window;
            this.txtsifree.Location = new System.Drawing.Point(392, 94);
            this.txtsifree.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtsifree.Name = "txtsifree";
            this.txtsifree.Size = new System.Drawing.Size(219, 33);
            this.txtsifree.TabIndex = 7;
            this.txtsifree.UseSystemPasswordChar = true;
            // 
            // button2
            // 
            this.button2.BackgroundImage = global::PERFORMANS.Properties.Resources.icons8_eye_30px;
            this.button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button2.Location = new System.Drawing.Point(617, 94);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(53, 33);
            this.button2.TabIndex = 6;
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.BackgroundImage = global::PERFORMANS.Properties.Resources.icons8_eye_30px;
            this.button3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button3.Location = new System.Drawing.Point(617, 34);
            this.button3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(53, 33);
            this.button3.TabIndex = 5;
            this.button3.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.MistyRose;
            this.label3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label3.Location = new System.Drawing.Point(205, 94);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(127, 27);
            this.label3.TabIndex = 2;
            this.label3.Text = "ŞİFRENİZ:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.MistyRose;
            this.label4.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label4.Location = new System.Drawing.Point(96, 41);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(236, 27);
            this.label4.TabIndex = 0;
            this.label4.Text = "KULLANICI ADINIZ:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.MistyRose;
            this.label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label1.Location = new System.Drawing.Point(5, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(297, 27);
            this.label1.TabIndex = 0;
            this.label1.Text = "T.C. KİMLİK NUMARANIZ:";
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.Transparent;
            this.groupBox2.Controls.Add(this.txtkullaniciadi);
            this.groupBox2.Controls.Add(this.btnkayit);
            this.groupBox2.Controls.Add(this.txtsifree);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Location = new System.Drawing.Point(1415, 587);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox2.Size = new System.Drawing.Size(715, 274);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Visible = false;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.Transparent;
            this.groupBox1.Controls.Add(this.txtsifre);
            this.groupBox1.Controls.Add(this.btnsifreac);
            this.groupBox1.Controls.Add(this.btntcac);
            this.groupBox1.Controls.Add(this.btngiris);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.msktc);
            this.groupBox1.Location = new System.Drawing.Point(171, 2);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox1.Size = new System.Drawing.Size(779, 276);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.BackgroundImage = global::PERFORMANS.Properties.Resources.Artboard_1;
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Font = new System.Drawing.Font("Monotype Corsiva", 13.8F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.panel1.Location = new System.Drawing.Point(0, 142);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1069, 561);
            this.panel1.TabIndex = 2;
            // 
            // pictureBox5
            // 
            this.pictureBox5.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox5.Image = global::PERFORMANS.Properties.Resources.icons8_close_40px;
            this.pictureBox5.Location = new System.Drawing.Point(969, 23);
            this.pictureBox5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pictureBox5.Name = "pictureBox5";
            this.pictureBox5.Size = new System.Drawing.Size(67, 50);
            this.pictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox5.TabIndex = 13;
            this.pictureBox5.TabStop = false;
            this.pictureBox5.Click += new System.EventHandler(this.pictureBox5_Click);
            // 
            // pictureBox6
            // 
            this.pictureBox6.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox6.Image = global::PERFORMANS.Properties.Resources.icons8_subtract_40px;
            this.pictureBox6.Location = new System.Drawing.Point(892, 31);
            this.pictureBox6.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pictureBox6.Name = "pictureBox6";
            this.pictureBox6.Size = new System.Drawing.Size(53, 42);
            this.pictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox6.TabIndex = 14;
            this.pictureBox6.TabStop = false;
            this.pictureBox6.Click += new System.EventHandler(this.pictureBox6_Click);
            // 
            // pictureBox4
            // 
            this.pictureBox4.Image = global::PERFORMANS.Properties.Resources.bayrak;
            this.pictureBox4.Location = new System.Drawing.Point(12, 2);
            this.pictureBox4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(99, 84);
            this.pictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox4.TabIndex = 9;
            this.pictureBox4.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.Image = global::PERFORMANS.Properties.Resources.ÖĞRENCİ_PERFORMANS_PROGRAMI;
            this.pictureBox1.Location = new System.Drawing.Point(171, 11);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(667, 62);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.PeachPuff;
            this.panel2.Controls.Add(this.pictureBox5);
            this.panel2.Controls.Add(this.pictureBox6);
            this.panel2.Controls.Add(this.pictureBox4);
            this.panel2.Controls.Add(this.pictureBox1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1069, 105);
            this.panel2.TabIndex = 3;
            // 
            // frmogrtgiris
            // 
            this.AcceptButton = this.btngiris;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1069, 697);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "frmogrtgiris";
            this.Text = "frmyonetici";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnsifreac;
        private System.Windows.Forms.Button btntcac;
        private System.Windows.Forms.Button btngiris;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MaskedTextBox msktc;
        private System.Windows.Forms.TextBox txtsifre;
        private System.Windows.Forms.TextBox txtkullaniciadi;
        private System.Windows.Forms.Button btnkayit;
        private System.Windows.Forms.TextBox txtsifree;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox5;
        private System.Windows.Forms.PictureBox pictureBox6;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panel2;
    }
}