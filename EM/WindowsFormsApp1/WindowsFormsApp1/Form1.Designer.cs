namespace WindowsFormsApp1
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.guna2CircleProgressBar1 = new Guna.UI2.WinForms.Guna2CircleProgressBar();
            this.myProgress = new Guna.UI2.WinForms.Guna2CirclePictureBox();
            this.button = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.guna2CircleProgressBar1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.myProgress)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Montserrat SemiBold", 21F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(158)))), ((int)(((byte)(161)))), ((int)(((byte)(176)))));
            this.label1.Location = new System.Drawing.Point(162, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(533, 39);
            this.label1.TabIndex = 0;
            this.label1.Text = "EMPLOYEE MANAGEMENT SYSTEM";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Montserrat SemiBold", 14F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(178)))), ((int)(((byte)(168)))), ((int)(((byte)(176)))));
            this.label2.Location = new System.Drawing.Point(256, 404);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(311, 26);
            this.label2.TabIndex = 1;
            this.label2.Text = "DeveloppedByAbdullaAshurov";
            // 
            // guna2CircleProgressBar1
            // 
            this.guna2CircleProgressBar1.Controls.Add(this.myProgress);
            this.guna2CircleProgressBar1.FillColor = System.Drawing.Color.Cyan;
            this.guna2CircleProgressBar1.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.guna2CircleProgressBar1.ForeColor = System.Drawing.Color.White;
            this.guna2CircleProgressBar1.Location = new System.Drawing.Point(306, 103);
            this.guna2CircleProgressBar1.Minimum = 0;
            this.guna2CircleProgressBar1.Name = "guna2CircleProgressBar1";
            this.guna2CircleProgressBar1.ProgressColor = System.Drawing.Color.IndianRed;
            this.guna2CircleProgressBar1.ProgressColor2 = System.Drawing.Color.IndianRed;
            this.guna2CircleProgressBar1.ProgressThickness = 10;
            this.guna2CircleProgressBar1.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle;
            this.guna2CircleProgressBar1.ShadowDecoration.Parent = this.guna2CircleProgressBar1;
            this.guna2CircleProgressBar1.Size = new System.Drawing.Size(232, 232);
            this.guna2CircleProgressBar1.TabIndex = 2;
            this.guna2CircleProgressBar1.Text = "guna2CircleProgressBar1";
            // 
            // myProgress
            // 
            this.myProgress.BackColor = System.Drawing.Color.Transparent;
            this.myProgress.Image = ((System.Drawing.Image)(resources.GetObject("myProgress.Image")));
            this.myProgress.ImageRotate = 0F;
            this.myProgress.Location = new System.Drawing.Point(24, 23);
            this.myProgress.Name = "myProgress";
            this.myProgress.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle;
            this.myProgress.ShadowDecoration.Parent = this.myProgress;
            this.myProgress.Size = new System.Drawing.Size(184, 186);
            this.myProgress.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.myProgress.TabIndex = 0;
            this.myProgress.TabStop = false;
            this.myProgress.UseTransparentBackground = true;
            // 
            // button
            // 
            this.button.FlatAppearance.BorderSize = 0;
            this.button.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button.ForeColor = System.Drawing.Color.White;
            this.button.Location = new System.Drawing.Point(793, 12);
            this.button.Name = "button";
            this.button.Size = new System.Drawing.Size(25, 25);
            this.button.TabIndex = 3;
            this.button.Text = "X";
            this.button.UseVisualStyleBackColor = true;
            this.button.Click += new System.EventHandler(this.button_Click);
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(54)))), ((int)(((byte)(73)))));
            this.ClientSize = new System.Drawing.Size(830, 439);
            this.Controls.Add(this.button);
            this.Controls.Add(this.guna2CircleProgressBar1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.guna2CircleProgressBar1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.myProgress)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private Guna.UI2.WinForms.Guna2CircleProgressBar guna2CircleProgressBar1;
        private Guna.UI2.WinForms.Guna2CirclePictureBox myProgress;
        private System.Windows.Forms.Button button;
        private System.Windows.Forms.Timer timer1;
    }
}

