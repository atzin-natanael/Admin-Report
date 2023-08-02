namespace Admin_Report
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            Titulo = new Label();
            openFileDialog1 = new OpenFileDialog();
            saveFileDialog1 = new SaveFileDialog();
            monthCalendar1 = new MonthCalendar();
            Btn_Surtido = new Button();
            pictureBox1 = new PictureBox();
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            progressBar1 = new ProgressBar();
            Loading = new Label();
            progressBar2 = new ProgressBar();
            progressBar3 = new ProgressBar();
            Btn_Empaque = new Button();
            Btn_Facturacion = new Button();
            label5 = new Label();
            label6 = new Label();
            Loading2 = new Label();
            Loading3 = new Label();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            // Titulo
            // 
            Titulo.Anchor = AnchorStyles.Top;
            Titulo.AutoSize = true;
            Titulo.Font = new Font("Virgo 01", 15.75F, FontStyle.Regular, GraphicsUnit.Point);
            Titulo.ForeColor = SystemColors.MenuHighlight;
            Titulo.Location = new Point(335, 25);
            Titulo.Name = "Titulo";
            Titulo.Size = new Size(370, 23);
            Titulo.TabIndex = 0;
            Titulo.Text = "Administrador de Surtido";
            Titulo.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // monthCalendar1
            // 
            monthCalendar1.Location = new Point(8, 11);
            monthCalendar1.Name = "monthCalendar1";
            monthCalendar1.TabIndex = 1;
            // 
            // Btn_Surtido
            // 
            Btn_Surtido.Anchor = AnchorStyles.Top;
            Btn_Surtido.BackColor = SystemColors.MenuHighlight;
            Btn_Surtido.Cursor = Cursors.Hand;
            Btn_Surtido.Font = new Font("Sora", 12F, FontStyle.Regular, GraphicsUnit.Point);
            Btn_Surtido.Location = new Point(539, 174);
            Btn_Surtido.Name = "Btn_Surtido";
            Btn_Surtido.Size = new Size(166, 50);
            Btn_Surtido.TabIndex = 2;
            Btn_Surtido.Text = "Elegir";
            Btn_Surtido.UseVisualStyleBackColor = false;
            Btn_Surtido.Click += Btn_Surtido_Click;
            // 
            // pictureBox1
            // 
            pictureBox1.Anchor = AnchorStyles.Bottom;
            pictureBox1.BackColor = Color.Transparent;
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.Location = new Point(565, 615);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(445, 113);
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox1.TabIndex = 14;
            pictureBox1.TabStop = false;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Top;
            label1.AutoSize = true;
            label1.Font = new Font("Sora", 15.75F, FontStyle.Regular, GraphicsUnit.Point);
            label1.ForeColor = SystemColors.MenuHighlight;
            label1.Location = new Point(23, 191);
            label1.Name = "label1";
            label1.Size = new Size(460, 33);
            label1.TabIndex = 15;
            label1.Text = "Seleccionar el archivo semanal de surtido:";
            label1.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Top;
            label2.AutoSize = true;
            label2.Font = new Font("Sora", 15.75F, FontStyle.Regular, GraphicsUnit.Point);
            label2.ForeColor = SystemColors.MenuHighlight;
            label2.Location = new Point(12, 508);
            label2.Name = "label2";
            label2.Size = new Size(261, 33);
            label2.TabIndex = 16;
            label2.Text = "Objetivo del programa:";
            label2.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            label3.Anchor = AnchorStyles.Top;
            label3.AutoSize = true;
            label3.Font = new Font("Sora", 9.749999F, FontStyle.Regular, GraphicsUnit.Point);
            label3.Location = new Point(8, 552);
            label3.Name = "label3";
            label3.Size = new Size(628, 60);
            label3.TabIndex = 17;
            label3.Text = resources.GetString("label3.Text");
            // 
            // label4
            // 
            label4.Anchor = AnchorStyles.Bottom;
            label4.AutoSize = true;
            label4.Font = new Font("Elephant", 14.25F, FontStyle.Regular, GraphicsUnit.Point);
            label4.ForeColor = SystemColors.ActiveCaptionText;
            label4.Location = new Point(8, 695);
            label4.Name = "label4";
            label4.Size = new Size(344, 25);
            label4.TabIndex = 18;
            label4.Text = "Desarrollado por: Atzin Not Found";
            label4.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // progressBar1
            // 
            progressBar1.Anchor = AnchorStyles.Top;
            progressBar1.Location = new Point(773, 189);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(173, 35);
            progressBar1.TabIndex = 19;
            progressBar1.Visible = false;
            // 
            // Loading
            // 
            Loading.Anchor = AnchorStyles.Top;
            Loading.AutoSize = true;
            Loading.Font = new Font("Century", 15.75F, FontStyle.Bold, GraphicsUnit.Point);
            Loading.ForeColor = SystemColors.InfoText;
            Loading.Location = new Point(797, 148);
            Loading.Name = "Loading";
            Loading.Size = new Size(136, 25);
            Loading.TabIndex = 20;
            Loading.Text = "Cargando...";
            Loading.TextAlign = ContentAlignment.MiddleCenter;
            Loading.Visible = false;
            // 
            // progressBar2
            // 
            progressBar2.Anchor = AnchorStyles.Top;
            progressBar2.Location = new Point(773, 289);
            progressBar2.Name = "progressBar2";
            progressBar2.Size = new Size(173, 35);
            progressBar2.TabIndex = 21;
            progressBar2.Visible = false;
            // 
            // progressBar3
            // 
            progressBar3.Anchor = AnchorStyles.Top;
            progressBar3.Location = new Point(773, 395);
            progressBar3.Name = "progressBar3";
            progressBar3.Size = new Size(173, 35);
            progressBar3.TabIndex = 22;
            progressBar3.Visible = false;
            // 
            // Btn_Empaque
            // 
            Btn_Empaque.Anchor = AnchorStyles.Top;
            Btn_Empaque.BackColor = Color.RosyBrown;
            Btn_Empaque.Cursor = Cursors.Hand;
            Btn_Empaque.Font = new Font("Sora", 12F, FontStyle.Regular, GraphicsUnit.Point);
            Btn_Empaque.Location = new Point(539, 274);
            Btn_Empaque.Name = "Btn_Empaque";
            Btn_Empaque.Size = new Size(166, 50);
            Btn_Empaque.TabIndex = 23;
            Btn_Empaque.Text = "Elegir";
            Btn_Empaque.UseVisualStyleBackColor = false;
            Btn_Empaque.Click += Btn_Empaque_Click;
            // 
            // Btn_Facturacion
            // 
            Btn_Facturacion.Anchor = AnchorStyles.Top;
            Btn_Facturacion.BackColor = Color.SeaGreen;
            Btn_Facturacion.Cursor = Cursors.Hand;
            Btn_Facturacion.Font = new Font("Sora", 12F, FontStyle.Regular, GraphicsUnit.Point);
            Btn_Facturacion.Location = new Point(539, 380);
            Btn_Facturacion.Name = "Btn_Facturacion";
            Btn_Facturacion.Size = new Size(166, 50);
            Btn_Facturacion.TabIndex = 24;
            Btn_Facturacion.Text = "Elegir";
            Btn_Facturacion.UseVisualStyleBackColor = false;
            Btn_Facturacion.Click += Btn_Facturacion_Click;
            // 
            // label5
            // 
            label5.Anchor = AnchorStyles.Top;
            label5.AutoSize = true;
            label5.Font = new Font("Sora", 15.75F, FontStyle.Regular, GraphicsUnit.Point);
            label5.ForeColor = Color.RosyBrown;
            label5.Location = new Point(23, 291);
            label5.Name = "label5";
            label5.Size = new Size(483, 33);
            label5.TabIndex = 25;
            label5.Text = "Seleccionar el archivo semanal de Empaque:";
            label5.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            label6.Anchor = AnchorStyles.Top;
            label6.AutoSize = true;
            label6.BackColor = Color.Transparent;
            label6.Font = new Font("Sora", 15.75F, FontStyle.Regular, GraphicsUnit.Point);
            label6.ForeColor = Color.SeaGreen;
            label6.Location = new Point(23, 397);
            label6.Name = "label6";
            label6.Size = new Size(510, 33);
            label6.TabIndex = 26;
            label6.Text = "Seleccionar el archivo semanal de Facturación:";
            label6.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // Loading2
            // 
            Loading2.Anchor = AnchorStyles.Top;
            Loading2.AutoSize = true;
            Loading2.Font = new Font("Century", 15.75F, FontStyle.Bold, GraphicsUnit.Point);
            Loading2.ForeColor = SystemColors.InfoText;
            Loading2.Location = new Point(797, 252);
            Loading2.Name = "Loading2";
            Loading2.Size = new Size(136, 25);
            Loading2.TabIndex = 27;
            Loading2.Text = "Cargando...";
            Loading2.TextAlign = ContentAlignment.MiddleCenter;
            Loading2.Visible = false;
            // 
            // Loading3
            // 
            Loading3.Anchor = AnchorStyles.Top;
            Loading3.AutoSize = true;
            Loading3.Font = new Font("Century", 15.75F, FontStyle.Bold, GraphicsUnit.Point);
            Loading3.ForeColor = SystemColors.InfoText;
            Loading3.Location = new Point(797, 367);
            Loading3.Name = "Loading3";
            Loading3.Size = new Size(136, 25);
            Loading3.TabIndex = 28;
            Loading3.Text = "Cargando...";
            Loading3.TextAlign = ContentAlignment.MiddleCenter;
            Loading3.Visible = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.DimGray;
            ClientSize = new Size(1008, 729);
            Controls.Add(Loading3);
            Controls.Add(Loading2);
            Controls.Add(label6);
            Controls.Add(label5);
            Controls.Add(Btn_Facturacion);
            Controls.Add(Btn_Empaque);
            Controls.Add(progressBar3);
            Controls.Add(progressBar2);
            Controls.Add(Loading);
            Controls.Add(progressBar1);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(pictureBox1);
            Controls.Add(Btn_Surtido);
            Controls.Add(monthCalendar1);
            Controls.Add(Titulo);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label Titulo;
        private OpenFileDialog openFileDialog1;
        private SaveFileDialog saveFileDialog1;
        private MonthCalendar monthCalendar1;
        private Button Btn_Surtido;
        private PictureBox pictureBox1;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private ProgressBar progressBar1;
        private Label Loading;
        private ProgressBar progressBar2;
        private ProgressBar progressBar3;
        private Button Btn_Empaque;
        private Button Btn_Facturacion;
        private Label label5;
        private Label label6;
        private Label Loading2;
        private Label Loading3;
    }
}