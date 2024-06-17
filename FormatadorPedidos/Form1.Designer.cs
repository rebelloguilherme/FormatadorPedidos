namespace FormatadorPedidos
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        // Declare UI components
        protected System.Windows.Forms.Button btnOpenFile;
        protected System.Windows.Forms.Button btnSaveFile;
        protected System.Windows.Forms.TextBox txtFilePath;

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
            btnOpenFile = new Button();
            btnSaveFile = new Button();
            txtFilePath = new TextBox();
            SuspendLayout();
            // 
            // btnOpenFile
            // 
            btnOpenFile.Location = new Point(50, 50);
            btnOpenFile.Name = "btnOpenFile";
            btnOpenFile.Size = new Size(75, 23);
            btnOpenFile.TabIndex = 0;
            btnOpenFile.Text = "Abrir arquivo";
            btnOpenFile.UseVisualStyleBackColor = true;
            btnOpenFile.Click += btnOpenFile_Click;
            // 
            // btnSaveFile
            // 
            btnSaveFile.Location = new Point(50, 100);
            btnSaveFile.Name = "btnSaveFile";
            btnSaveFile.Size = new Size(75, 23);
            btnSaveFile.TabIndex = 1;
            btnSaveFile.Text = "Formatar";
            btnSaveFile.UseVisualStyleBackColor = true;
            btnSaveFile.Click += btnSaveFile_Click;
            // 
            // txtFilePath
            // 
            txtFilePath.Location = new Point(150, 50);
            txtFilePath.Name = "txtFilePath";
            txtFilePath.Size = new Size(600, 23);
            txtFilePath.TabIndex = 2;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(768, 139);
            Controls.Add(txtFilePath);
            Controls.Add(btnSaveFile);
            Controls.Add(btnOpenFile);
            Name = "Form1";
            Text = "Fomatador de Pedidos";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
