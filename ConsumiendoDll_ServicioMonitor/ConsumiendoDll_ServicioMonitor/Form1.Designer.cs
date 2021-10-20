
namespace ConsumiendoDll_ServicioMonitor
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnTest = new System.Windows.Forms.Button();
            this.tmrRestar = new System.Windows.Forms.Timer(this.components);
            this.tmrMonitorMQTKT = new System.Windows.Forms.Timer(this.components);
            this.txtErrores = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(641, 195);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(262, 94);
            this.btnTest.TabIndex = 0;
            this.btnTest.Text = "btnTest";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // tmrRestar
            // 
            this.tmrRestar.Tick += new System.EventHandler(this.tmrRestar_Tick);
            // 
            // tmrMonitorMQTKT
            // 
            this.tmrMonitorMQTKT.Enabled = true;
            this.tmrMonitorMQTKT.Tick += new System.EventHandler(this.tmrMonitorMQTKT_Tick);
            // 
            // txtErrores
            // 
            this.txtErrores.Location = new System.Drawing.Point(12, 329);
            this.txtErrores.Multiline = true;
            this.txtErrores.Name = "txtErrores";
            this.txtErrores.Size = new System.Drawing.Size(891, 206);
            this.txtErrores.TabIndex = 1;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(915, 547);
            this.Controls.Add(this.txtErrores);
            this.Controls.Add(this.btnTest);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Timer tmrRestar;
        private System.Windows.Forms.Timer tmrMonitorMQTKT;
        private System.Windows.Forms.TextBox txtErrores;
    }
}

