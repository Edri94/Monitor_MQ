
namespace MonitorMQTKT
{
    partial class MonitorMQTKT
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

        #region Código generado por el Diseñador de componentes

        /// <summary> 
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tmrMonitorMQTKT = new System.Timers.Timer();
            this.tmrRestar = new System.Timers.Timer();
            ((System.ComponentModel.ISupportInitialize)(this.tmrMonitorMQTKT)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tmrRestar)).BeginInit();
            // 
            // tmrMonitorMQTKT
            // 
            this.tmrMonitorMQTKT.Enabled = true;
            this.tmrMonitorMQTKT.Interval = 30000D;
            this.tmrMonitorMQTKT.Elapsed += new System.Timers.ElapsedEventHandler(this.tmrMonitorMQTKT_Elapsed);
            // 
            // tmrRestar
            // 
            this.tmrRestar.AutoReset = false;
            this.tmrRestar.Elapsed += new System.Timers.ElapsedEventHandler(this.tmrRestar_Elapsed);
            // 
            // MonitorMQTKT
            // 
            this.ServiceName = "MonitorMQTKT";
            ((System.ComponentModel.ISupportInitialize)(this.tmrMonitorMQTKT)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tmrRestar)).EndInit();

        }

        #endregion

        private System.Timers.Timer tmrMonitorMQTKT;
        private System.Timers.Timer tmrRestar;
    }
}
