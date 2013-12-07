namespace AuditTool
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnAudit = new System.Windows.Forms.Button();
            this.btnServer = new System.Windows.Forms.Button();
            this.btnClient = new System.Windows.Forms.Button();
            this.txtMessage = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkTimer = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // btnAudit
            // 
            this.btnAudit.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnAudit.Location = new System.Drawing.Point(466, 291);
            this.btnAudit.Name = "btnAudit";
            this.btnAudit.Size = new System.Drawing.Size(106, 59);
            this.btnAudit.TabIndex = 0;
            this.btnAudit.Text = "進行稽核";
            this.btnAudit.UseVisualStyleBackColor = true;
            this.btnAudit.Click += new System.EventHandler(this.btnAudit_Click);
            // 
            // btnServer
            // 
            this.btnServer.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnServer.Location = new System.Drawing.Point(466, 43);
            this.btnServer.Name = "btnServer";
            this.btnServer.Size = new System.Drawing.Size(106, 59);
            this.btnServer.TabIndex = 1;
            this.btnServer.Text = "啟動伺服器";
            this.btnServer.UseVisualStyleBackColor = true;
            this.btnServer.Click += new System.EventHandler(this.btnServer_Click);
            // 
            // btnClient
            // 
            this.btnClient.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnClient.Location = new System.Drawing.Point(466, 108);
            this.btnClient.Name = "btnClient";
            this.btnClient.Size = new System.Drawing.Size(106, 58);
            this.btnClient.TabIndex = 2;
            this.btnClient.Text = "傳輸資料";
            this.btnClient.UseVisualStyleBackColor = true;
            this.btnClient.Click += new System.EventHandler(this.btnClient_Click);
            // 
            // txtMessage
            // 
            this.txtMessage.Font = new System.Drawing.Font("微軟正黑體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtMessage.Location = new System.Drawing.Point(12, 43);
            this.txtMessage.Multiline = true;
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(431, 307);
            this.txtMessage.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(8, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "訊息記錄：";
            // 
            // checkTimer
            // 
            this.checkTimer.Enabled = true;
            this.checkTimer.Tick += new System.EventHandler(this.checkTimer_Tick);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 362);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.btnClient);
            this.Controls.Add(this.btnServer);
            this.Controls.Add(this.btnAudit);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "檔案稽核小工具";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAudit;
        private System.Windows.Forms.Button btnServer;
        private System.Windows.Forms.Button btnClient;
        private System.Windows.Forms.TextBox txtMessage;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Timer checkTimer;
    }
}

