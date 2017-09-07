namespace Demo_Excel_Export
{
    partial class runMonthly
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(runMonthly));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.startDate = new System.Windows.Forms.DateTimePicker();
            this.endDate = new System.Windows.Forms.DateTimePicker();
            this.run = new System.Windows.Forms.Button();
            this.sendMail = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Strat Date";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "End Date";
            // 
            // startDate
            // 
            this.startDate.Location = new System.Drawing.Point(104, 41);
            this.startDate.Name = "startDate";
            this.startDate.Size = new System.Drawing.Size(200, 20);
            this.startDate.TabIndex = 2;
            this.startDate.Value = new System.DateTime(2016, 8, 11, 0, 0, 0, 0);
            
            // 
            // endDate
            // 
            this.endDate.Location = new System.Drawing.Point(103, 71);
            this.endDate.Name = "endDate";
            this.endDate.Size = new System.Drawing.Size(201, 20);
            this.endDate.TabIndex = 3;
            this.endDate.Value = new System.DateTime(2016, 8, 11, 0, 0, 0, 0);
            // 
            // run
            // 
            this.run.Location = new System.Drawing.Point(366, 17);
            this.run.Name = "run";
            this.run.Size = new System.Drawing.Size(123, 29);
            this.run.TabIndex = 4;
            this.run.Text = "Daily";
            this.run.UseVisualStyleBackColor = true;
            this.run.Click += new System.EventHandler(this.run_Click);
            // 
            // sendMail
            // 
            this.sendMail.Location = new System.Drawing.Point(366, 58);
            this.sendMail.Name = "sendMail";
            this.sendMail.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.sendMail.Size = new System.Drawing.Size(123, 29);
            this.sendMail.TabIndex = 5;
            this.sendMail.Text = "Send Mail";
            this.sendMail.UseVisualStyleBackColor = true;
            this.sendMail.Click += new System.EventHandler(this.sendMail_Click);
            // 
            // runMonthly
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(564, 128);
            this.Controls.Add(this.sendMail);
            this.Controls.Add(this.run);
            this.Controls.Add(this.endDate);
            this.Controls.Add(this.startDate);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "runMonthly";
            this.Text = "BTRCDaily";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker startDate;
        private System.Windows.Forms.DateTimePicker endDate;
        private System.Windows.Forms.Button run;
        private System.Windows.Forms.Button sendMail;
       
        
    }
}