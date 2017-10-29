namespace WindowsApplication
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.MainLabel = new System.Windows.Forms.Label();
            this.MainSourceButton = new System.Windows.Forms.Button();
            this.RangeInput = new System.Windows.Forms.TextBox();
            this.RangeLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // MainLabel
            // 
            this.MainLabel.AutoSize = true;
            this.MainLabel.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.MainLabel.Location = new System.Drawing.Point(29, 9);
            this.MainLabel.Name = "MainLabel";
            this.MainLabel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.MainLabel.Size = new System.Drawing.Size(350, 154);
            this.MainLabel.TabIndex = 0;
            this.MainLabel.Text = resources.GetString("MainLabel.Text");
            this.MainLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.MainLabel.Click += new System.EventHandler(this.MainLabel_Click);
            // 
            // MainSourceButton
            // 
            this.MainSourceButton.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.MainSourceButton.Location = new System.Drawing.Point(252, 183);
            this.MainSourceButton.Name = "MainSourceButton";
            this.MainSourceButton.Size = new System.Drawing.Size(75, 23);
            this.MainSourceButton.TabIndex = 1;
            this.MainSourceButton.Text = "Start";
            this.MainSourceButton.UseVisualStyleBackColor = true;
            this.MainSourceButton.Click += new System.EventHandler(this.MainSourceButton_Click);
            // 
            // RangeInput
            // 
            this.RangeInput.Location = new System.Drawing.Point(147, 183);
            this.RangeInput.Name = "RangeInput";
            this.RangeInput.Size = new System.Drawing.Size(70, 21);
            this.RangeInput.TabIndex = 2;
            this.RangeInput.TextChanged += new System.EventHandler(this.RangeInput_TextChanged);
            // 
            // RangeLabel
            // 
            this.RangeLabel.AutoSize = true;
            this.RangeLabel.Location = new System.Drawing.Point(30, 190);
            this.RangeLabel.Name = "RangeLabel";
            this.RangeLabel.Size = new System.Drawing.Size(101, 12);
            this.RangeLabel.TabIndex = 3;
            this.RangeLabel.Text = "请输入精度范围：";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(410, 262);
            this.Controls.Add(this.RangeLabel);
            this.Controls.Add(this.RangeInput);
            this.Controls.Add(this.MainSourceButton);
            this.Controls.Add(this.MainLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "MainForm";
            this.Text = "Excel数据处理软件";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label MainLabel;
        private System.Windows.Forms.Button MainSourceButton;
        private System.Windows.Forms.TextBox RangeInput;
        private System.Windows.Forms.Label RangeLabel;
    }
}

