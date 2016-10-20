partial class Form1
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
        this.groupBox1 = new System.Windows.Forms.GroupBox();
        this.label3 = new System.Windows.Forms.Label();
        this.fwTxt = new System.Windows.Forms.TextBox();
        this.label2 = new System.Windows.Forms.Label();
        this.typTxt = new System.Windows.Forms.TextBox();
        this.label1 = new System.Windows.Forms.Label();
        this.button1 = new System.Windows.Forms.Button();
        this.textBox1 = new System.Windows.Forms.TextBox();
        this.dataGridView1 = new System.Windows.Forms.DataGridView();
        this.groupBox1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
        this.SuspendLayout();
        // 
        // groupBox1
        // 
        this.groupBox1.Controls.Add(this.label3);
        this.groupBox1.Controls.Add(this.fwTxt);
        this.groupBox1.Controls.Add(this.label2);
        this.groupBox1.Controls.Add(this.typTxt);
        this.groupBox1.Controls.Add(this.label1);
        this.groupBox1.Controls.Add(this.button1);
        this.groupBox1.Controls.Add(this.textBox1);
        this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
        this.groupBox1.Location = new System.Drawing.Point(0, 0);
        this.groupBox1.Name = "groupBox1";
        this.groupBox1.Size = new System.Drawing.Size(503, 131);
        this.groupBox1.TabIndex = 0;
        this.groupBox1.TabStop = false;
        this.groupBox1.Text = "groupBox1";
        // 
        // label3
        // 
        this.label3.AutoSize = true;
        this.label3.Location = new System.Drawing.Point(11, 87);
        this.label3.Name = "label3";
        this.label3.Size = new System.Drawing.Size(22, 12);
        this.label3.TabIndex = 36;
        this.label3.Text = "Fw.";
        // 
        // fwTxt
        // 
        this.fwTxt.Location = new System.Drawing.Point(50, 84);
        this.fwTxt.Name = "fwTxt";
        this.fwTxt.Size = new System.Drawing.Size(111, 22);
        this.fwTxt.TabIndex = 35;
        // 
        // label2
        // 
        this.label2.AutoSize = true;
        this.label2.Location = new System.Drawing.Point(11, 59);
        this.label2.Name = "label2";
        this.label2.Size = new System.Drawing.Size(32, 12);
        this.label2.TabIndex = 34;
        this.label2.Text = "Type.";
        // 
        // typTxt
        // 
        this.typTxt.Location = new System.Drawing.Point(50, 56);
        this.typTxt.Name = "typTxt";
        this.typTxt.Size = new System.Drawing.Size(111, 22);
        this.typTxt.TabIndex = 33;
        // 
        // label1
        // 
        this.label1.AutoSize = true;
        this.label1.Location = new System.Drawing.Point(11, 24);
        this.label1.Name = "label1";
        this.label1.Size = new System.Drawing.Size(15, 12);
        this.label1.TabIndex = 32;
        this.label1.Text = "IP";
        // 
        // button1
        // 
        this.button1.Location = new System.Drawing.Point(167, 21);
        this.button1.Name = "button1";
        this.button1.Size = new System.Drawing.Size(76, 85);
        this.button1.TabIndex = 31;
        this.button1.Text = "Connect";
        this.button1.UseVisualStyleBackColor = true;
        this.button1.Click += new System.EventHandler(this.button1_Click);
        // 
        // textBox1
        // 
        this.textBox1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
        this.textBox1.Location = new System.Drawing.Point(50, 21);
        this.textBox1.Name = "textBox1";
        this.textBox1.Size = new System.Drawing.Size(111, 27);
        this.textBox1.TabIndex = 30;
        this.textBox1.Text = "192.168.1.1";
        // 
        // dataGridView1
        // 
        this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dataGridView1.Location = new System.Drawing.Point(0, 131);
        this.dataGridView1.Name = "dataGridView1";
        this.dataGridView1.RowTemplate.Height = 24;
        this.dataGridView1.Size = new System.Drawing.Size(503, 209);
        this.dataGridView1.TabIndex = 1;
        // 
        // Form1
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(503, 340);
        this.Controls.Add(this.dataGridView1);
        this.Controls.Add(this.groupBox1);
        this.Name = "Form1";
        this.Text = "WISE_RESTAPI_Main";
        this.Load += new System.EventHandler(this.Form1_Load);
        this.groupBox1.ResumeLayout(false);
        this.groupBox1.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
        this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.GroupBox groupBox1;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.TextBox typTxt;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Button button1;
    private System.Windows.Forms.TextBox textBox1;
    private System.Windows.Forms.DataGridView dataGridView1;
    private System.Windows.Forms.Label label3;
    private System.Windows.Forms.TextBox fwTxt;
}

