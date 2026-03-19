namespace PowerPointAddIn1
{
    partial class QuizPaneControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label1 = new System.Windows.Forms.Label();
            this.rtbQuizOutput = new System.Windows.Forms.RichTextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.numQuestionCount = new System.Windows.Forms.NumericUpDown();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.numQuestionCount)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(36, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "AI Quiz Generator";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // rtbQuizOutput
            // 
            this.rtbQuizOutput.Location = new System.Drawing.Point(39, 181);
            this.rtbQuizOutput.Name = "rtbQuizOutput";
            this.rtbQuizOutput.ReadOnly = true;
            this.rtbQuizOutput.Size = new System.Drawing.Size(146, 150);
            this.rtbQuizOutput.TabIndex = 1;
            this.rtbQuizOutput.Text = "";
            this.rtbQuizOutput.TextChanged += new System.EventHandler(this.rtbQuizOutput_TextChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(72, 382);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Generate Quiz";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // numQuestionCount
            // 
            this.numQuestionCount.Location = new System.Drawing.Point(39, 99);
            this.numQuestionCount.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numQuestionCount.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numQuestionCount.Name = "numQuestionCount";
            this.numQuestionCount.Size = new System.Drawing.Size(120, 22);
            this.numQuestionCount.TabIndex = 3;
            this.numQuestionCount.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
            this.numQuestionCount.ValueChanged += new System.EventHandler(this.numQuestionCount_ValueChanged);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // QuizPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.numQuestionCount);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.rtbQuizOutput);
            this.Controls.Add(this.label1);
            this.Name = "QuizPaneControl";
            this.Size = new System.Drawing.Size(233, 490);
            this.Load += new System.EventHandler(this.QuizPaneControl_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numQuestionCount)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox rtbQuizOutput;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.NumericUpDown numQuestionCount;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
    }
}
