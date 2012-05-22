namespace OneNoteJournal
{
    partial class OptionsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comboBoxDateFormat = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.labelExample = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right ) ) );
            this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOK.Location = new System.Drawing.Point( 116, 149 );
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size( 75, 23 );
            this.buttonOK.TabIndex = 0;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right ) ) );
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point( 197, 149 );
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size( 75, 23 );
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( ( ( System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom )
                        | System.Windows.Forms.AnchorStyles.Left )
                        | System.Windows.Forms.AnchorStyles.Right ) ) );
            this.groupBox1.Controls.Add( this.label1 );
            this.groupBox1.Controls.Add( this.groupBox2 );
            this.groupBox1.Controls.Add( this.comboBoxDateFormat );
            this.groupBox1.Location = new System.Drawing.Point( 12, 12 );
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size( 260, 131 );
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Options";
            // 
            // comboBoxDateFormat
            // 
            this.comboBoxDateFormat.FormattingEnabled = true;
            this.comboBoxDateFormat.Items.AddRange( new object[] {
            "MM/dd/yyyy",
            "MM-dd-yyyy",
            "dddd, dd MMMM yyyy",
            "ddd, dd MMM yyyy",
            "MMMM dd"} );
            this.comboBoxDateFormat.Location = new System.Drawing.Point( 104, 19 );
            this.comboBoxDateFormat.Name = "comboBoxDateFormat";
            this.comboBoxDateFormat.Size = new System.Drawing.Size( 138, 21 );
            this.comboBoxDateFormat.TabIndex = 0;
            this.comboBoxDateFormat.SelectedIndexChanged += new System.EventHandler( this.comboBoxDateFormat_SelectedIndexChanged );
            this.comboBoxDateFormat.TextUpdate += new System.EventHandler( this.comboBoxDateTime_TextUpdate );
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add( this.labelExample );
            this.groupBox2.Location = new System.Drawing.Point( 18, 46 );
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size( 224, 68 );
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Example";
            // 
            // labelExample
            // 
            this.labelExample.Anchor = ( (System.Windows.Forms.AnchorStyles)( ( ( System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left )
                        | System.Windows.Forms.AnchorStyles.Right ) ) );
            this.labelExample.Location = new System.Drawing.Point( 20, 32 );
            this.labelExample.Name = "labelExample";
            this.labelExample.Size = new System.Drawing.Size( 187, 23 );
            this.labelExample.TabIndex = 0;
            this.labelExample.Text = "...";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point( 15, 22 );
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size( 68, 13 );
            this.label1.TabIndex = 2;
            this.label1.Text = "Date Format:";
            // 
            // OptionsForm
            // 
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF( 6F, 13F );
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size( 284, 184 );
            this.Controls.Add( this.groupBox1 );
            this.Controls.Add( this.buttonCancel );
            this.Controls.Add( this.buttonOK );
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "OptionsForm";
            this.Text = "OneNote Journal - Options";
            this.groupBox1.ResumeLayout( false );
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout( false );
            this.ResumeLayout( false );

        }

        #endregion

        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label labelExample;
        private System.Windows.Forms.ComboBox comboBoxDateFormat;
    }
}