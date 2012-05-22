namespace OneNoteJournal
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager( typeof( MainForm ) );
            this.notifyIconMain = new System.Windows.Forms.NotifyIcon( this.components );
            this.contextMenuStripMain = new System.Windows.Forms.ContextMenuStrip( this.components );
            this.createJournalEntryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.contextMenuStripMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // notifyIconMain
            // 
            this.notifyIconMain.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyIconMain.BalloonTipText = "Press \"Win + J\"";
            this.notifyIconMain.BalloonTipTitle = "OneNote Journal v1.1";
            this.notifyIconMain.ContextMenuStrip = this.contextMenuStripMain;
            this.notifyIconMain.Icon = ( (System.Drawing.Icon)( resources.GetObject( "notifyIconMain.Icon" ) ) );
            this.notifyIconMain.Text = "OneNote Journal v1.1";
            this.notifyIconMain.Visible = true;
            // 
            // contextMenuStripMain
            // 
            this.contextMenuStripMain.Items.AddRange( new System.Windows.Forms.ToolStripItem[] {
            this.optionsToolStripMenuItem,
            this.createJournalEntryToolStripMenuItem,
            this.toolStripSeparator1,
            this.exitToolStripMenuItem} );
            this.contextMenuStripMain.Name = "contextMenuStripMain";
            this.contextMenuStripMain.Size = new System.Drawing.Size( 173, 76 );
            // 
            // createJournalEntryToolStripMenuItem
            // 
            this.createJournalEntryToolStripMenuItem.Name = "createJournalEntryToolStripMenuItem";
            this.createJournalEntryToolStripMenuItem.Size = new System.Drawing.Size( 172, 22 );
            this.createJournalEntryToolStripMenuItem.Text = "Create Journal Entry";
            this.createJournalEntryToolStripMenuItem.Click += new System.EventHandler( this.createJournalEntryToolStripMenuItem_Click );
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size( 172, 22 );
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler( this.exitToolStripMenuItem_Click );
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size( 172, 22 );
            this.optionsToolStripMenuItem.Text = "Options";
            this.optionsToolStripMenuItem.Click += new System.EventHandler( this.optionsToolStripMenuItem_Click );
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size( 169, 6 );
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF( 6F, 13F );
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size( 322, 69 );
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "MainForm";
            this.ShowInTaskbar = false;
            this.Text = "OneNote Journal";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler( this.MainForm_FormClosing );
            this.Load += new System.EventHandler( this.MainForm_Load );
            this.Shown += new System.EventHandler( this.MainForm_Shown );
            this.contextMenuStripMain.ResumeLayout( false );
            this.ResumeLayout( false );

        }

        #endregion

        private System.Windows.Forms.NotifyIcon notifyIconMain;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripMain;
        private System.Windows.Forms.ToolStripMenuItem createJournalEntryToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
    }
}

