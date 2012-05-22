using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Runtime.InteropServices;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace OneNoteJournal
{
    public partial class MainForm : Form
    {
        private HotKey hotKey = new HotKey();
        private string strPageDateTimeFormat = "dddd, dd MMMM yyyy";
        private string strSettingsFile = ".OneNoteJournal.Settings";

        public MainForm()
        {
            InitializeComponent();

            hotKey.KeyPressed += new EventHandler<KeyPressedEventArgs>( HotKey_KeyPressed );
            hotKey.RegisterHotKey( OneNoteJournal.ModifierKeys.Win, Keys.J );
        }

        private void LoadOptions()
        {
            string strSettingsFile = Path.Combine( Environment.CurrentDirectory, this.strSettingsFile );
            if ( File.Exists( strSettingsFile ) )
            {
                XElement xElement = XElement.Load( strSettingsFile );
                string strPageDateTimeFormatXML = (
                        from xItem in xElement.Descendants( "PageDateTimeFormat" )
                        select xItem.Value
                    ).First();

                if ( !String.IsNullOrWhiteSpace( strPageDateTimeFormatXML ) )
                {
                    this.strPageDateTimeFormat = strPageDateTimeFormatXML;
                }
            }
         }

        private void SaveOptions()
        { 
            XElement xElement = 
                new XElement( "OneNoteJournal",
                    new XElement( "Options",
                        new XElement( "PageDateTimeFormat", this.strPageDateTimeFormat )
                    )
                );

            xElement.Save( Path.Combine( Environment.CurrentDirectory, this.strSettingsFile ) );
        }

        private void CreateJournalEntry()
        {
            OneNoteProvider oneNoteProvider = new OneNoteProvider();

            try
            {
                string strNotebook = "Journal";
                string strNotebookId = oneNoteProvider.GetNotebookIdFromName( strNotebook );

                if ( !String.IsNullOrEmpty( strNotebookId ) )
                {
                    string strSectionGroup = DateTime.Now.ToString( "yyyy" );
                    string strSectionGroupId = oneNoteProvider.CreateSectionGroupFromParentId( strNotebookId, strSectionGroup );

                    if ( !String.IsNullOrEmpty( strSectionGroupId ) )
                    {
                        string strSection = DateTime.Now.ToString( "MMMM" );
                        string strSectionId = oneNoteProvider.CreateSectionFromParentId( strSectionGroupId, strSection );

                        if ( !String.IsNullOrEmpty( strSectionId ) )
                        {
                            string strPage = DateTime.Now.ToString( this.strPageDateTimeFormat );
                            string strPageId = oneNoteProvider.CreatePageFromParentId( strSectionId, strPage );

                            if ( !String.IsNullOrEmpty( strPageId ) )
                            {
                                oneNoteProvider.NavigateToId( strPageId );
                            }
                            else
                            {
                                MessageBox.Show( "Unable to create/open Page.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error );
                            }
                        }
                        else
                        {
                            MessageBox.Show( "Unable to create Section.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error );
                        }
                    }
                    else
                    {
                        MessageBox.Show( "Unable to create Section Folder.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error );
                    }
                }
                else
                {
                    MessageBox.Show( "Unable to access \"Journal\" Notebook.\nPlease ensure it exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop );
                }
            }
            catch ( Exception e )
            {
                MessageBox.Show( "Fatal Error\n\n" + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop );
            }
        }

        ///////////////////////////////////////////////////////////////////////////////////

        private void MainForm_Shown( object sender, EventArgs e )
        {
            this.WindowState = FormWindowState.Minimized;
            this.Hide();

            this.notifyIconMain.ShowBalloonTip( 2 * 1000 );
        }

        void HotKey_KeyPressed( object sender, KeyPressedEventArgs e )
        {
            CreateJournalEntry();
        }

        private void createJournalEntryToolStripMenuItem_Click( object sender, EventArgs e )
        {
            CreateJournalEntry();
        }

        private void exitToolStripMenuItem_Click( object sender, EventArgs e )
        {
            this.Close();
        }

        private void optionsToolStripMenuItem_Click( object sender, EventArgs e )
        {
            OptionsForm options = new OptionsForm();
            options.DateTimeFormat = this.strPageDateTimeFormat;
            if ( DialogResult.OK == options.ShowDialog() )
            {
                if ( !String.IsNullOrWhiteSpace( options.DateTimeFormat ) )
                {
                    this.strPageDateTimeFormat = options.DateTimeFormat;
                }
            }
        }

        private void MainForm_FormClosing( object sender, FormClosingEventArgs e )
        {
            SaveOptions();
        }

        private void MainForm_Load( object sender, EventArgs e )
        {
            LoadOptions();
        }
    }
}
