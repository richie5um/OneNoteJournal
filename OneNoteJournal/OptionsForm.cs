using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OneNoteJournal
{
    public partial class OptionsForm : Form
    {
        public OptionsForm()
        {
            InitializeComponent();
        }

        public string DateTimeFormat
        {
            get { return this.comboBoxDateFormat.Text; }
            set { this.comboBoxDateFormat.Text = value; }
        }

        private void UpdateExample()
        {
            try
            {
                this.labelExample.Text = DateTime.Now.ToString( this.comboBoxDateFormat.Text );
            }
            catch ( Exception exception )
            {
                this.labelExample.Text = "Invalid format";
            }
        }

        private void comboBoxDateTime_TextUpdate( object sender, EventArgs e )
        {
            UpdateExample();
        }

        private void comboBoxDateFormat_SelectedIndexChanged( object sender, EventArgs e )
        {
            UpdateExample();
        }
    }
}
