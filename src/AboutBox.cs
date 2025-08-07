using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointSlideThumbnailsAddIn
{
    partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();

            string rtfPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "readme.rtf");
            if (System.IO.File.Exists(rtfPath))
            {
                richTextBox1.LoadFile(rtfPath);
            }
            else
            {
                richTextBox1.Text = "Readme file not found.";
            }

            this.Text = Properties.Strings.about_title + Properties.Strings.title;
            okButton.Text = Properties.Strings.about_ok;

            // Enable clickable links
            richTextBox1.LinkClicked += (sender, e) =>
            {
                try
                {
                    System.Diagnostics.Process.Start(e.LinkText);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening link: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
        }

        #region Assembly Attribute Accessors

        #endregion
    }
}
