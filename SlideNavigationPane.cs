using System;
using System.Windows.Forms;

namespace PowerPointSlideThumbnailsAddIn
{
    public partial class SlideNavigationPane : UserControl
    {
        public event EventHandler LeftArrowClicked;
        public event EventHandler RightArrowClicked;

        public SlideNavigationPane()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.btnLeft = new Button();
            this.btnRight = new Button();
            // 
            // btnLeft
            // 
            this.btnLeft.Font = new System.Drawing.Font("Segoe UI", 36F, System.Drawing.FontStyle.Bold);
            this.btnLeft.Text = "<";
            this.btnLeft.Width = 120;
            this.btnLeft.Height = 120;
            this.btnLeft.Left = 10;
            this.btnLeft.Top = 10;
            this.btnLeft.Click += (s, e) => LeftArrowClicked?.Invoke(this, EventArgs.Empty);
            // 
            // btnRight
            // 
            this.btnRight.Font = new System.Drawing.Font("Segoe UI", 36F, System.Drawing.FontStyle.Bold);
            this.btnRight.Text = ">";
            this.btnRight.Width = 120;
            this.btnRight.Height = 120;
            this.btnRight.Left = 150;
            this.btnRight.Top = 10;
            this.btnRight.Click += (s, e) => RightArrowClicked?.Invoke(this, EventArgs.Empty);
            // 
            // SlideNavigationPane
            // 
            this.Controls.Add(this.btnLeft);
            this.Controls.Add(this.btnRight);
            this.Width = 290;
            this.Height = 140;
        }

        private Button btnLeft;
        private Button btnRight;
    }
}
