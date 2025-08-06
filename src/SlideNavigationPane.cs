using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace PowerPointSlideThumbnailsAddIn
{
    public partial class SlideNavigationPane : UserControl
    {
        public event EventHandler LeftArrowClicked;
        public event EventHandler RightArrowClicked;
        public event EventHandler BackToGridClicked;
        public event EventHandler EndButtonClicked;
        public event EventHandler DockToBottomClicked;
        public event EventHandler DockToRightClicked;
        public event EventHandler btnAbout_Click;

        private PrivateFontCollection _privateFonts = new PrivateFontCollection();
        private bool _fontLoaded = false;
        private Button btnLeft;
        private Button btnRight;
        private Button btnBackToGrid;
        private Button btnEnd;
        private Button btnDockBottom;
        private Button btnDockRight;
        private Button btnAbout;
        private Panel linePanel;

        private Label lblBackToGrid;
        private Label lblEnd;

        private ToolTip toolTipLeft;
        private ToolTip toolTipRight;
        private ToolTip toolTipBackToGrid;
        private ToolTip toolTipEnd;
        private ToolTip toolTipDockBottom;
        private ToolTip toolTipDockRight;
        private ToolTip toolTipAbout;

        public SlideNavigationPane()
        {
            LoadMaterialIconsFont(); // Load font before InitializeComponent
            InitializeComponent();
            this.BackColor = SystemColors.Control;
        }

        private void LoadMaterialIconsFont()
        {
            string fontPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fonts/MaterialSymbolsOutlined.ttf");
            if (File.Exists(fontPath))
            {
                _privateFonts.AddFontFile(fontPath);
                _fontLoaded = true;
            }
        }

        private void InitializeComponent()
        {
            this.btnLeft = new Button();
            this.btnRight = new Button();
            this.btnEnd = new Button();
            this.btnDockBottom = new Button();
            this.btnDockRight = new Button();
            this.btnBackToGrid = new Button();
            this.btnAbout = new Button();
            this.linePanel = new Panel();

            this.lblBackToGrid = new Label();
            this.lblEnd = new Label();

            this.toolTipLeft = new ToolTip();
            this.toolTipRight = new ToolTip();
            this.toolTipBackToGrid = new ToolTip();
            this.toolTipEnd = new ToolTip();
            this.toolTipDockBottom = new ToolTip();
            this.toolTipDockRight = new ToolTip();
            this.toolTipAbout = new ToolTip();

            // 
            // btnLeft
            // 
            if (_fontLoaded)
            {
                this.btnLeft.Font = new Font(_privateFonts.Families[0], 40);
                this.btnLeft.Text = "\ue5de"; // Material Icons: arrow_left
            }
            else
            {
                this.btnLeft.Font = new Font("Courier New", 36F, FontStyle.Bold);
                this.btnLeft.Text = "<";
            }
            this.btnLeft.Width = 120;
            this.btnLeft.Height = 80;
            this.btnLeft.Left = 10;
            this.btnLeft.TextAlign = ContentAlignment.MiddleCenter;
            this.btnLeft.Padding = new Padding(0, 10, 0, 0);
            this.btnLeft.Click += (s, e) => LeftArrowClicked?.Invoke(this, EventArgs.Empty);

            toolTipLeft.SetToolTip(this.btnLeft, "上一頁");
            // 
            // btnRight
            // 
            if (_fontLoaded)
            {
                this.btnRight.Font = new Font(_privateFonts.Families[0], 40);
                this.btnRight.Text = "\ue5df"; // Material Icons: arrow_right
            }
            else
            {
                this.btnRight.Font = new Font("Courier New", 36F, FontStyle.Bold);
                this.btnRight.Text = ">";
            }
            this.btnRight.Width = 120;
            this.btnRight.Height = 80;
            this.btnRight.Left = 140;
            this.btnRight.TextAlign = ContentAlignment.MiddleCenter;
            this.btnRight.Padding = new Padding(0, 10, 0, 0);
            this.btnRight.Click += (s, e) => RightArrowClicked?.Invoke(this, EventArgs.Empty);

            toolTipRight.SetToolTip(this.btnRight, "下一頁");
            // 
            // linePanel
            // 
            this.linePanel.Height = 2;
            this.linePanel.Width = 250;
            this.linePanel.Left = 10;
            this.linePanel.Top = 90;
            this.linePanel.BackColor = Color.LightGray;
            // 
            // btnBackToGrid
            // 
            if (_fontLoaded)
            {
                this.btnBackToGrid.Font = new Font(_privateFonts.Families[0], 40);
                this.btnBackToGrid.Text = "\ue9b0"; // Material Icons: grid_view
            }
            else
            {
                this.btnBackToGrid.Font = new Font("Calibri", 15, FontStyle.Bold);
                this.btnBackToGrid.Text = "Slideshow";
            }
            this.btnBackToGrid.Width = 120;
            this.btnBackToGrid.Height = 80;
            this.btnBackToGrid.Left = 10;
            this.btnBackToGrid.Top = 100;
            this.btnBackToGrid.TextAlign = ContentAlignment.MiddleCenter;
            this.btnBackToGrid.Padding = new Padding(0, 10, 0, 0);
            this.btnBackToGrid.Click += (s, e) => BackToGridClicked?.Invoke(this, EventArgs.Empty);

            toolTipBackToGrid.SetToolTip(this.btnBackToGrid, "返回投影片網格");

            lblBackToGrid.Text = "返回";
            lblBackToGrid.Font = new Font("微軟正黑體, Microsoft JhengHei, Microsoft JhengHei UI, 新細明體, PMingLiU", 14, FontStyle.Bold);
            lblBackToGrid.AutoSize = false;
            lblBackToGrid.TextAlign = ContentAlignment.TopCenter;
            lblBackToGrid.Width = btnBackToGrid.Width;
            lblBackToGrid.Height = 20; // Adjust as needed
            lblBackToGrid.Left = btnBackToGrid.Left;
            lblBackToGrid.Top = btnBackToGrid.Top + btnBackToGrid.Height + 2;
            // 
            // btnEnd
            // 
            if (_fontLoaded)
            {
                this.btnEnd.Font = new Font(_privateFonts.Families[0], 36);
                this.btnEnd.Text = "\uf2f6"; // Material Icons: computer_cancel
            }
            else
            {
                this.btnEnd.Font = new Font("Calibri ", 15, FontStyle.Bold);
                this.btnEnd.Text = "End";
            }
            this.btnEnd.Width = 120;
            this.btnEnd.Height = 80;
            this.btnEnd.Left = 140;
            this.btnEnd.Top = 190;
            this.btnEnd.TextAlign = ContentAlignment.MiddleCenter;
            this.btnEnd.Padding = new Padding(0, 10, 0, 0);
            this.btnEnd.FlatStyle = FlatStyle.Flat;
            this.btnEnd.FlatAppearance.BorderSize = 0;
            this.btnEnd.Click += (s, e) => EndButtonClicked?.Invoke(this, EventArgs.Empty);

            toolTipEnd.SetToolTip(this.btnEnd, "結束放映");

            lblEnd.Text = "結束";
            lblEnd.Font = new Font("微軟正黑體, Microsoft JhengHei, Microsoft JhengHei UI, 新細明體, PMingLiU", 14, FontStyle.Bold);
            lblEnd.AutoSize = false;
            lblEnd.TextAlign = ContentAlignment.TopCenter;
            lblEnd.Width = btnEnd.Width;
            lblEnd.Height = 20; // Adjust as needed
            lblEnd.Left = btnEnd.Left;
            lblEnd.Top = btnEnd.Top + btnEnd.Height + 2;
            // 
            // btnAbout
            // 
            if (_fontLoaded)
            {
                this.btnAbout.Font = new Font(_privateFonts.Families[0], 16);
                this.btnAbout.Text = "\ue88e"; // Material Icons: info
            }
            else
            {
                this.btnAbout.Font = new Font("Calibri ", 15, FontStyle.Bold);
                this.btnAbout.Text = "About me";
            }
            this.btnAbout.Width = 35;
            this.btnAbout.Height = 35;
            this.btnAbout.Left = btnEnd.Left + btnEnd.Width - this.btnAbout.Width;
            this.btnAbout.Top = btnEnd.Bottom + 100;
            this.btnAbout.TextAlign = ContentAlignment.MiddleCenter;
            this.btnAbout.FlatStyle = FlatStyle.Flat;
            this.btnAbout.FlatAppearance.BorderSize = 0;
            this.btnAbout.Click += (s, e) => btnAbout_Click?.Invoke(this, EventArgs.Empty);

            toolTipAbout.SetToolTip(this.btnAbout, "關於插件");
            // 
            // btnDockBottom
            // 
            if (_fontLoaded)
            {
                this.btnDockBottom.Font = new Font(_privateFonts.Families[0], 20);
                this.btnDockBottom.Text = "\uf72a"; // Material Icons: bottom_panel_close
            }
            else
            {
                this.btnDockBottom.Font = new Font("Calibri ", 8);
                this.btnDockBottom.Text = "Dock to bottom";
            }
            this.btnDockBottom.Width = 280;
            this.btnDockBottom.Height = 60;
            this.btnDockBottom.Left = 0;
            this.btnDockBottom.Top = this.Height - this.btnDockBottom.Height;
            this.btnDockBottom.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.btnDockBottom.TextAlign = ContentAlignment.MiddleCenter;
            this.btnDockBottom.Padding = new Padding(0, 5, 0, 0);
            this.btnDockBottom.FlatStyle = FlatStyle.Flat;
            this.btnDockBottom.FlatAppearance.BorderSize = 0;
            this.btnDockBottom.Click += (s, e) => DockToBottomClicked?.Invoke(this, EventArgs.Empty);

            toolTipDockBottom.SetToolTip(this.btnDockBottom, "停靠底部");
            // 
            // btnDockRight
            // 
            if (_fontLoaded)
            {
                this.btnDockRight.Font = new Font(_privateFonts.Families[0], 20);
                this.btnDockRight.Text = "\uf705"; // Material Icons: right_panel_close
            }
            else
            {
                this.btnDockRight.Font = new Font("Calibri ", 8);
                this.btnDockRight.Text = "Dock to right";
            }
            this.btnDockRight.Width = 60; // Swapped width and height
            this.btnDockRight.Height = 103;
            this.btnDockRight.Top = 25;
            // Place at the right side of the bottom pane
            this.btnDockRight.Left = this.Width - this.btnDockRight.Width;
            this.btnDockRight.Anchor = AnchorStyles.Right;
            this.btnDockRight.TextAlign = ContentAlignment.MiddleCenter;
            this.btnDockRight.Padding = new Padding(0, 5, 0, 0);
            this.btnDockRight.FlatStyle = FlatStyle.Flat;
            this.btnDockRight.FlatAppearance.BorderSize = 0;
            this.btnDockRight.Visible = false;
            this.btnDockRight.Click += (s, e) => DockToRightClicked?.Invoke(this, EventArgs.Empty);

            toolTipDockRight.SetToolTip(this.btnDockRight, "停靠右側");
            // 
            // SlideNavigationPane
            // 
            this.Controls.Add(this.btnLeft);
            this.Controls.Add(this.btnRight);
            this.Controls.Add(this.linePanel);
            this.Controls.Add(this.btnBackToGrid);
            this.Controls.Add(this.lblBackToGrid);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.lblEnd);
            this.Controls.Add(this.btnDockBottom);
            this.Controls.Add(this.btnDockRight);
            this.Controls.Add(this.btnAbout);
            this.Width = 280;
            this.Height = _fontLoaded ? 300 : 280;
            // Ensure btnDockBottom and btnDockRight stay at the bottom on resize
            this.Resize += (s, e) =>
            {
                this.btnDockBottom.Top = this.Height - this.btnDockBottom.Height;
                this.btnDockRight.Left = this.Width - this.btnDockRight.Width;
            };
        }

        public void UpdateButtonLayoutForDock(bool isDockedBottom)
        {
            if (isDockedBottom)
            {
                btnLeft.Left = 50;
                btnRight.Left = btnLeft.Right + 10;

                // Place Back to Grid button to the right of the arrows
                btnBackToGrid.Left = btnRight.Right + 100;
                btnBackToGrid.Top = btnRight.Top;
                lblBackToGrid.Visible = false;

                // Place End button to the Back to Grid button
                btnEnd.Left = btnBackToGrid.Right + 100;
                btnEnd.Top = btnRight.Top;
                lblEnd.Visible = false;

                linePanel.Visible = false;

                btnAbout.Left = btnEnd.Left + btnEnd.Width + 100;
                btnAbout.Top = btnEnd.Top;

                btnDockBottom.Visible = false;
                btnDockRight.Visible = true;
            }
            else
            {
                // Restore original position
                btnLeft.Left = 10;
                btnRight.Left = btnLeft.Right + 10;

                btnBackToGrid.Left = 10;
                btnBackToGrid.Top = 100;
                lblBackToGrid.Visible = true;

                btnEnd.Left = 140;
                btnEnd.Top = 190;
                lblEnd.Visible = true;

                linePanel.Visible = true;

                btnAbout.Left = btnEnd.Left + btnEnd.Width - btnAbout.Width;
                btnAbout.Top = btnEnd.Bottom + 100;

                btnDockBottom.Visible = true;
                btnDockRight.Visible = false;
            }
        }
    }
}
