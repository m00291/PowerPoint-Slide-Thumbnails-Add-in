using System;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
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

        private Label debugmsg;
        private bool _isDockedBottom;
        private int taskpaneleftpadding = 10;
        private int taskpanerightpadding = 4;
        private int taskpanebottompadding = 10;
        private int btngap = 10;
        private int bottom_btngap = 100;
        private int btnwidth = 120;
        private int btnheight = 80;
        private int lblheight = 20;
        private int btnAboutWidth = 35;
        private int btnAboutHeight = 35;
        private int btnDockBottomWidth = 280;
        private int btnDockBottomHeight = 60;
        private int btnDockRightWidth = 60;
        private int btnDockRightHeight = 103;

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
            var culture = Thread.CurrentThread.CurrentUICulture;
            //Thread.CurrentThread.CurrentUICulture = new CultureInfo("zh-TW"); // test with Traditional Chinese

            LoadMaterialIconsFont(); // Load font before InitializeComponent
            InitializeComponent();
            this.BackColor = SystemColors.Control;
            this.SizeChanged += TaskPaneChangeSize;
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

            this.debugmsg = new Label();

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
            this.btnLeft.Width = btnwidth;
            this.btnLeft.Height = btnheight;
            this.btnLeft.Left = taskpaneleftpadding;
            this.btnLeft.TextAlign = ContentAlignment.MiddleCenter;
            this.btnLeft.Padding = new Padding(0, 10, 0, 0);
            this.btnLeft.Click += (s, e) => LeftArrowClicked?.Invoke(this, EventArgs.Empty);

            toolTipLeft.SetToolTip(this.btnLeft, Properties.Strings.toolTipLeft);
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
            this.btnRight.Width = btnwidth;
            this.btnRight.Height = btnheight;
            this.btnRight.Left = btnLeft.Right + btngap;
            this.btnRight.TextAlign = ContentAlignment.MiddleCenter;
            this.btnRight.Padding = new Padding(0, 10, 0, 0);
            this.btnRight.Click += (s, e) => RightArrowClicked?.Invoke(this, EventArgs.Empty);

            toolTipRight.SetToolTip(this.btnRight, Properties.Strings.toolTipRight);
            // 
            // linePanel
            // 
            this.linePanel.Height = 2;
            this.linePanel.Width = this.Width - taskpaneleftpadding - taskpanerightpadding;
            this.linePanel.Left = taskpaneleftpadding;
            this.linePanel.Top = btnLeft.Bottom + btngap;
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
            this.btnBackToGrid.Width = btnwidth;
            this.btnBackToGrid.Height = btnheight;
            this.btnBackToGrid.Left = taskpaneleftpadding;
            this.btnBackToGrid.Top = linePanel.Bottom + btngap;
            this.btnBackToGrid.TextAlign = ContentAlignment.MiddleCenter;
            this.btnBackToGrid.Padding = new Padding(0, 10, 0, 0);
            this.btnBackToGrid.Click += (s, e) => BackToGridClicked?.Invoke(this, EventArgs.Empty);

            toolTipBackToGrid.SetToolTip(this.btnBackToGrid, Properties.Strings.toolTipBackToGrid);

            lblBackToGrid.Text = Properties.Strings.lblBackToGrid;
            lblBackToGrid.Font = new Font("微軟正黑體, Microsoft JhengHei, Microsoft JhengHei UI, 新細明體, PMingLiU", 14, FontStyle.Bold);
            lblBackToGrid.AutoSize = false;
            lblBackToGrid.TextAlign = ContentAlignment.TopCenter;
            lblBackToGrid.Width = btnBackToGrid.Width;
            lblBackToGrid.Height = lblheight;
            lblBackToGrid.Left = btnBackToGrid.Left;
            lblBackToGrid.Top = btnBackToGrid.Bottom + 2;
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
            this.btnEnd.Width = btnwidth;
            this.btnEnd.Height = btnheight;
            this.btnEnd.Left = btnBackToGrid.Right + btngap;
            this.btnEnd.Top = btnBackToGrid.Bottom + btngap;
            this.btnEnd.TextAlign = ContentAlignment.MiddleCenter;
            this.btnEnd.Padding = new Padding(0, 10, 0, 0);
            //this.btnEnd.FlatStyle = FlatStyle.Flat;
            //this.btnEnd.FlatAppearance.BorderSize = 0;
            this.btnEnd.Click += (s, e) => EndButtonClicked?.Invoke(this, EventArgs.Empty);

            toolTipEnd.SetToolTip(this.btnEnd, Properties.Strings.toolTipEnd);

            lblEnd.Text = Properties.Strings.lblEnd;
            lblEnd.Font = new Font("微軟正黑體, Microsoft JhengHei, Microsoft JhengHei UI, 新細明體, PMingLiU", 14, FontStyle.Bold);
            lblEnd.AutoSize = false;
            lblEnd.TextAlign = ContentAlignment.TopCenter;
            lblEnd.Width = btnEnd.Width;
            lblEnd.Height = lblheight;
            lblEnd.Left = btnEnd.Left;
            lblEnd.Top = btnEnd.Bottom + 2;
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
            this.btnAbout.Width = btnAboutWidth;
            this.btnAbout.Height = btnAboutHeight;
            this.btnAbout.Left = this.Width - taskpanerightpadding - btnAboutWidth;
            this.btnAbout.Top = btnEnd.Bottom + 100;
            this.btnAbout.TextAlign = ContentAlignment.MiddleCenter;
            //this.btnAbout.FlatStyle = FlatStyle.Flat;
            //this.btnAbout.FlatAppearance.BorderSize = 0;
            this.btnAbout.Click += (s, e) => btnAbout_Click?.Invoke(this, EventArgs.Empty);

            toolTipAbout.SetToolTip(this.btnAbout, Properties.Strings.toolTipAbout);

            //
            // debugmsg
            this.debugmsg.Text = "";
            this.debugmsg.Font = new Font("Calibri", 10, FontStyle.Bold);
            this.debugmsg.AutoSize = true;
            this.debugmsg.Left = 10;
            this.debugmsg.Top = btnAbout.Top + btnAbout.Height + 5;

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
            this.btnDockBottom.Width = btnDockBottomWidth;
            this.btnDockBottom.Height = btnDockBottomHeight;
            this.btnDockBottom.Left = 0;
            this.btnDockBottom.Top = this.Height - this.btnDockBottom.Height;
            this.btnDockBottom.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.btnDockBottom.TextAlign = ContentAlignment.MiddleCenter;
            this.btnDockBottom.Padding = new Padding(0, 5, 0, 0);
            //this.btnDockBottom.FlatStyle = FlatStyle.Flat;
            //this.btnDockBottom.FlatAppearance.BorderSize = 0;
            this.btnDockBottom.Click += (s, e) => DockToBottomClicked?.Invoke(this, EventArgs.Empty);

            toolTipDockBottom.SetToolTip(this.btnDockBottom, Properties.Strings.toolTipDockBottom);
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
            this.btnDockRight.Width = btnDockRightWidth;
            this.btnDockRight.Height = btnDockRightHeight;
            this.btnDockRight.Top = this.Height - this.btnDockRight.Height;
            this.btnDockRight.Left = this.Width - this.btnDockRight.Width;
            this.btnDockRight.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.btnDockRight.TextAlign = ContentAlignment.MiddleCenter;
            this.btnDockRight.Padding = new Padding(0, 5, 0, 0);
            //this.btnDockRight.FlatStyle = FlatStyle.Flat;
            //this.btnDockRight.FlatAppearance.BorderSize = 0;
            this.btnDockRight.Visible = false;
            this.btnDockRight.Click += (s, e) => DockToRightClicked?.Invoke(this, EventArgs.Empty);

            toolTipDockRight.SetToolTip(this.btnDockRight, Properties.Strings.toolTipDockRight);
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
            //this.Controls.Add(this.debugmsg);
            btnDockBottom.BringToFront();
            btnDockRight.BringToFront();
            // Ensure btnDockBottom and btnDockRight stay at the bottom on resize
            this.Resize += (s, e) =>
            {
                this.btnDockBottom.Top = this.Height - this.btnDockBottom.Height;
                this.btnDockRight.Left = this.Width - this.btnDockRight.Width;
            };
        }

        public void TaskPaneChangeSize(object sender, EventArgs e)
        {
            btnwidth = 120;
            btnheight = 80;
            btnDockBottomWidth = 280;
            btnDockBottomHeight = 60;
            btnDockRightWidth = 60;
            btnDockRightHeight = 100;

            if (_isDockedBottom)
            {
                btnheight = this.Height - taskpanebottompadding;

                btnLeft.Left = 50;
                btnRight.Left = btnLeft.Right + btngap;

                // Place Back to Grid button to the right of the arrows
                btnBackToGrid.Left = btnRight.Right + bottom_btngap;
                btnBackToGrid.Top = btnRight.Top;
                lblBackToGrid.Visible = false;

                btnEnd.Left = btnBackToGrid.Right + bottom_btngap;
                btnEnd.Top = btnRight.Top;
                lblEnd.Visible = false;

                linePanel.Visible = false;

                btnAbout.Left = btnEnd.Right + bottom_btngap;
                btnAbout.Top = btnEnd.Top;

                debugmsg.Left = btnAbout.Right + 5;
                debugmsg.Top = btnAbout.Top;

                btnDockRight.Height = this.Height;
                btnDockRight.Top = this.Height - this.btnDockRight.Height;

                btnDockBottom.Visible = false;
                btnDockRight.Visible = true;
            }
            else
            {
                btnwidth = (this.Width - taskpaneleftpadding - taskpanerightpadding - btngap) / 2;

                btnLeft.Left = taskpaneleftpadding;
                btnRight.Left = btnLeft.Right + btngap;

                btnBackToGrid.Left = taskpaneleftpadding;
                btnBackToGrid.Top = linePanel.Bottom + btngap;
                lblBackToGrid.Visible = true;
                AutoShrinkFont(lblBackToGrid);

                btnEnd.Left = btnBackToGrid.Right + btngap;
                btnEnd.Top = btnBackToGrid.Bottom + btngap;
                lblEnd.Visible = true;
                AutoShrinkFont(lblEnd);

                linePanel.Visible = true;

                btnAbout.Left = this.Width - taskpanerightpadding - btnAboutWidth;
                btnAbout.Top = btnEnd.Bottom + 100;

                debugmsg.Left = 10;
                debugmsg.Top = btnAbout.Bottom + 5;

                btnDockBottom.Width = this.Width;

                btnDockBottom.Visible = true;
                btnDockRight.Visible = false;
            }

            btnLeft.Width = btnwidth;
            btnLeft.Height = btnheight;
            btnRight.Width = btnwidth;
            btnRight.Height = btnheight;

            linePanel.Width = this.Width - taskpaneleftpadding - taskpanerightpadding;

            btnBackToGrid.Width = btnwidth;
            btnBackToGrid.Height = btnheight;
            lblBackToGrid.Width = btnwidth;

            btnEnd.Width = btnwidth;
            btnEnd.Height = btnheight;
            lblEnd.Width = btnwidth;

            if (_isDockedBottom)
            {
            }
            else
            {
                btnRight.Left = btnLeft.Right + btngap;
                btnEnd.Left = btnBackToGrid.Right + btngap;
                lblEnd.Left = btnEnd.Left;
            }

            // debug msg
            /*
            debugmsg.Text = "Width: " + this.Width + " | Height: " + this.Height;
            debugmsg.Visible = false;
            */
            // debug msg end
        }

        private void AutoShrinkFont(Label lbl)
        {
            if (string.IsNullOrEmpty(lbl.Text))
                return;

            int minFontSize = 4; // Minimum font size
            int maxFontSize = 14; // Maximum font size

            using (Graphics g = lbl.CreateGraphics())
            {
                for (int size = maxFontSize; size >= minFontSize; size--)
                {
                    Font testFont = new Font(lbl.Font.FontFamily, size, lbl.Font.Style);
                    SizeF textSize = g.MeasureString(lbl.Text, testFont);

                    if (textSize.Width <= lbl.Width)
                    {
                        lbl.Font = testFont;
                        break;
                    }
                }
            }
        }

        public void UpdateButtonLayoutForDock(bool isDockedBottom)
        {
            if (isDockedBottom)
            {
                _isDockedBottom = true;
            }
            else
            {
                _isDockedBottom = false;
            }
            TaskPaneChangeSize(this, EventArgs.Empty);
        }
    }
}
