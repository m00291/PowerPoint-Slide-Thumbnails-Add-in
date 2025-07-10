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
        public event EventHandler EndButtonClicked;
        public event EventHandler DockToBottomClicked;
        public event EventHandler DockToRightClicked;

        private PrivateFontCollection _privateFonts = new PrivateFontCollection();
        private bool _fontLoaded = false;
        private Button btnLeft;
        private Button btnRight;
        private Button btnEnd;
        private Button btnDockBottom;
        private Button btnDockRight;
        private Button btnBackToGrid;
        private Panel linePanel;

        public SlideNavigationPane()
        {
            LoadMaterialIconsFont(); // Load font before InitializeComponent
            InitializeComponent();
            this.BackColor = SystemColors.Control;
        }

        private void LoadMaterialIconsFont()
        {
            string fontPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fonts", "MaterialSymbolsOutlined.ttf");
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
            this.linePanel = new Panel();
            // 
            // btnLeft
            // 
            if (_fontLoaded)
            {
                this.btnLeft.Font = new Font(_privateFonts.Families[0], 40, FontStyle.Regular, GraphicsUnit.Point);
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
            this.btnLeft.Click += (s, e) => LeftArrowClicked?.Invoke(this, EventArgs.Empty);
            // 
            // btnRight
            // 
            if (_fontLoaded)
            {
                this.btnRight.Font = new Font(_privateFonts.Families[0], 40, FontStyle.Regular, GraphicsUnit.Point);
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
            this.btnRight.Click += (s, e) => RightArrowClicked?.Invoke(this, EventArgs.Empty);
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
                this.btnBackToGrid.Font = new Font(_privateFonts.Families[0], 40, FontStyle.Regular, GraphicsUnit.Point);
                this.btnBackToGrid.Text = "\ue9b0"; // Material Icons: grid_view
            }
            else
            {
                this.btnBackToGrid.Font = new Font("Calibri", 24F, FontStyle.Bold);
                this.btnBackToGrid.Text = "Grid";
            }
            this.btnBackToGrid.Width = 120;
            this.btnBackToGrid.Height = 80;
            this.btnBackToGrid.Left = 10;
            this.btnBackToGrid.Top = 100;
            this.btnBackToGrid.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // btnEnd
            // 
            if (_fontLoaded)
            {
                this.btnEnd.Font = new Font(_privateFonts.Families[0], 36, FontStyle.Regular, GraphicsUnit.Point);
                this.btnEnd.Text = "\uf2f6"; // Material Icons: computer_cancel
            }
            else
            {
                this.btnEnd.Font = new Font("Calibri ", 20F, FontStyle.Bold);
                this.btnEnd.Text = "End";
            }
            this.btnEnd.Width = 120;
            this.btnEnd.Height = 80;
            this.btnEnd.Left = 140;
            this.btnEnd.Top = 190;
            this.btnEnd.TextAlign = ContentAlignment.MiddleCenter;
            this.btnEnd.FlatStyle = FlatStyle.Flat;
            this.btnEnd.FlatAppearance.BorderSize = 0;
            this.btnEnd.Click += (s, e) => EndButtonClicked?.Invoke(this, EventArgs.Empty);
            // 
            // btnDockBottom
            // 
            this.btnDockBottom.Font = new Font(_privateFonts.Families[0], 20, FontStyle.Regular, GraphicsUnit.Point);
            this.btnDockBottom.Text = "\uf72a"; // Material Icons: bottom_panel_close
            this.btnDockBottom.Width = 280;
            this.btnDockBottom.Height = 60;
            this.btnDockBottom.Left = 0;
            this.btnDockBottom.Top = this.Height - this.btnDockBottom.Height;
            this.btnDockBottom.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.btnDockBottom.TextAlign = ContentAlignment.MiddleCenter;
            this.btnDockBottom.FlatStyle = FlatStyle.Flat;
            this.btnDockBottom.FlatAppearance.BorderSize = 0;
            this.btnDockBottom.Click += (s, e) => DockToBottomClicked?.Invoke(this, EventArgs.Empty);
            // 
            // btnDockRight
            // 
            this.btnDockRight.Font = new Font(_privateFonts.Families[0], 20, FontStyle.Regular, GraphicsUnit.Point);
            this.btnDockRight.Text = "\uf705"; // Material Icons: right_panel_close
            this.btnDockRight.Width = 60; // Swapped width and height
            this.btnDockRight.Height = 103;
            this.btnDockRight.Top = 25;
            // Place at the right side of the bottom pane
            this.btnDockRight.Left = this.Width - this.btnDockRight.Width;
            this.btnDockRight.Anchor = AnchorStyles.Right;
            this.btnDockRight.TextAlign = ContentAlignment.MiddleCenter;
            this.btnDockRight.FlatStyle = FlatStyle.Flat;
            this.btnDockRight.FlatAppearance.BorderSize = 0;
            this.btnDockRight.Visible = false;
            this.btnDockRight.Click += (s, e) => DockToRightClicked?.Invoke(this, EventArgs.Empty);
            // 
            // SlideNavigationPane
            // 
            this.Controls.Add(this.btnLeft);
            this.Controls.Add(this.btnRight);
            this.Controls.Add(this.linePanel);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.btnDockBottom);
            this.Controls.Add(this.btnDockRight);
            this.Controls.Add(this.btnBackToGrid);
            this.Width = 240;
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

                // Place End button to the right of the arrows
                btnEnd.Left = btnBackToGrid.Right + 100;
                btnEnd.Top = btnRight.Top;

                linePanel.Visible = false;

                btnDockBottom.Visible = false;
                btnDockRight.Visible = true;
            }
            else
            {
                btnLeft.Left = 10;
                btnRight.Left = btnLeft.Right + 10;

                // Restore original position
                btnBackToGrid.Left = 10;
                btnBackToGrid.Top = 100;

                btnEnd.Left = 140;
                btnEnd.Top = 190;

                linePanel.Visible = true;

                btnDockBottom.Visible = true;
                btnDockRight.Visible = false;
            }
        }
    }
}
