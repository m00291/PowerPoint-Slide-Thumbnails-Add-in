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

        private PrivateFontCollection _privateFonts = new PrivateFontCollection();
        private bool _fontLoaded = false;
        private Button btnLeft;
        private Button btnRight;
        private Button btnEnd;
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
            this.btnLeft.Top = 10;
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
            this.btnRight.Top = 10;
            this.btnRight.TextAlign = ContentAlignment.MiddleCenter;
            this.btnRight.Click += (s, e) => RightArrowClicked?.Invoke(this, EventArgs.Empty);
            // 
            // linePanel
            // 
            this.linePanel.Height = 2;
            this.linePanel.Width = 250;
            this.linePanel.Left = 10;
            this.linePanel.Top = 100;
            this.linePanel.BackColor = Color.LightGray;
            // 
            // btnEnd
            // 
            if (_fontLoaded)
            {
                this.btnEnd.Font = new Font(_privateFonts.Families[0], 36, FontStyle.Regular, GraphicsUnit.Point);
                this.btnEnd.Text = "\uf2f6"; // Material Icons: cancel_presentation
            }
            else
            {
                this.btnEnd.Font = new Font("Calibri ", 20F, FontStyle.Bold);
                this.btnEnd.Text = "End";
            }
            this.btnEnd.Width = 120;
            this.btnEnd.Height = 80;
            this.btnEnd.Left = 140;
            this.btnEnd.Top = 110;
            this.btnEnd.TextAlign = ContentAlignment.MiddleCenter;
            this.btnEnd.FlatStyle = FlatStyle.Flat;
            this.btnEnd.FlatAppearance.BorderSize = 0;
            this.btnEnd.Click += (s, e) => EndButtonClicked?.Invoke(this, EventArgs.Empty);
            // 
            // SlideNavigationPane
            // 
            this.Controls.Add(this.btnLeft);
            this.Controls.Add(this.btnRight);
            this.Controls.Add(this.linePanel);
            this.Controls.Add(this.btnEnd);
            this.Width = 240;
            this.Height = _fontLoaded ? 230 : 210;
        }
    }
}
