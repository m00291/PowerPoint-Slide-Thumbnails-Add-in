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

        private PrivateFontCollection _privateFonts = new PrivateFontCollection();
        private bool _fontLoaded = false;

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
            this.btnRight.Left = 150;
            this.btnRight.Top = 10;
            this.btnRight.Click += (s, e) => RightArrowClicked?.Invoke(this, EventArgs.Empty);
            // 
            // SlideNavigationPane
            // 
            this.Controls.Add(this.btnLeft);
            this.Controls.Add(this.btnRight);
            this.Width = 290;
            this.Height = 110; // Reduce height to fit in 320px task pane
        }

        private Button btnLeft;
        private Button btnRight;
    }
}
