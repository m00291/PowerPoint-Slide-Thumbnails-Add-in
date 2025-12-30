using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointSlideThumbnailsAddIn
{
    public partial class ThisAddIn
    {
        private PowerPoint.Application pptApp;
        private PowerPoint.Presentation currentPresentation;

        private Microsoft.Office.Tools.CustomTaskPane navigationTaskPane;
        private SlideNavigationPane navigationPaneControl;
        private bool isSyncingSelection = false;
        public int currentSlideIndex = 1;
        public int originalZoom = 100;

        private const int TaskPaneBottomHeight = 130; // Height for bottom dock
        private const int TaskPaneRightWidth = 270;   // Width for right dock

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            pptApp = this.Application;
            pptApp.SlideShowBegin += PptApp_SlideShowBegin;
            pptApp.SlideShowEnd += PptApp_SlideShowEnd;
            pptApp.WindowSelectionChange += PptApp_WindowSelectionChange;
            pptApp.SlideShowNextSlide += PptApp_SlideShowNextSlide;
        }

        private void PptApp_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            try
            {
                CollapseRibbon();

                currentPresentation = Wn.Presentation;
                bool presenterViewWasOn = false;
                try
                {
                    // Check if Presenter View is enabled
                    if (currentPresentation.SlideShowSettings.ShowPresenterView == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        presenterViewWasOn = true;
                        currentPresentation.SlideShowSettings.ShowPresenterView = Microsoft.Office.Core.MsoTriState.msoFalse;
                    }
                }
                catch { }

                if (presenterViewWasOn)
                {
                    // End current slideshow and restart with Presenter View disabled
                    try { Wn.View.Exit(); } catch { }
                    try { currentPresentation.SlideShowSettings.Run(); } catch { }
                    return;
                }

                var window = currentPresentation.Windows[1];
                window.ViewType = PowerPoint.PpViewType.ppViewSlideSorter;
                originalZoom = window.View.Zoom;
                window.View.Zoom = 130;

                // Show navigation task pane
                if (navigationPaneControl == null)
                {
                    navigationPaneControl = new SlideNavigationPane();
                    navigationPaneControl.LeftArrowClicked += NavigationPaneControl_LeftArrowClicked;
                    navigationPaneControl.RightArrowClicked += NavigationPaneControl_RightArrowClicked;
                    navigationPaneControl.BackToGridClicked += NavigationPaneControl_BackToGridClicked;
                    navigationPaneControl.EndButtonClicked += NavigationPaneControl_EndButtonClicked;
                    navigationPaneControl.DockToBottomClicked += NavigationPaneControl_DockToBottomClicked;
                    navigationPaneControl.DockToRightClicked += NavigationPaneControl_DockToRightClicked;
                    navigationPaneControl.btnAbout_Click += NavigationPaneControl_About_Click;
                }
                if (navigationTaskPane == null)
                {
                    navigationTaskPane = this.CustomTaskPanes.Add(navigationPaneControl, " ");
                    //navigationTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                    //navigationTaskPane.Width = TaskPaneRightWidth;
                    navigationTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
                    navigationTaskPane.Height = TaskPaneBottomHeight;
                }
                navigationTaskPane.Visible = true;

                navigationTaskPane.VisibleChanged += (s, e) =>
                {
                    if (pptApp.SlideShowWindows.Count > 0)
                    {
                        try
                        {
                            new Thread(() =>
                            {
                                navigationTaskPane.Visible = true;
                            }).Start();
                        }
                        catch { }
                    }
                };

                // Show/hide DockToBottom button based on dock position
                bool isBottom = navigationTaskPane.DockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
                navigationPaneControl.UpdateButtonLayoutForDock(isBottom);
            }
            catch { }
        }

        private void CollapseRibbon()
        {
            try
            {
                if (pptApp.CommandBars["Ribbon"] != null)
                {
                    float height = pptApp.CommandBars["Ribbon"].Controls[1].Height;
                    if (height >= 100)
                    {
                        pptApp.CommandBars.ExecuteMso("MinimizeRibbon");
                    }
                }
            }
            catch { }
        }

        private void UncollapseRibbon()
        {
            try
            {
                if (pptApp.CommandBars["Ribbon"] != null)
                {
                    float height = pptApp.CommandBars["Ribbon"].Controls[1].Height;
                    if (height < 100)
                    {
                        pptApp.CommandBars.ExecuteMso("MinimizeRibbon");
                    }
                }
            }
            catch { }
        }

        private void NavigationPaneControl_BackToGridClicked(object sender, EventArgs e)
        {
            try
            {
                CollapseRibbon();

                isSyncingSelection = true;
                var window = currentPresentation.Windows[1];
                window.ViewType = PowerPoint.PpViewType.ppViewSlideSorter;

                if (pptApp.SlideShowWindows.Count > 0)
                {
                    window.View.GotoSlide(currentSlideIndex);
                    window.Selection.SlideRange.Select();
                }
                isSyncingSelection = false;
            }
            catch { isSyncingSelection = false; }
        }

        private void NavigationPaneControl_EndButtonClicked(object sender, EventArgs e)
        {
            try
            {
                if (pptApp.SlideShowWindows.Count > 0)
                {
                    pptApp.SlideShowWindows[1].View.Exit();
                }
            }
            catch { }
        }

        private void NavigationPaneControl_About_Click(object sender, EventArgs e)
        {
            AboutBox aboutBox = new AboutBox();
            aboutBox.ShowDialog();
        }

        private void NavigationPaneControl_DockToBottomClicked(object sender, EventArgs e)
        {
            if (navigationTaskPane != null)
            {
                navigationTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
                navigationPaneControl.UpdateButtonLayoutForDock(true);
                navigationTaskPane.Height = TaskPaneBottomHeight;
            }
        }

        private void NavigationPaneControl_DockToRightClicked(object sender, EventArgs e)
        {
            if (navigationTaskPane != null)
            {
                navigationTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                navigationPaneControl.UpdateButtonLayoutForDock(false);
                navigationTaskPane.Width = TaskPaneRightWidth;
            }
        }

        private void PptApp_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            try
            {
                UncollapseRibbon();

                if (currentPresentation != null)
                {
                    var window = currentPresentation.Windows[1];
                    window.View.Zoom = originalZoom;
                    window.ViewType = PowerPoint.PpViewType.ppViewNormal;
                }
                // Hide and remove navigation task pane
                if (navigationTaskPane != null)
                {
                    navigationTaskPane.Visible = false;
                    this.CustomTaskPanes.Remove(navigationTaskPane);
                    navigationTaskPane = null;
                    navigationPaneControl = null; // Ensure new control is created next time
                }
            }
            catch { }
        }

        private void NavigationPaneControl_LeftArrowClicked(object sender, EventArgs e)
        {
            try
            {
                if (pptApp.SlideShowWindows.Count > 0)
                {
                    var view = pptApp.SlideShowWindows[1].View;
                    view.Previous();
                    currentSlideIndex = view.CurrentShowPosition;
                }
            }
            catch { }
        }

        private void NavigationPaneControl_RightArrowClicked(object sender, EventArgs e)
        {
            try
            {
                if (pptApp.SlideShowWindows.Count > 0)
                {
                    var view = pptApp.SlideShowWindows[1].View;
                    view.Next();
                    currentSlideIndex = view.CurrentShowPosition;
                }
            }
            catch { }
        }

        private void PptApp_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            // This event is triggered when the next slide is shown in the slideshow
            try
            {
                if (currentPresentation != null && currentPresentation.Windows.Count > 0)
                {
                    var window = currentPresentation.Windows[1];
                    if (window.ViewType == PowerPoint.PpViewType.ppViewSlideSorter)
                    {
                        int slideIndex = Wn.View.CurrentShowPosition;
                        isSyncingSelection = true;
                        window.View.GotoSlide(slideIndex);
                        window.Selection.SlideRange.Select();
                        isSyncingSelection = false;
                    }
                }
            }
            catch { isSyncingSelection = false; }
        }

        private void PptApp_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            // This event is triggered when the selection changes in the SlideSorter view
            try
            {
                if (isSyncingSelection) return;
                // Only act if in Slide Sorter view and a slide is selected
                if (currentPresentation != null && currentPresentation.Windows[1].ViewType == PowerPoint.PpViewType.ppViewSlideSorter)
                {
                    if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides && Sel.SlideRange != null && Sel.SlideRange.Count > 0)
                    {
                        var slideIndex = Sel.SlideRange[1].SlideIndex;
                        // Find the running slideshow window
                        if (pptApp.SlideShowWindows.Count > 0)
                        {
                            var slideShowView = pptApp.SlideShowWindows[1].View;
                            slideShowView.GotoSlide(slideIndex);
                            currentSlideIndex = slideIndex;
                        }
                    }
                }
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Unsubscribe from PowerPoint events
            if (pptApp != null)
            {
                pptApp.SlideShowBegin -= PptApp_SlideShowBegin;
                pptApp.SlideShowEnd -= PptApp_SlideShowEnd;
                pptApp.WindowSelectionChange -= PptApp_WindowSelectionChange;
                pptApp.SlideShowNextSlide -= PptApp_SlideShowNextSlide;
            }

            // Remove and dispose of navigation task pane and control
            if (navigationTaskPane != null)
            {
                try
                {
                    navigationTaskPane.Visible = false;
                    this.CustomTaskPanes.Remove(navigationTaskPane);
                }
                catch { }
                navigationTaskPane = null;
            }
            if (navigationPaneControl != null)
            {
                navigationPaneControl.Dispose();
                navigationPaneControl = null;
            }
            currentPresentation = null;
            pptApp = null;
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

}
