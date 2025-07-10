using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

                // Show navigation task pane
                if (navigationPaneControl == null)
                {
                    navigationPaneControl = new SlideNavigationPane();
                    navigationPaneControl.LeftArrowClicked += NavigationPaneControl_LeftArrowClicked;
                    navigationPaneControl.RightArrowClicked += NavigationPaneControl_RightArrowClicked;
                    navigationPaneControl.EndButtonClicked += NavigationPaneControl_EndButtonClicked;
                    navigationPaneControl.DockToBottomClicked += NavigationPaneControl_DockToBottomClicked;
                }
                if (navigationTaskPane == null)
                {
                    navigationTaskPane = this.CustomTaskPanes.Add(navigationPaneControl, " ");
                    navigationTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                    navigationTaskPane.Width = 280;
                }
                navigationTaskPane.Visible = true;
                // Show/hide DockToBottom button based on dock position
                bool isBottom = navigationTaskPane.DockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
                navigationPaneControl.SetDockToBottomButtonVisible(!isBottom);
                navigationPaneControl.UpdateEndButtonLayoutForDock(isBottom);
            }
            catch { }
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

        private void NavigationPaneControl_DockToBottomClicked(object sender, EventArgs e)
        {
            if (navigationTaskPane != null)
            {
                navigationTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom;
                navigationPaneControl.SetDockToBottomButtonVisible(false);
                navigationPaneControl.UpdateEndButtonLayoutForDock(true);
            }
        }

        private void PptApp_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            try
            {
                if (currentPresentation != null)
                {
                    var window = currentPresentation.Windows[1];
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
                }
            }
            catch { }
        }

        private void PptApp_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
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
                        }
                    }
                }
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
