using System;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointTimer.Model;
using PowerPointTimer.Util;

namespace PowerPointTimer
{
    public partial class ThisAddIn
    {
        private TimerData _activeTimerData;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.SlideShowEnd += OnSlideShowEnd;
            this.Application.SlideShowNextSlide += OnNextSlide;
        }

        private void OnNextSlide(SlideShowWindow Wn)
        {
            OnSlideExiting();
            OnSlideEntering(Wn.View.Slide);
        }

        private void OnSlideEntering(Slide slide)
        {
            var timerShape = FindTimerOnSlide(slide);
            if (timerShape == null)
                return; // no timer on this slide

            _activeTimerData = new TimerData(timerShape);
        }

        private void OnSlideExiting()
        {
            _activeTimerData?.Dispose();
            _activeTimerData = null;
        }

        private void OnSlideShowEnd(Presentation Pres)
        {
            _activeTimerData?.Dispose();
            _activeTimerData = null;
        }

        private Shape FindTimerOnSlide(Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Tags[Constants.TimerTagName] == Constants.DigitalTimerTagValue)
                {
                    return shape;
                }
            }
            return null;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
