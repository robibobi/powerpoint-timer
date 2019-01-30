using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointTimer.Model;
using PowerPointTimer.Util;

namespace PowerPointTimer
{
    public partial class ThisAddIn
    {
        private List<TimerData> _activeTimers = new List<TimerData>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.SlideShowEnd += OnSlideShowEnd;
            this.Application.SlideShowNextSlide += OnNextSlide;
        }

        private void OnNextSlide(SlideShowWindow Wn)
        {
            // Dispose timers from previous slide
            DisposeAllActiveTimers();
            // Check activated slide for new timers
            Slide activatedSlide = Wn.View.Slide;
            _activeTimers = FindTimersOnSlide(activatedSlide)
                .Select(timerShape => new TimerData(timerShape))
                .ToList();
        }       

        private void OnSlideShowEnd(Presentation Pres)
        {
            DisposeAllActiveTimers();
        }

        private void DisposeAllActiveTimers()
        {
            _activeTimers.ForEach(t => t.Dispose());
            _activeTimers.Clear();
        }

        private IEnumerable<Shape> FindTimersOnSlide(Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Tags[Constants.TimerTagName] == Constants.DigitalTimerTagValue)
                {
                    yield return shape;
                }
            }
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
