using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop;
using PowerPointTimer.Util;
using Microsoft.Office.Core;

namespace PowerPointTimer
{
    public partial class TimerRibbon
    {
        private void AddTimerButton_Click(object sender, RibbonControlEventArgs e)
        {
            var slide = SlideHelpers.FindActiveSlide();
            SlideHelpers.AddDigitalTimerTextToSlide(slide);
        }
    }
}
