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
        private void TimerRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void AddTimerButton_Click(object sender, RibbonControlEventArgs e)
        {
            var slide = SlideHelpers.FindActiveSlide();

            // Test
            var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.Text = "00:10:00";
            textBox.Tags.Add("TimerTag", "Timer");

        }
    }
}
