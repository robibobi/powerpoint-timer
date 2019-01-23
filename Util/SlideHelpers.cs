using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointTimer.Util
{
    static class SlideHelpers
    {
        public static Slide FindActiveSlide()
        {
            var app = Globals.ThisAddIn.Application;
            if (app.ActivePresentation.Slides.Count > 0)
            {
                return app.ActiveWindow.View.Slide;
            }
            else
            {
                return null;
            }
        }

        public static void AddDigitalTimerTextToSlide(Slide slide)
        {
            // Test
            var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 120, 60);
            textBox.TextFrame.TextRange.Text = "05:00";
            textBox.TextFrame.TextRange.Font.Size = 28;
            textBox.Tags.Add(Constants.TimerTagName,
                Constants.DigitalTimerTagValue);

        }
    }
}
