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
    }
}
