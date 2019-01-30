using Microsoft.Office.Tools.Ribbon;
using PowerPointTimer.Util;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointTimer
{
    public partial class TimerRibbon
    {
        private void AddDigitalTimerButton_Click(object sender, RibbonControlEventArgs e)
        {
            var slide = FindActiveSlide();
            if(slide != null)
            {
                AddDigitalTimerTextToSlide(slide);
            }
        }

        private Slide FindActiveSlide()
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

        private void AddDigitalTimerTextToSlide(Slide slide)
        {
            // Center the textBox on the Slide
            // 1. Calculate the top left corner 
            const float textBoxWidth = 180;
            const float textBoxHeight = 45;
            Presentation currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;
            float slideHeight = currentPresentation.PageSetup.SlideHeight;
            float slideWidth = currentPresentation.PageSetup.SlideWidth;
            float x = (slideWidth - textBoxWidth) * 0.5f;
            float y = (slideHeight - textBoxHeight) * 0.5f;
            // 2. Place the slide
            var textBox = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                x, y, textBoxWidth, textBoxHeight);
            textBox.TextFrame.TextRange.Text = Constants.DefaultTimeString;
            textBox.TextFrame.TextRange.Font.Size = 64;
            textBox.Tags.Add(Constants.TimerTagName,
                Constants.DigitalTimerTagValue);
        }
    }
}
