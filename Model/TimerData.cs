using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace PowerPointTimer.Model
{
    class TimerData : IDisposable
    {
        private static readonly TimeSpan OneSecond = TimeSpan.FromSeconds(1);
        private Timer _timer;

        public string Duration { get; }

        public Shape TimerShape { get; }

        public TimerData(Shape timerShape)
        {
            _timer = new Timer(1000);
            _timer.Elapsed += TimerTick;
            _timer.Start();
            TimerShape = timerShape;
            Duration = GetShapeText();
        }

        private void TimerTick(object _, ElapsedEventArgs __)
        {
            string timeText = GetShapeText();
            if(TimeSpan.TryParseExact(timeText, "mm\\:ss",
                CultureInfo.InvariantCulture, out TimeSpan time))
            {
                if (time.TotalSeconds == 0)
                    return;
                time = time - OneSecond;
                SetShapeText(time.ToString("mm\\:ss"));
            } else
            {
                SetShapeText($"Invalid format: {Duration}");
            }
        }

        private void SetShapeText(string text)
        {
            TimerShape.TextFrame.TextRange.Text = text;
        }

        private string GetShapeText()
        {
            if (TimerShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return TimerShape.TextFrame2.TextRange.Text;
            }

            return "No text found.";
        }

        public void Dispose()
        {
            SetShapeText(Duration);
            _timer.Dispose();
        }
    }
}
