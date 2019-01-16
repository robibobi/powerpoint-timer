using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
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
            Duration = TimerShape.TextFrame.TextRange.Text;
        }

        private void TimerTick(object _, ElapsedEventArgs __)
        {
            string timeText = GetShapeText();
            if(TimeSpan.TryParse(timeText, out TimeSpan time))
            {
                if (time.TotalSeconds == 0)
                    return;
                time = time - OneSecond;
                SetShapeText(time.ToString());
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
            return TimerShape.TextFrame.TextRange.Text;
        }

        public void Dispose()
        {
            SetShapeText(Duration);
            _timer.Dispose();
        }
    }
}
