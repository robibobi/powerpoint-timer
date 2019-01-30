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

        private readonly Timer _timer;
        private readonly string _duration;
        private readonly Shape _timerShape;

        public TimerData(Shape timerShape)
        {
            _timer = new Timer(OneSecond.TotalMilliseconds);
            _timer.Elapsed += TimerTick;
            _timer.Start();
            _timerShape = timerShape;
            _duration = GetShapeText();
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
                SetShapeText($"Invalid format: {_duration}");
            }
        }

        private void SetShapeText(string text)
        {
            _timerShape.TextFrame.TextRange.Text = text;
        }

        private string GetShapeText()
        {
            if (_timerShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return _timerShape.TextFrame2.TextRange.Text;
            }

            return "No text found.";
        }

        public void Dispose()
        {
            SetShapeText(_duration);
            _timer.Dispose();
        }
    }
}
