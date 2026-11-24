using System;

namespace InteractiveDisplay
{
    internal sealed class PalmCalibration
    {
        private readonly string[] instructions = { "Palm setup", "Swipe pen tip", "Swipe one finger", "Swipe palm" };
        private int step;
        private int penWidth;
        private int fingerWidth;
        private int peak;
        private bool unavailable;

        public int Threshold { get; private set; }
        public bool InProgress { get { return step != 0; } }
        public bool NeedsWidth { get { return InProgress || Threshold > 0; } }
        public string Readings { get; private set; }
        public string Label { get { return InProgress ? instructions[step] :
            Threshold > 0 ? "Palm on" : unavailable ? "Palm unavailable" : instructions[0]; } }

        public void Toggle()
        {
            bool start = !InProgress && Threshold == 0;
            Restore(0);
            step = start ? 1 : 0;
        }

        public void Restore(int threshold)
        {
            step = 0;
            Threshold = threshold;
            unavailable = false;
            penWidth = 0;
            fingerWidth = 0;
            peak = 0;
            Readings = null;
        }

        public void StartSample()
        {
            peak = 0;
        }

        public void Sample(int width)
        {
            peak = Math.Max(peak, width);
        }

        public void CompleteSample()
        {
            if (step == 1)
                penWidth = peak;
            else if (step == 2)
                fingerWidth = peak;
            else
            {
                int normal = Math.Max(penWidth, fingerWidth);
                unavailable = peak < Math.Max(18, Math.Max(normal + 12, normal * 1.65));
                if (!unavailable)
                    Threshold = Math.Max(normal + 8, (normal + peak) / 2);
                Readings = string.Format("Contact widths in pixels — pen {0}, finger {1}, palm {2}.", penWidth, fingerWidth, peak);
            }
            step = (step + 1) % instructions.Length;
        }
    }
}
