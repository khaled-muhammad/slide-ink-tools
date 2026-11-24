using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

[assembly: ComVisible(false)]

namespace InteractiveDisplay
{
    [ComVisible(true)]
    [Guid("FC353AEF-32D1-41FB-9AC6-4BCAA006AEC5")]
    [ProgId("InteractiveDisplay.PowerPoint")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class PowerPointAddIn : IDTExtensibility2
    {
        private const string ProductName = "Slide Ink Tools";
        private PowerPoint.Application application;
        private ShowSession session;
        private ShowSession pending;
        private Timer finishTimer;

        public void OnConnection(object host, ConnectMode mode, object addIn, ref Array custom)
        {
            application = (PowerPoint.Application)host;
            application.SlideShowBegin += BeginShow;
            application.SlideShowNextSlide += ChangeSlide;
            application.SlideShowEnd += EndShow;
        }

        public void OnDisconnection(DisconnectMode mode, ref Array custom)
        {
            Stop();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Stop();
        }

        public void OnStartupComplete(ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }

        private void BeginShow(PowerPoint.SlideShowWindow window)
        {
            if (session != null || pending != null)
                return;

            try
            {
                session = new ShowSession(window);
            }
            catch (Exception error)
            {
                Report(ProductName + " could not start.", error);
            }
        }

        private void ChangeSlide(PowerPoint.SlideShowWindow window)
        {
            if (session != null)
                session.SuspendForTransition(window);
        }

        private void EndShow(PowerPoint.Presentation presentation)
        {
            if (session == null || !session.Owns(presentation))
                return;

            pending = session;
            session = null;
            pending.Dispose();
            finishTimer = new Timer { Interval = 100 };
            finishTimer.Tick += FinishShow;
            finishTimer.Start();
        }

        private void FinishShow(object sender, EventArgs args)
        {
            if (pending.IsShowWindowOpen)
                return;

            finishTimer.Dispose();
            finishTimer = null;
            ShowSession finished = pending;
            pending = null;

            if (!finished.Ink.HasMarks)
                return;

            while (MessageBox.Show(
                "Keep the marks as editable objects on the slides? Use PowerPoint Save or Save As afterward to write the deck to disk.",
                ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    PowerPointInkWriter.AddToPresentation(application, finished.Presentation, finished.Ink);
                    return;
                }
                catch (Exception error)
                {
                    Report("PowerPoint could not keep the marks. Choose Yes to retry or No to discard.", error);
                }
            }
        }

        private void Stop()
        {
            if (application != null)
            {
                application.SlideShowBegin -= BeginShow;
                application.SlideShowNextSlide -= ChangeSlide;
                application.SlideShowEnd -= EndShow;
                application = null;
            }

            if (session != null)
            {
                session.Dispose();
                session = null;
            }

            if (finishTimer != null)
            {
                finishTimer.Dispose();
                finishTimer = null;
            }
            pending = null;
        }

        private static void Report(string message, Exception error)
        {
            MessageBox.Show(message + Environment.NewLine + error.Message, ProductName,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
