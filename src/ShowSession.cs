using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace InteractiveDisplay
{
    internal enum Tool { Navigate, Pen, Highlighter, Eraser }

    internal sealed class ShowSession : IDisposable
    {
        private readonly PowerPoint.SlideShowWindow show;
        private readonly IntPtr showHandle;
        private readonly InkForm inkWindow;
        private readonly InputForm inputWindow;
        private readonly ToolbarWindow toolbar;
        private readonly ToolbarButton browseButton;
        private readonly ToolbarButton penButton;
        private readonly ToolbarButton highlighterButton;
        private readonly ToolbarButton eraserButton;
        private readonly ToolbarButton palmButton;
        private readonly ToolbarButton blackoutButton;
        private readonly Dictionary<Color, ToolbarButton> colorButtons = new Dictionary<Color, ToolbarButton>();
        private readonly Dictionary<int, ToolbarButton> penWidthButtons = new Dictionary<int, ToolbarButton>();
        private readonly PalmCalibration palm = new PalmCalibration();
        private readonly Timer timer;
        private Rectangle viewport;
        private SizeF slideSize;
        private Point? toolbarPlacement;
        private DisplayPreferences display;
        private IntPtr monitor;
        private bool displayChanged = true;
        private Tool tool = Tool.Navigate;
        private Color selectedColor = Color.FromArgb(225, 42, 49);
        private int penWidth = 3;
        private InkPage page;
        private int slideId;
        private Stroke activeStroke;
        private int activePointer = -1;
        private Point pointerOrigin;
        private PointF previousPoint;
        private int palmWidth;
        private bool movedFromOrigin;
        private bool palmContact;
        private bool eraserActive;
        private bool erasedInGesture;
        private bool draggingToolbar;
        private Point dragOrigin;
        private Point dragOffset;
        private Point dragPosition;
        private DateTime resumeAfter;
        private bool blackout;
        private bool insertingSlide;
        private bool disposed;

        public readonly SlideInk Ink = new SlideInk();
        public readonly PowerPoint.Presentation Presentation;

        public ShowSession(PowerPoint.SlideShowWindow window)
        {
            show = window;
            showHandle = new IntPtr(window.HWND);
            Presentation = window.Presentation;
            PowerPoint.PageSetup setup = Presentation.PageSetup;
            slideSize = new SizeF(setup.SlideWidth, setup.SlideHeight);
            inkWindow = new InkForm(PaintInk);
            inputWindow = new InputForm(StartPointer, MovePointer, FinishPointer, OnPointerCanceled, id => id == activePointer);
            toolbar = new ToolbarWindow();
            toolbar.DisplayTopologyChanged += delegate { displayChanged = true; };
            toolbar.StateChanged += ToolbarStateChanged;
            toolbar.DragStarted += StartDrag;
            toolbar.DragMoved += MoveDrag;
            toolbar.DragFinished += delegate { FinishDrag(); };

            browseButton = AddButton(ToolbarGlyph.Browse, "Browse", delegate { SelectTool(Tool.Navigate); });
            penButton = AddButton(ToolbarGlyph.Pen, "Pen", delegate { SelectTool(Tool.Pen); });
            highlighterButton = AddButton(ToolbarGlyph.Highlight, "Highlight", delegate { SelectTool(Tool.Highlighter); });
            eraserButton = AddButton(ToolbarGlyph.Eraser, "Eraser", delegate { SelectTool(Tool.Eraser); });
            AddButton(ToolbarGlyph.Undo, "Undo", delegate { if (page != null) page.Undo(); inkWindow.QueueFrame(); });
            AddButton(ToolbarGlyph.Previous, "Prev", delegate { Navigate(false); });
            AddButton(ToolbarGlyph.Next, "Next", delegate { Navigate(true); });
            AddButton(ToolbarGlyph.NewSlide, "New slide", InsertBlankSlide);
            AddButton(ToolbarGlyph.ExitShow, "Exit presentation", delegate { show.View.Exit(); });
            AddPenWidth(1);
            AddPenWidth(3);
            AddPenWidth(6);
            AddColor("Red", selectedColor);
            AddColor("Blue", Color.DodgerBlue);
            AddColor("Black", Color.FromArgb(23, 25, 30));
            AddColor("Yellow", Color.Gold);
            palmButton = AddButton(ToolbarGlyph.Palm, palm.Label, TogglePalm, true);
            AddButton(ToolbarGlyph.Clear, "Clear", delegate { if (page != null) page.Clear(); inkWindow.QueueFrame(); }, true);
            blackoutButton = AddButton(ToolbarGlyph.Blackout, "Blank", ToggleBlackout, true);
            toolbar.AddViewControls();
            browseButton.Active = true;

            timer = new Timer { Interval = 70 };
            timer.Tick += delegate { RefreshShow(); };
            timer.Start();
            RefreshShow();
        }

        public bool Owns(PowerPoint.Presentation presentation)
        {
            return Presentation.Equals(presentation);
        }

        public bool IsShowWindowOpen { get { return NativeWindowApi.IsWindow(showHandle) && NativeWindowApi.IsWindowVisible(showHandle); } }

        public void SuspendForTransition(PowerPoint.SlideShowWindow window)
        {
            if (window.HWND == showHandle.ToInt64())
                PauseInput();
        }

        private void PauseInput()
        {
            CancelPointer();
            resumeAfter = DateTime.UtcNow.AddMilliseconds(200);
            inputWindow.Hide();
            inkWindow.Hide();
        }

        private void RefreshShow()
        {
            if (disposed)
                return;

            try
            {
                Rectangle showArea;
                if (!IsShowWindowOpen || !NativeWindowApi.TryClientRectangle(showHandle, out showArea))
                {
                    HideOverlay();
                    return;
                }

                using (new DpiScope())
                {
                    PowerPoint.SlideShowView view = show.View;
                    PowerPoint.PpSlideShowState state = view.State;
                    if (state == PowerPoint.PpSlideShowState.ppSlideShowDone)
                    {
                        HideOverlay();
                        return;
                    }

                    blackout = state == PowerPoint.PpSlideShowState.ppSlideShowBlackScreen ||
                        state == PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen;
                    blackoutButton.Active = blackout;
                    int currentId = view.Slide.SlideID;
                    if (slideId != currentId)
                    {
                        PauseInput();
                        slideId = currentId;
                        page = Ink.Page(slideId);
                    }

                    Rectangle newViewport = FitSlide(showArea, slideSize);
                    bool viewportChanged = newViewport != viewport;
                    if (viewportChanged)
                    {
                        CancelPointer();
                        viewport = newViewport;
                        inkWindow.Bounds = viewport;
                        inputWindow.Bounds = viewport;
                        inkWindow.QueueFrame();
                    }
                    RefreshDisplay();
                    if (viewportChanged)
                        PositionToolbar();
                    ApplyVisibility();
                }
            }
            catch (COMException)
            {
                HideOverlay();
            }
        }

        private void RefreshDisplay()
        {
            IntPtr current = NativeWindowApi.MonitorFromWindow(showHandle, 2); // MONITOR_DEFAULTTONEAREST
            if (!displayChanged && monitor == current)
                return;

            displayChanged = false;
            monitor = current;
            DisplayPreferences next = DisplayPreferences.ForMonitor(current);
            if (display == null || !display.MatchesMonitor(next))
            {
                draggingToolbar = false;
                toolbarPlacement = next.LoadPlacement();
            }
            if (display == null || !display.MatchesPointerScale(next))
            {
                CancelPointer();
                palm.Restore(next.LoadPalmThreshold());
                UpdatePalmButton();
            }
            display = next;
        }

        private static Rectangle FitSlide(Rectangle bounds, SizeF slide)
        {
            float scale = Math.Min(bounds.Width / slide.Width, bounds.Height / slide.Height);
            int width = (int)Math.Round(slide.Width * scale);
            int height = (int)Math.Round(slide.Height * scale);
            return new Rectangle(bounds.Left + (bounds.Width - width) / 2, bounds.Top + (bounds.Height - height) / 2, width, height);
        }

        private void HideOverlay()
        {
            CancelPointer();
            inputWindow.Hide();
            inkWindow.Hide();
            toolbar.Hide();
        }

        private void ApplyVisibility()
        {
            if (viewport.IsEmpty)
                return;

            bool showInk = !blackout && DateTime.UtcNow >= resumeAfter;
            SetVisible(inkWindow, showInk);
            SetVisible(inputWindow, showInk && tool != Tool.Navigate && !toolbar.IsClosed);
            SetVisible(toolbar, true);
        }

        private void SetVisible(OverlayForm window, bool visible)
        {
            if (!visible)
            {
                window.Hide();
                return;
            }
            if (window.Visible)
                return;

            window.Show(new WindowOwner(showHandle));
            if (window == inkWindow)
                inkWindow.QueueFrame();
            if (window != toolbar && toolbar.Visible)
                NativeWindowApi.RaiseWithoutActivation(toolbar.Handle);
        }

        private ToolbarButton AddButton(ToolbarGlyph glyph, string label, Action action, bool expandedOnly = false)
        {
            return toolbar.Add(glyph, label, delegate
            {
                try { action(); }
                catch (COMException) { RefreshShow(); }
            }, expandedOnly);
        }

        private void AddColor(string name, Color color)
        {
            ToolbarButton button = AddButton(ToolbarGlyph.Swatch, name, delegate
            {
                selectedColor = color;
                foreach (var choice in colorButtons)
                    choice.Value.Active = choice.Key == color;
                if (tool == Tool.Navigate || tool == Tool.Eraser)
                    SelectTool(Tool.Pen);
            }, true);
            button.Swatch = color;
            button.Active = color == selectedColor;
            colorButtons.Add(color, button);
        }

        private void AddPenWidth(int width)
        {
            ToolbarButton button = AddButton(ToolbarGlyph.StrokeWidth, "Pen width: " + width + " pt", delegate
            {
                penWidthButtons[penWidth].Active = false;
                penWidth = width;
                penWidthButtons[penWidth].Active = true;
                if (tool != Tool.Pen)
                    SelectTool(Tool.Pen);
            }, true);
            button.StrokeWidth = width;
            button.Active = width == penWidth;
            penWidthButtons.Add(width, button);
        }

        private void SelectTool(Tool next)
        {
            CancelPointer();
            if (palm.InProgress && next == Tool.Navigate)
            {
                palm.Toggle();
                UpdatePalmButton();
            }
            tool = next;
            browseButton.Active = next == Tool.Navigate;
            penButton.Active = next == Tool.Pen;
            highlighterButton.Active = next == Tool.Highlighter;
            eraserButton.Active = next == Tool.Eraser;
            ApplyVisibility();
        }

        private void ToolbarStateChanged(object sender, EventArgs args)
        {
            if (toolbar.IsClosed)
                SelectTool(Tool.Navigate);
            else if (!toolbar.IsExpanded && palm.InProgress)
            {
                CancelPointer();
                palm.Toggle();
                UpdatePalmButton();
            }
            PositionToolbar();
            ApplyVisibility();
        }

        private void PositionToolbar()
        {
            Size size = toolbar.LayoutForViewport(viewport.Width);
            var bounds = new Rectangle(DesiredToolbarLocation(size), size);
            if (toolbar.Bounds != bounds)
                toolbar.Bounds = bounds;
        }

        private Point DesiredToolbarLocation(Size size)
        {
            Point offset = new Point(16, 16);
            if (draggingToolbar)
                offset = dragPosition;
            else if (toolbarPlacement.HasValue)
                offset = DisplayPreferences.OffsetFromPlacement(toolbarPlacement.Value, viewport.Size);
            int x = Math.Max(0, Math.Min(offset.X, viewport.Width - size.Width));
            int y = Math.Max(0, Math.Min(offset.Y, viewport.Height - size.Height));
            return new Point(viewport.Left + x, viewport.Top + y);
        }

        private void StartDrag(object sender, MouseEventArgs args)
        {
            if (args.Button != MouseButtons.Left)
                return;
            draggingToolbar = true;
            dragOrigin = Cursor.Position;
            dragOffset = new Point(toolbar.Left - viewport.Left, toolbar.Top - viewport.Top);
            dragPosition = dragOffset;
            ((Control)sender).Capture = true;
        }

        private void MoveDrag(object sender, MouseEventArgs args)
        {
            if (!draggingToolbar)
                return;
            Point cursor = Cursor.Position;
            int dx = cursor.X - dragOrigin.X;
            int dy = cursor.Y - dragOrigin.Y;
            if (dragPosition == dragOffset && Math.Abs(dx) + Math.Abs(dy) < 6)
                return;

            ((ToolbarButton)sender).SuppressDragClick();
            dragPosition = new Point(dragOffset.X + dx, dragOffset.Y + dy);
            Point location = DesiredToolbarLocation(toolbar.Size);
            if (toolbar.Location != location)
                toolbar.Location = location;
        }

        private void FinishDrag()
        {
            if (!draggingToolbar)
                return;

            draggingToolbar = false;
            if (dragPosition == dragOffset || viewport.IsEmpty)
                return;

            var offset = new Point(toolbar.Left - viewport.Left, toolbar.Top - viewport.Top);
            Point placement = DisplayPreferences.PlacementFromOffset(offset, viewport.Size);
            toolbarPlacement = placement;
            if (display != null)
                display.SavePlacement(placement);
        }

        private void Navigate(bool forward)
        {
            PauseInput();
            if (forward)
                show.View.Next();
            else
                show.View.Previous();
        }

        private void InsertBlankSlide()
        {
            if (insertingSlide)
                return;

            insertingSlide = true;
            PowerPoint.Slide inserted = null;
            try
            {
                PowerPoint.SlideShowView view = show.View;
                if (view.IsNamedShow == Office.MsoTriState.msoTrue ||
                    Presentation.SlideShowSettings.RangeType != PowerPoint.PpSlideShowRangeType.ppShowAll)
                {
                    ShowInsertionMessage("Start a standard all-slide show to add and open a slide from the toolbar.");
                    return;
                }

                int index = view.Slide.SlideIndex + 1;
                PauseInput();
                Presentation.Application.StartNewUndoEntry();
                inserted = Presentation.Slides.Add(index, PowerPoint.PpSlideLayout.ppLayoutBlank);
                view.GotoSlide(inserted.SlideIndex, Office.MsoTriState.msoTrue);
            }
            catch (COMException error)
            {
                string message = "PowerPoint could not add and open a blank slide.";
                if (inserted != null)
                {
                    try { inserted.Delete(); }
                    catch (COMException) { message = "A blank slide was added to the deck. PowerPoint could not open it in this show."; }
                }
                RefreshShow();
                ShowInsertionMessage(message + Environment.NewLine + error.Message);
                return;
            }
            finally
            {
                insertingSlide = false;
            }

            if (palm.InProgress)
            {
                palm.Toggle();
                UpdatePalmButton();
            }
            if (tool == Tool.Navigate || tool == Tool.Eraser)
                SelectTool(Tool.Pen);
            RefreshShow();
        }

        private void ShowInsertionMessage(string message)
        {
            MessageBox.Show(new WindowOwner(showHandle), message, "New slide", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ToggleBlackout()
        {
            CancelPointer();
            show.View.State = blackout ? PowerPoint.PpSlideShowState.ppSlideShowRunning
                : PowerPoint.PpSlideShowState.ppSlideShowBlackScreen;
            RefreshShow();
        }

        private PointF ToSlide(Point screen)
        {
            return new PointF(
                Math.Min(slideSize.Width, Math.Max(0, (screen.X - viewport.Left) * slideSize.Width / viewport.Width)),
                Math.Min(slideSize.Height, Math.Max(0, (screen.Y - viewport.Top) * slideSize.Height / viewport.Height)));
        }

        private void TogglePalm()
        {
            CancelPointer();
            bool wasEnabled = palm.Threshold > 0;
            palm.Toggle();
            if (wasEnabled && display != null)
                display.SavePalmThreshold(0);
            if (palm.InProgress && (tool == Tool.Navigate || tool == Tool.Eraser))
                SelectTool(Tool.Pen);
            UpdatePalmButton();
        }

        private void UpdatePalmButton()
        {
            palmButton.Active = palm.Threshold > 0 || palm.InProgress;
            inputWindow.MeasureTouch = palm.NeedsWidth;
            toolbar.SetHint(palmButton, palm.Label, palm.Readings);
            toolbar.SetStatus(palm.InProgress || palm.Label == "Palm unavailable" ? palm.Label : null);
            PositionToolbar();
        }

        private void StartPointer(PointerSample sample)
        {
            if (activePointer != -1 || page == null || tool == Tool.Navigate || !viewport.Contains(sample.Position))
                return;

            activePointer = sample.Id;
            pointerOrigin = sample.Position;
            previousPoint = ToSlide(sample.Position);
            if (palm.InProgress)
            {
                palm.StartSample();
                palm.Sample(sample.ContactWidth);
                return;
            }

            if (tool == Tool.Eraser || sample.PenEraser)
            {
                eraserActive = true;
                EraseBetween(previousPoint, previousPoint, 12);
                return;
            }

            if (sample.ContactWidth >= palm.Threshold && palm.Threshold > 0)
            {
                palmContact = true;
                palmWidth = sample.ContactWidth;
                return;
            }

            Color color = tool == Tool.Highlighter ? Color.FromArgb(100, selectedColor) : selectedColor;
            activeStroke = new Stroke(color, tool == Tool.Highlighter ? 14 : penWidth);
            activeStroke.Append(previousPoint);
            QueueSegment(previousPoint, previousPoint, activeStroke.Width);
        }

        private void MovePointer(PointerSample sample)
        {
            if (activePointer != sample.Id)
                return;

            int dx = sample.Position.X - pointerOrigin.X;
            int dy = sample.Position.Y - pointerOrigin.Y;
            movedFromOrigin |= (long)dx * dx + (long)dy * dy >= 24 * 24;
            if (palm.InProgress)
            {
                palm.Sample(sample.ContactWidth);
                return;
            }

            PointF point = ToSlide(sample.Position);
            if (!eraserActive && palm.Threshold > 0 && sample.ContactWidth >= palm.Threshold)
            {
                palmContact = true;
                palmWidth = Math.Max(palmWidth, sample.ContactWidth);
                if (activeStroke != null)
                {
                    inkWindow.QueueFrame(ToPixels(activeStroke.Bounds));
                    activeStroke = null;
                }
            }

            if (palmContact && !eraserActive)
            {
                if (!movedFromOrigin)
                    return;
                eraserActive = true;
                previousPoint = ToSlide(pointerOrigin);
            }

            if (eraserActive)
            {
                float radius = Math.Max(12, palmWidth * slideSize.Width / viewport.Width / 2);
                EraseBetween(previousPoint, point, radius);
            }
            else if (activeStroke != null && previousPoint != point)
            {
                activeStroke.Append(point);
                QueueSegment(previousPoint, point, activeStroke.Width);
            }
            previousPoint = point;
        }

        private void EraseBetween(PointF from, PointF to, float radius)
        {
            float dx = to.X - from.X;
            float dy = to.Y - from.Y;
            int steps = Math.Max(1, (int)Math.Ceiling(Math.Sqrt(dx * dx + dy * dy) / radius));
            Rectangle dirty = Rectangle.Empty;
            for (int step = 0; step <= steps; step++)
            {
                var point = new PointF(from.X + dx * step / steps, from.Y + dy * step / steps);
                for (int index = page.Marks.Count - 1; index >= 0; index--)
                {
                    Stroke stroke = page.Marks[index];
                    if (!stroke.IsNear(point, radius))
                        continue;
                    if (!erasedInGesture)
                    {
                        page.Remember();
                        erasedInGesture = true;
                    }
                    Rectangle pixels = ToPixels(stroke.Bounds);
                    dirty = dirty.IsEmpty ? pixels : Rectangle.Union(dirty, pixels);
                    page.Marks.RemoveAt(index);
                }
            }
            if (!dirty.IsEmpty)
                inkWindow.QueueFrame(dirty);
        }

        private void FinishPointer(PointerSample sample)
        {
            if (activePointer != sample.Id)
                return;

            MovePointer(sample);
            if (palm.InProgress && movedFromOrigin)
            {
                palm.CompleteSample();
                if (!palm.InProgress && display != null)
                    display.SavePalmThreshold(palm.Threshold);
                UpdatePalmButton();
            }
            if (activeStroke != null)
            {
                page.Remember();
                page.Marks.Add(activeStroke);
                activeStroke = null;
            }
            CancelPointer();
        }

        private void OnPointerCanceled(int id)
        {
            if (id == -1 || id == activePointer)
                CancelPointer();
        }

        private void CancelPointer()
        {
            activePointer = -1;
            if (activeStroke != null)
                inkWindow.QueueFrame(ToPixels(activeStroke.Bounds));
            activeStroke = null;
            erasedInGesture = false;
            eraserActive = false;
            palmContact = false;
            palmWidth = 0;
            movedFromOrigin = false;
        }

        private void QueueSegment(PointF start, PointF end, float width)
        {
            float radius = width / 2;
            inkWindow.QueueFrame(ToPixels(RectangleF.FromLTRB(Math.Min(start.X, end.X) - radius,
                Math.Min(start.Y, end.Y) - radius, Math.Max(start.X, end.X) + radius, Math.Max(start.Y, end.Y) + radius)));
        }

        private Rectangle ToPixels(RectangleF bounds)
        {
            float scaleX = viewport.Width / slideSize.Width;
            float scaleY = viewport.Height / slideSize.Height;
            return Rectangle.FromLTRB((int)Math.Floor(bounds.Left * scaleX) - 3, (int)Math.Floor(bounds.Top * scaleY) - 3,
                (int)Math.Ceiling(bounds.Right * scaleX) + 3, (int)Math.Ceiling(bounds.Bottom * scaleY) + 3);
        }

        private void PaintInk(Graphics canvas, Rectangle dirty)
        {
            if (page == null)
                return;
            canvas.SmoothingMode = SmoothingMode.AntiAlias;
            canvas.ScaleTransform(viewport.Width / slideSize.Width, viewport.Height / slideSize.Height);
            foreach (Stroke stroke in page.Marks)
                if (ToPixels(stroke.Bounds).IntersectsWith(dirty))
                    DrawStroke(canvas, stroke);
            if (activeStroke != null)
                DrawStroke(canvas, activeStroke);
        }

        private static void DrawStroke(Graphics canvas, Stroke stroke)
        {
            if (stroke.Points.Count == 1)
            {
                PointF point = stroke.Points[0];
                float radius = stroke.Width / 2;
                using (var brush = new SolidBrush(stroke.Color))
                    canvas.FillEllipse(brush, point.X - radius, point.Y - radius, stroke.Width, stroke.Width);
                return;
            }

            using (var pen = new Pen(stroke.Color, stroke.Width)
            {
                StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round
            })
                canvas.DrawLines(pen, stroke.Points.ToArray());
        }

        public void Dispose()
        {
            disposed = true;
            timer.Dispose();
            toolbar.Dispose();
            inputWindow.Dispose();
            inkWindow.Dispose();
        }
    }

    internal class OverlayForm : Form
    {
        private readonly bool transparentToInput;
        public event EventHandler DisplayTopologyChanged;

        public OverlayForm(string title, bool transparentToInput)
        {
            Text = title;
            this.transparentToInput = transparentToInput;
            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
            DoubleBuffered = true;
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams parameters = base.CreateParams;
                parameters.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
                if (transparentToInput)
                    parameters.ExStyle |= 0x20; // WS_EX_TRANSPARENT on the layered ink window
                return parameters;
            }
        }

        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0x0021 || message.Msg == 0x024B) // Mouse or pointer activation
            {
                message.Result = new IntPtr(3);
                return;
            }
            if (message.Msg == 0x007e || message.Msg == 0x0219) // Display or device change
            {
                EventHandler handler = DisplayTopologyChanged;
                if (handler != null)
                    handler(this, EventArgs.Empty);
            }
            base.WndProc(ref message);
        }
    }

    internal sealed class WindowOwner : IWin32Window
    {
        public WindowOwner(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; private set; }
    }

    internal sealed class DpiScope : IDisposable
    {
        private readonly IntPtr previous = NativeWindowApi.SetThreadDpiAwarenessContext(new IntPtr(-4));
        public void Dispose()
        {
            if (previous != IntPtr.Zero)
                NativeWindowApi.SetThreadDpiAwarenessContext(previous);
        }
    }

    internal static class NativeWindowApi
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRectangle { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out NativeRectangle rectangle);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref Point point);
        [DllImport("user32.dll")] private static extern IntPtr GetMessageExtraInfo();
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
        [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] internal static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] internal static extern IntPtr MonitorFromWindow(IntPtr window, uint flags);
        [DllImport("user32.dll")] internal static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
        public static bool TryClientRectangle(IntPtr window, out Rectangle rectangle)
        {
            using (new DpiScope())
            {
                NativeRectangle client;
                var origin = Point.Empty;
                if (GetClientRect(window, out client) && ClientToScreen(window, ref origin)
                    && client.Right > client.Left && client.Bottom > client.Top)
                {
                    rectangle = new Rectangle(origin.X, origin.Y, client.Right - client.Left, client.Bottom - client.Top);
                    return true;
                }
                rectangle = Rectangle.Empty;
                return false;
            }
        }

        public static bool IsPromotedPointerMouseMessage()
        {
            return (GetMessageExtraInfo().ToInt64() & 0xffffff00L) == 0xff515700L;
        }

        public static void RaiseWithoutActivation(IntPtr window)
        {
            SetWindowPos(window, IntPtr.Zero, 0, 0, 0, 0, 0x0013); // no move, resize, or activation
        }

    }
}
