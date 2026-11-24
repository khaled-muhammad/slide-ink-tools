using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace InteractiveDisplay
{
    internal sealed class ToolbarWindow : OverlayForm
    {
        private const int ButtonSize = 44;
        private const int Gap = 4;
        private const int PaddingSize = 8;
        private const int LauncherSize = 58;

        private readonly List<ToolbarButton> primary = new List<ToolbarButton>();
        private readonly List<ToolbarButton> additional = new List<ToolbarButton>();
        private readonly ToolTip tips = new ToolTip();
        private readonly Label status;
        private readonly ToolbarButton launcher;
        private ToolbarButton expand;
        private int dividerY;

        public event EventHandler StateChanged;
        public event MouseEventHandler DragStarted;
        public event MouseEventHandler DragMoved;
        public event EventHandler DragFinished;

        public bool IsExpanded { get; private set; }
        public bool IsClosed { get; private set; }

        public ToolbarWindow() : base("Interactive Display Tools", false)
        {
            BackColor = Color.FromArgb(24, 34, 47);
            AttachDrag(Add(ToolbarGlyph.Grip, "Move", null));

            status = new Label
            {
                ForeColor = Color.FromArgb(210, 224, 239),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Visible = false
            };
            Controls.Add(status);

            launcher = CreateButton(ToolbarGlyph.Pen, "Open ink tools", delegate { SetClosed(false); });
            launcher.Circular = true;
            launcher.Visible = false;
            AttachDrag(launcher);
        }

        public ToolbarButton Add(ToolbarGlyph glyph, string label, Action action, bool expandedOnly = false)
        {
            ToolbarButton button = CreateButton(glyph, label, action);
            (expandedOnly ? additional : primary).Add(button);
            return button;
        }

        public void AddViewControls()
        {
            expand = Add(ToolbarGlyph.Expand, "Expand tools", ToggleExpanded);
            Add(ToolbarGlyph.Close, "Close toolbar", delegate { SetClosed(true); });
        }

        private ToolbarButton CreateButton(ToolbarGlyph glyph, string label, Action action)
        {
            var button = new ToolbarButton(glyph);
            SetHint(button, label);
            if (action != null)
                button.Click += delegate { action(); };
            Controls.Add(button);
            return button;
        }

        public void SetHint(ToolbarButton button, string label, string details = null)
        {
            button.Text = label;
            button.AccessibleName = label;
            button.AccessibleDescription = details;
            tips.SetToolTip(button, details == null ? label : label + Environment.NewLine + details);
        }

        public void SetStatus(string message)
        {
            status.Text = message ?? string.Empty;
        }

        private void AttachDrag(ToolbarButton button)
        {
            button.MouseDown += delegate(object sender, MouseEventArgs args)
            {
                var handler = DragStarted;
                if (handler != null)
                    handler(sender, args);
            };
            button.MouseMove += delegate(object sender, MouseEventArgs args)
            {
                var handler = DragMoved;
                if (handler != null)
                    handler(sender, args);
            };
            button.MouseUp += delegate { RaiseDragFinished(); };
            button.MouseCaptureChanged += delegate
            {
                if (!button.Capture)
                    RaiseDragFinished();
            };
        }

        private void RaiseDragFinished()
        {
            var handler = DragFinished;
            if (handler != null)
                handler(this, EventArgs.Empty);
        }

        private void ToggleExpanded()
        {
            IsExpanded = !IsExpanded;
            expand.Glyph = IsExpanded ? ToolbarGlyph.Collapse : ToolbarGlyph.Expand;
            SetHint(expand, IsExpanded ? "Compact tools" : "Expand tools");
            RaiseStateChanged();
        }

        private void SetClosed(bool closed)
        {
            IsClosed = closed;
            Opacity = closed ? 0.82 : 1;
            RaiseStateChanged();
        }

        private void RaiseStateChanged()
        {
            var handler = StateChanged;
            if (handler != null)
                handler(this, EventArgs.Empty);
            UpdateWindowRegion();
        }

        public Size LayoutForViewport(int availableWidth)
        {
            SuspendLayout();
            try
            {
                foreach (ToolbarButton button in primary)
                    button.Visible = !IsClosed;
                foreach (ToolbarButton button in additional)
                    button.Visible = !IsClosed && IsExpanded;
                launcher.Visible = IsClosed;
                status.Visible = !IsClosed && IsExpanded && status.Text.Length > 0;
                dividerY = 0;

                if (IsClosed)
                {
                    launcher.Bounds = new Rectangle(5, 5, LauncherSize - 10, LauncherSize - 10);
                    return new Size(LauncherSize, LauncherSize);
                }

                int columns = Math.Max(1, (availableWidth - 2 * PaddingSize + Gap) / (ButtonSize + Gap));
                columns = Math.Min(columns, Math.Max(primary.Count, IsExpanded ? additional.Count : 0));
                int width = columns * (ButtonSize + Gap) - Gap + 2 * PaddingSize;
                int y = PlaceButtons(primary, PaddingSize, columns);
                if (IsExpanded)
                {
                    dividerY = y + 4;
                    y = PlaceButtons(additional, y + 9, columns);
                    if (status.Text.Length > 0)
                    {
                        status.Bounds = new Rectangle(PaddingSize + 4, y + 1, width - 2 * PaddingSize - 8, 26);
                        y += 27;
                    }
                }
                return new Size(width, y + PaddingSize);
            }
            finally
            {
                ResumeLayout(false);
                Invalidate();
            }
        }

        private static int PlaceButtons(List<ToolbarButton> buttons, int top, int columns)
        {
            for (int index = 0; index < buttons.Count; index++)
            {
                int x = PaddingSize + index % columns * (ButtonSize + Gap);
                int y = top + index / columns * (ButtonSize + Gap);
                buttons[index].Bounds = new Rectangle(x, y, ButtonSize, ButtonSize);
            }
            int rows = (buttons.Count + columns - 1) / columns;
            return top + rows * (ButtonSize + Gap) - Gap;
        }

        protected override void OnSizeChanged(EventArgs args)
        {
            base.OnSizeChanged(args);
            UpdateWindowRegion();
        }

        private void UpdateWindowRegion()
        {
            if (Width <= 0 || Height <= 0)
                return;

            using (GraphicsPath outline = Outline(new RectangleF(0, 0, Width, Height)))
                Region = new Region(outline);
        }

        private GraphicsPath Outline(RectangleF bounds)
        {
            if (!IsClosed)
                return ToolbarButton.RoundedRectangle(bounds, 17);

            var path = new GraphicsPath();
            path.AddEllipse(bounds);
            return path;
        }

        protected override void OnPaint(PaintEventArgs args)
        {
            base.OnPaint(args);
            args.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath outline = Outline(new RectangleF(0.7f, 0.7f, Width - 1.4f, Height - 1.4f)))
            using (var border = new Pen(Color.FromArgb(65, 83, 103), 1))
            {
                args.Graphics.DrawPath(border, outline);
                if (dividerY > 0)
                    args.Graphics.DrawLine(border, PaddingSize + 4, dividerY, Width - PaddingSize - 4, dividerY);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                tips.Dispose();
            base.Dispose(disposing);
        }
    }

    internal enum ToolbarGlyph
    {
        Grip, Browse, Pen, Highlight, Eraser, Palm, Undo, Clear, Previous, Next, NewSlide, ExitShow,
        Blackout, Expand, Collapse, Close, Swatch, StrokeWidth
    }

    internal sealed class ToolbarButton : Button
    {
        private ToolbarGlyph glyph;
        private bool active;
        private bool hovered;
        private bool pressed;
        private bool suppressClick;

        public Color Swatch { get; set; }
        public float StrokeWidth { get; set; }
        public bool Circular { get; set; }

        public ToolbarGlyph Glyph
        {
            get { return glyph; }
            set { glyph = value; Invalidate(); }
        }

        public bool Active
        {
            get { return active; }
            set
            {
                if (active == value)
                    return;
                active = value;
                Invalidate();
            }
        }

        public ToolbarButton(ToolbarGlyph icon)
        {
            glyph = icon;
            UseVisualStyleBackColor = false;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw, true);
            if (icon == ToolbarGlyph.Grip)
                Cursor = Cursors.SizeAll;
        }

        public void SuppressDragClick()
        {
            suppressClick = true;
        }

        protected override void OnMouseEnter(EventArgs args)
        {
            base.OnMouseEnter(args);
            hovered = true;
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs args)
        {
            base.OnMouseLeave(args);
            hovered = false;
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs args)
        {
            suppressClick = false;
            pressed = args.Button == MouseButtons.Left;
            Invalidate();
            base.OnMouseDown(args);
        }

        protected override void OnMouseUp(MouseEventArgs args)
        {
            pressed = false;
            Invalidate();
            base.OnMouseUp(args);
            suppressClick = false;
        }

        protected override void OnMouseCaptureChanged(EventArgs args)
        {
            if (!Capture)
            {
                pressed = false;
                Invalidate();
            }
            base.OnMouseCaptureChanged(args);
        }

        protected override void OnClick(EventArgs args)
        {
            if (!suppressClick)
                base.OnClick(args);
        }

        protected override void OnSizeChanged(EventArgs args)
        {
            base.OnSizeChanged(args);
            if (!Circular || Width <= 0 || Height <= 0)
                return;

            using (var outline = new GraphicsPath())
            {
                outline.AddEllipse(0, 0, Width, Height);
                Region = new Region(outline);
            }
        }

        protected override void OnPaint(PaintEventArgs args)
        {
            Graphics canvas = args.Graphics;
            GraphicsState state = canvas.Save();
            using (var background = new SolidBrush(Parent == null ? BackColor : Parent.BackColor))
                canvas.FillRectangle(background, ClientRectangle);
            canvas.SmoothingMode = SmoothingMode.AntiAlias;
            if (active || hovered || pressed)
            {
                var region = new RectangleF(1, 1, Width - 2, Height - 2);
                using (var surface = new SolidBrush(active ? Color.FromArgb(41, 76, 106) : Color.FromArgb(36, 51, 69)))
                {
                    if (Circular)
                        canvas.FillEllipse(surface, region);
                    else
                    {
                        using (GraphicsPath outline = RoundedRectangle(region, 11))
                            canvas.FillPath(surface, outline);
                    }
                }
            }

            canvas.TranslateTransform((Width - 24) / 2f, (Height - 24) / 2f);
            Color foreground = active || Circular ? Color.FromArgb(137, 214, 255) : Color.FromArgb(226, 235, 246);
            using (var ink = new Pen(foreground, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            {
                DrawGlyph(canvas, ink);
            }
            canvas.Restore(state);
        }

        private void DrawGlyph(Graphics canvas, Pen ink)
        {
            switch (glyph)
            {
                case ToolbarGlyph.Grip:
                    using (var brush = new SolidBrush(ink.Color))
                    {
                        for (int row = 0; row < 3; row++)
                        {
                            canvas.FillEllipse(brush, 8, 6 + row * 5, 2.2f, 2.2f);
                            canvas.FillEllipse(brush, 14, 6 + row * 5, 2.2f, 2.2f);
                        }
                    }
                    break;
                case ToolbarGlyph.Browse:
                    canvas.DrawPolygon(ink, new[] { new Point(6, 4), new Point(18, 13), new Point(13, 14),
                        new Point(15, 19), new Point(12, 20), new Point(9, 15), new Point(6, 18) });
                    break;
                case ToolbarGlyph.Pen:
                    canvas.DrawPolygon(ink, new[] { new Point(5, 19), new Point(7, 13), new Point(16, 4),
                        new Point(20, 8), new Point(11, 17) });
                    canvas.DrawLine(ink, 13, 7, 17, 11);
                    break;
                case ToolbarGlyph.Highlight:
                    canvas.DrawPolygon(ink, new[] { new Point(8, 13), new Point(16, 5), new Point(20, 9),
                        new Point(12, 17), new Point(7, 17) });
                    canvas.DrawLine(ink, 9, 12, 13, 16);
                    canvas.DrawLine(ink, 4, 20, 17, 20);
                    break;
                case ToolbarGlyph.Eraser:
                    canvas.DrawPolygon(ink, new[] { new Point(4, 14), new Point(13, 5), new Point(20, 12),
                        new Point(13, 19), new Point(9, 19) });
                    canvas.DrawLine(ink, 9, 9, 16, 16);
                    canvas.DrawLine(ink, 13, 19, 21, 19);
                    break;
                case ToolbarGlyph.Palm:
                    using (var hand = new GraphicsPath())
                    {
                        hand.AddLines(new[] { new Point(8, 19), new Point(4, 13), new Point(5, 11),
                            new Point(8, 14), new Point(8, 6), new Point(10, 5), new Point(11, 6),
                            new Point(11, 11), new Point(11, 4), new Point(13, 4), new Point(14, 5),
                            new Point(14, 11), new Point(14, 6), new Point(16, 6), new Point(17, 7),
                            new Point(17, 12), new Point(17, 9), new Point(19, 9), new Point(20, 11) });
                        hand.AddBezier(20, 11, 20, 17, 18, 21, 14, 21);
                        hand.AddBezier(14, 21, 10, 21, 10, 21, 8, 19);
                        canvas.DrawPath(ink, hand);
                    }
                    break;
                case ToolbarGlyph.Undo:
                    canvas.DrawLines(ink, new[] { new Point(9, 5), new Point(5, 10), new Point(9, 14) });
                    using (var curve = new GraphicsPath())
                    {
                        curve.AddLine(5, 10, 15, 10);
                        curve.AddBezier(15, 10, 22, 10, 22, 19, 15, 19);
                        canvas.DrawPath(ink, curve);
                    }
                    break;
                case ToolbarGlyph.Clear:
                    canvas.DrawLine(ink, 5, 7, 19, 7);
                    canvas.DrawLines(ink, new[] { new Point(9, 7), new Point(10, 4), new Point(14, 4), new Point(15, 7) });
                    canvas.DrawLines(ink, new[] { new Point(7, 7), new Point(8, 20), new Point(16, 20), new Point(17, 7) });
                    canvas.DrawLine(ink, 11, 11, 11, 16);
                    canvas.DrawLine(ink, 13, 11, 13, 16);
                    break;
                case ToolbarGlyph.Blackout:
                    using (var screen = RoundedRectangle(new RectangleF(4, 5, 16, 13), 3))
                    {
                        canvas.DrawPath(ink, screen);
                        canvas.DrawLine(ink, 7, 15, 17, 8);
                    }
                    break;
                case ToolbarGlyph.Previous:
                    Chevron(canvas, ink, new Point(14, 6), new Point(8, 12), new Point(14, 18));
                    break;
                case ToolbarGlyph.Next:
                    Chevron(canvas, ink, new Point(10, 6), new Point(16, 12), new Point(10, 18));
                    break;
                case ToolbarGlyph.NewSlide:
                    using (var slide = RoundedRectangle(new RectangleF(3, 5, 18, 14), 3))
                    {
                        canvas.DrawPath(ink, slide);
                        canvas.DrawLine(ink, 8.5f, 12, 15.5f, 12);
                        canvas.DrawLine(ink, 12, 8.5f, 12, 15.5f);
                    }
                    break;
                case ToolbarGlyph.ExitShow:
                    canvas.DrawLines(ink, new[] { new Point(13, 5), new Point(5, 5), new Point(5, 19), new Point(13, 19) });
                    canvas.DrawLine(ink, 10, 12, 21, 12);
                    Chevron(canvas, ink, new Point(17, 8), new Point(21, 12), new Point(17, 16));
                    break;
                case ToolbarGlyph.StrokeWidth:
                    ink.Width = StrokeWidth;
                    canvas.DrawLine(ink, 5, 12, 19, 12);
                    break;
                case ToolbarGlyph.Expand:
                    Chevron(canvas, ink, new Point(6, 10), new Point(12, 16), new Point(18, 10));
                    break;
                case ToolbarGlyph.Collapse:
                    Chevron(canvas, ink, new Point(6, 14), new Point(12, 8), new Point(18, 14));
                    break;
                case ToolbarGlyph.Close:
                    canvas.DrawLine(ink, 7, 7, 17, 17);
                    canvas.DrawLine(ink, 7, 17, 17, 7);
                    break;
                case ToolbarGlyph.Swatch:
                    using (var brush = new SolidBrush(Swatch))
                        canvas.FillEllipse(brush, 5, 5, 14, 14);
                    if (active)
                        canvas.DrawEllipse(ink, 2, 2, 20, 20);
                    else
                    {
                        using (var border = new Pen(Color.FromArgb(125, 157, 186)))
                            canvas.DrawEllipse(border, 5, 5, 14, 14);
                    }
                    break;
            }
        }

        private static void Chevron(Graphics canvas, Pen ink, Point start, Point elbow, Point end)
        {
            canvas.DrawLines(ink, new[] { start, elbow, end });
        }

        internal static GraphicsPath RoundedRectangle(RectangleF rectangle, float radius)
        {
            float diameter = Math.Min(radius * 2, Math.Min(rectangle.Width, rectangle.Height));
            var outline = new GraphicsPath();
            outline.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
            outline.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
            outline.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
            outline.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
            outline.CloseFigure();
            return outline;
        }
    }
}
