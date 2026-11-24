using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace InteractiveDisplay
{
    internal sealed class InkForm : OverlayForm
    {
        private readonly Action<Graphics, Rectangle> paint;
        private LayeredInkBuffer buffer;
        private Rectangle dirty;
        private bool frameQueued;

        public InkForm(Action<Graphics, Rectangle> paint) : base("Slide Ink Tools Overlay", true)
        {
            this.paint = paint;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams parameters = base.CreateParams;
                parameters.ExStyle |= 0x00080000; // WS_EX_LAYERED
                return parameters;
            }
        }

        public void QueueFrame()
        {
            QueueFrame(ClientRectangle);
        }

        public void QueueFrame(Rectangle pixels)
        {
            pixels = Rectangle.Intersect(pixels, ClientRectangle);
            if (pixels.Width <= 0 || pixels.Height <= 0 || !Visible || !IsHandleCreated || IsDisposed)
                return;

            dirty = dirty.IsEmpty ? pixels : Rectangle.Union(dirty, pixels);
            if (frameQueued)
                return;

            frameQueued = true;
            BeginInvoke((Action)RenderFrame);
        }

        private void RenderFrame()
        {
            frameQueued = false;
            Rectangle pixels = dirty;
            dirty = Rectangle.Empty;
            if (!Visible || IsDisposed || Width <= 0 || Height <= 0)
                return;

            using (new DpiScope())
            {
                if (buffer == null || buffer.Size != ClientSize)
                {
                    ReleaseBuffer();
                    try { buffer = new LayeredInkBuffer(ClientSize); }
                    catch (Win32Exception error)
                    {
                        System.Diagnostics.Trace.TraceError("Slide Ink Tools ink surface: " + error.Message);
                        return;
                    }
                    pixels = ClientRectangle;
                }
                pixels = Rectangle.Intersect(pixels, ClientRectangle);
                if (pixels.Width > 0 && pixels.Height > 0)
                    buffer.Paint(Handle, Bounds.Location, pixels, paint);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                ReleaseBuffer();
            base.Dispose(disposing);
        }

        private void ReleaseBuffer()
        {
            if (buffer == null)
                return;
            buffer.Dispose();
            buffer = null;
        }
    }

    internal sealed class LayeredInkBuffer : IDisposable
    {
        private IntPtr context;
        private IntPtr dib;
        private IntPtr selectedBefore;
        private Bitmap bitmap;
        private Graphics graphics;
        private bool presentWhole = true;

        public readonly Size Size;

        public LayeredInkBuffer(Size size)
        {
            Size = size;
            var info = new BitmapInfoHeader
            {
                Size = (uint)Marshal.SizeOf(typeof(BitmapInfoHeader)),
                Width = size.Width,
                Height = -size.Height,
                Planes = 1,
                BitCount = 32
            };

            try
            {
                context = CreateCompatibleDC(IntPtr.Zero);
                if (context == IntPtr.Zero)
                    throw new Win32Exception();
                IntPtr pixels;
                dib = CreateDIBSection(context, ref info, 0, out pixels, IntPtr.Zero, 0);
                if (dib == IntPtr.Zero)
                    throw new Win32Exception();
                selectedBefore = SelectObject(context, dib);
                if (selectedBefore == IntPtr.Zero)
                    throw new Win32Exception();
                bitmap = new Bitmap(size.Width, size.Height, size.Width * 4, PixelFormat.Format32bppPArgb, pixels);
                graphics = Graphics.FromImage(bitmap);
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        public unsafe void Paint(IntPtr window, Point position, Rectangle dirty, Action<Graphics, Rectangle> draw)
        {
            graphics.ResetTransform();
            graphics.ResetClip();
            graphics.SmoothingMode = SmoothingMode.None;
            graphics.CompositingMode = CompositingMode.SourceCopy;
            if (dirty.Size == Size)
                graphics.Clear(Color.Transparent);
            else
                graphics.FillRectangle(Brushes.Transparent, dirty);
            graphics.CompositingMode = CompositingMode.SourceOver;
            graphics.SetClip(dirty);
            draw(graphics, dirty);
            graphics.Flush(FlushIntention.Sync);

            Size dimensions = Size;
            Point origin = Point.Empty;
            var blend = new Blend { Opacity = 255, AlphaFormat = 1 };
            var rect = new DirtyRectangle { Left = dirty.Left, Top = dirty.Top, Right = dirty.Right, Bottom = dirty.Bottom };
            var update = new LayeredWindowInfo
            {
                Size = (uint)sizeof(LayeredWindowInfo),
                Position = &position,
                Dimensions = &dimensions,
                Source = context,
                Origin = &origin,
                Blend = &blend,
                Flags = 2, // ULW_ALPHA
                Dirty = presentWhole ? null : &rect
            };
            if (UpdateLayeredWindowIndirect(window, ref update))
                presentWhole = false;
            else
                presentWhole = !UpdateLayeredWindow(window, IntPtr.Zero, ref position, ref dimensions,
                    context, ref origin, 0, ref blend, 2);
        }

        public void Dispose()
        {
            if (graphics != null)
                graphics.Dispose();
            if (bitmap != null)
                bitmap.Dispose();
            if (selectedBefore != IntPtr.Zero)
                SelectObject(context, selectedBefore);
            if (dib != IntPtr.Zero)
                DeleteObject(dib);
            if (context != IntPtr.Zero)
                DeleteDC(context);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct BitmapInfoHeader
        {
            public uint Size;
            public int Width, Height;
            public ushort Planes, BitCount;
            public uint Compression, SizeImage;
            public int XPelsPerMeter, YPelsPerMeter;
            public uint ColorsUsed, ColorsImportant;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Blend { public byte Operation, Flags, Opacity, AlphaFormat; }

        [StructLayout(LayoutKind.Sequential)]
        private struct DirtyRectangle { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private unsafe struct LayeredWindowInfo
        {
            public uint Size;
            public IntPtr Destination;
            public Point* Position;
            public Size* Dimensions;
            public IntPtr Source;
            public Point* Origin;
            public uint ColorKey;
            public Blend* Blend;
            public uint Flags;
            public DirtyRectangle* Dirty;
        }

        [DllImport("gdi32.dll", SetLastError = true)] private static extern IntPtr CreateCompatibleDC(IntPtr source);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern bool DeleteDC(IntPtr context);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern IntPtr SelectObject(IntPtr context, IntPtr item);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern bool DeleteObject(IntPtr item);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern IntPtr CreateDIBSection(IntPtr context,
            ref BitmapInfoHeader info, uint usage, out IntPtr bits, IntPtr file, uint offset);
        [DllImport("user32.dll")] private static extern bool UpdateLayeredWindowIndirect(IntPtr window, ref LayeredWindowInfo info);
        [DllImport("user32.dll")] private static extern bool UpdateLayeredWindow(IntPtr window, IntPtr screen,
            ref Point position, ref Size size, IntPtr source, ref Point origin, uint colorKey, ref Blend blend, uint flags);
    }
}
