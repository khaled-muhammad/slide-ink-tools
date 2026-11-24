using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace InteractiveDisplay
{
    internal enum PointerKind : uint { Unknown, Pointer, Touch, Pen, Mouse }

    internal struct PointerSample
    {
        public readonly int Id;
        public readonly Point Position;
        public readonly PointerKind Kind;
        public readonly int ContactWidth;
        public readonly bool PenEraser;

        public PointerSample(int id, Point position, PointerKind kind = PointerKind.Unknown,
            int contactWidth = 0, bool penEraser = false)
        {
            Id = id;
            Position = position;
            Kind = kind;
            ContactWidth = contactWidth;
            PenEraser = penEraser;
        }
    }

    internal sealed class InputForm : OverlayForm
    {
        private const int DisablePressAndHold = 0x00000001;
        private const int DisablePenTapFeedback = 0x00000008;
        private const int DisablePenBarrelFeedback = 0x00000010;
        private const int DisableFlicks = 0x00010000;
        private const uint TouchContactVisualization = 1;
        private const uint GesturePressAndTap = 11;

        private readonly Action<PointerSample> start;
        private readonly Action<PointerSample> move;
        private readonly Action<PointerSample> finish;
        private readonly Action<int> cancel;
        private readonly Func<int, bool> acceptsUpdate;
        private readonly PointerInfo[] history = new PointerInfo[32];
        public bool MeasureTouch { get; set; }

        public InputForm(Action<PointerSample> start, Action<PointerSample> move, Action<PointerSample> finish,
            Action<int> cancel, Func<int, bool> acceptsUpdate)
            : base("Interactive Display Input", false)
        {
            this.start = start;
            this.move = move;
            this.finish = finish;
            this.cancel = cancel;
            this.acceptsUpdate = acceptsUpdate;
            BackColor = Color.Black;
            // A fully transparent layered window would let pointer events through.
            Opacity = 1 / 255d;
            Cursor = Cursors.Cross;
        }

        protected override void OnHandleCreated(EventArgs args)
        {
            base.OnHandleCreated(args);
            // The Win32 feedback types in this range cover touch and pen contact and gesture visuals.
            int enabled = 0;
            for (uint feedback = TouchContactVisualization; feedback <= GesturePressAndTap; feedback++)
                SetWindowFeedbackSetting(Handle, feedback, 0, sizeof(int), ref enabled);
        }

        protected override void WndProc(ref Message message)
        {
            int id = (int)(message.WParam.ToInt64() & 0xffff);
            long packed = message.LParam.ToInt64();
            var point = new Point(unchecked((short)packed), unchecked((short)(packed >> 16)));
            switch (message.Msg)
            {
                case 0x02cc: // WM_TABLET_QUERYSYSTEMGESTURESTATUS
                    message.Result = new IntPtr(DisablePressAndHold | DisablePenTapFeedback | DisablePenBarrelFeedback | DisableFlicks);
                    return;
                case 0x0246: start(ReadSample(id, point, true)); return; // WM_POINTERDOWN
                case 0x0245: MovePointer(id, point); return; // WM_POINTERUPDATE
                case 0x0247: finish(ReadSample(id, point, false)); return; // WM_POINTERUP
                case 0x024c: cancel(id); return; // WM_POINTERCAPTURECHANGED
                case 0x0201:
                case 0x0200:
                case 0x0202:
                    if (NativeWindowApi.IsPromotedPointerMouseMessage())
                        return;
                    var mouse = new PointerSample(0, PointToScreen(point), PointerKind.Mouse);
                    if (message.Msg == 0x0201) { Capture = true; start(mouse); }
                    if (message.Msg == 0x0200) move(mouse);
                    if (message.Msg == 0x0202) { finish(mouse); Capture = false; }
                    return;
                case 0x0215: cancel(0); break; // WM_CAPTURECHANGED
                case 0x001f: cancel(-1); break; // WM_CANCELMODE
            }
            base.WndProc(ref message);
        }

        private void MovePointer(int id, Point point)
        {
            if (!acceptsUpdate(id))
                return;

            // Collect the samples before passing them to the UI, which can process more window messages.
            PointerSample latest = ReadSample(id, point, false);
            uint count = (uint)history.Length;
            if (GetPointerInfoHistory((uint)id, ref count, history))
            {
                for (int index = (int)Math.Min(count, (uint)history.Length) - 1; index > 0; index--)
                    move(new PointerSample(id, history[index].Position, latest.Kind));
            }
            move(latest);
        }

        private PointerSample ReadSample(int id, Point point, bool readPenEraser)
        {
            PointerKind kind;
            GetPointerType((uint)id, out kind);
            int width = 0;
            bool eraser = false;
            if (kind == PointerKind.Touch && MeasureTouch)
            {
                TouchInfo touch;
                if (GetPointerTouchInfo((uint)id, out touch) && (touch.Mask & 1) != 0)
                    width = touch.RawContact.ShortSide > 0 ? touch.RawContact.ShortSide : touch.Contact.ShortSide;
            }
            if (kind == PointerKind.Pen && readPenEraser)
            {
                PenInfo pen;
                eraser = GetPointerPenInfo((uint)id, out pen) && (pen.Flags & 6) != 0;
            }
            return new PointerSample(id, point, kind, width, eraser);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PointerInfo
        {
            public PointerKind Type;
            public uint Id, Frame, Flags;
            public IntPtr Device, Target;
            public Point Position, Himetric, RawPosition, RawHimetric;
            public uint Time, History;
            public int Input;
            public uint Keys;
            public ulong Performance;
            public uint ButtonChange;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ContactRect
        {
            public int Left, Top, Right, Bottom;
            public int ShortSide { get { return Math.Max(0, Math.Min(Right - Left, Bottom - Top)); } }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct TouchInfo
        {
            public PointerInfo Pointer;
            public uint Flags, Mask;
            public ContactRect Contact, RawContact;
            public uint Orientation, Pressure;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PenInfo
        {
            public PointerInfo Pointer;
            public uint Flags, Mask, Pressure, Rotation;
            public int TiltX, TiltY;
        }

        [DllImport("user32.dll")] private static extern bool GetPointerType(uint id, out PointerKind kind);
        [DllImport("user32.dll")] private static extern bool GetPointerInfoHistory(uint id, ref uint entries, [Out] PointerInfo[] history);
        [DllImport("user32.dll")] private static extern bool GetPointerTouchInfo(uint id, out TouchInfo touch);
        [DllImport("user32.dll")] private static extern bool GetPointerPenInfo(uint id, out PenInfo pen);
        [DllImport("user32.dll")] private static extern bool SetWindowFeedbackSetting(
            IntPtr window, uint feedback, uint flags, uint size, [In] ref int configuration);
    }
}
