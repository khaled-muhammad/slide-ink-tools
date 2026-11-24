using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace InteractiveDisplay
{
    internal sealed class DisplayPreferences
    {
        private const string RegistryRoot = @"HKEY_CURRENT_USER\Software\InteractiveDisplay\Displays\";
        private const string PositionValue = "ToolbarPosition";
        private const int PlacementScale = 10000;

        private readonly IntPtr monitor;
        private readonly Size pixels;
        private readonly string registryPath;
        private readonly string palmValue;

        private DisplayPreferences(IntPtr monitor, Size pixels, string devicePath)
        {
            this.monitor = monitor;
            this.pixels = pixels;
            if (devicePath != null)
            {
                using (SHA256 hash = SHA256.Create())
                {
                    byte[] identity = hash.ComputeHash(Encoding.Unicode.GetBytes(devicePath.ToUpperInvariant()));
                    registryPath = RegistryRoot + BitConverter.ToString(identity).Replace("-", "");
                }
                palmValue = string.Format(CultureInfo.InvariantCulture, "PalmThreshold_{0}x{1}", pixels.Width, pixels.Height);
            }
        }

        public static DisplayPreferences ForMonitor(IntPtr monitor)
        {
            var info = new MonitorInfo { Size = (uint)Marshal.SizeOf(typeof(MonitorInfo)) };
            if (monitor == IntPtr.Zero || !GetMonitorInfo(monitor, ref info))
                return new DisplayPreferences(monitor, Size.Empty, null);

            string devicePath = null;
            for (uint index = 0; ; index++)
            {
                var connected = new DisplayDevice { Size = (uint)Marshal.SizeOf(typeof(DisplayDevice)) };
                if (!EnumDisplayDevices(info.DeviceName, index, ref connected, 1)) // EDD_GET_DEVICE_INTERFACE_NAME
                    break;
                if ((connected.StateFlags & 1) == 0) // DISPLAY_DEVICE_ACTIVE
                    continue;
                if (string.IsNullOrEmpty(connected.DeviceId) ||
                    (devicePath != null && !StringComparer.OrdinalIgnoreCase.Equals(devicePath, connected.DeviceId)))
                {
                    // Mirrored displays can share one desktop monitor handle.
                    devicePath = null;
                    break;
                }
                devicePath = connected.DeviceId;
            }

            var size = new Size(info.Bounds.Right - info.Bounds.Left, info.Bounds.Bottom - info.Bounds.Top);
            return new DisplayPreferences(monitor, size, devicePath);
        }

        public bool MatchesMonitor(DisplayPreferences other)
        {
            if (registryPath == null && other.registryPath == null)
                return monitor == other.monitor;
            return registryPath == other.registryPath;
        }

        public bool MatchesPointerScale(DisplayPreferences other)
        {
            return MatchesMonitor(other) && pixels == other.pixels;
        }

        public Point? LoadPlacement()
        {
            string saved = Read(PositionValue) as string;
            if (saved == null)
                return null;

            string[] coordinates = saved.Split(',');
            int x, y;
            if (coordinates.Length != 2 ||
                !int.TryParse(coordinates[0], NumberStyles.None, CultureInfo.InvariantCulture, out x) ||
                !int.TryParse(coordinates[1], NumberStyles.None, CultureInfo.InvariantCulture, out y) ||
                x > PlacementScale || y > PlacementScale)
                return null;
            return new Point(x, y);
        }

        public void SavePlacement(Point placement)
        {
            Write(PositionValue, string.Format(CultureInfo.InvariantCulture, "{0},{1}", placement.X, placement.Y),
                RegistryValueKind.String);
        }

        public int LoadPalmThreshold()
        {
            object saved = Read(palmValue);
            if (!(saved is int))
                return 0;
            int threshold = (int)saved;
            return threshold >= 8 && threshold <= Math.Min(pixels.Width, pixels.Height) ? threshold : 0;
        }

        public void SavePalmThreshold(int threshold)
        {
            Write(palmValue, threshold, RegistryValueKind.DWord);
        }

        public static Point PlacementFromOffset(Point offset, Size viewport)
        {
            return new Point(ScaleOffset(offset.X, viewport.Width), ScaleOffset(offset.Y, viewport.Height));
        }

        public static Point OffsetFromPlacement(Point placement, Size viewport)
        {
            return new Point((int)Math.Round((double)placement.X * viewport.Width / PlacementScale),
                (int)Math.Round((double)placement.Y * viewport.Height / PlacementScale));
        }

        private static int ScaleOffset(int offset, int span)
        {
            return span <= 0 ? 0 : (int)Math.Round((double)Math.Max(0, Math.Min(span, offset)) * PlacementScale / span);
        }

        private object Read(string name)
        {
            if (registryPath != null)
            {
                try { return Registry.GetValue(registryPath, name, null); }
                catch (SecurityException) { }
                catch (UnauthorizedAccessException) { }
                catch (IOException) { }
            }
            return null;
        }

        private void Write(string name, object value, RegistryValueKind kind)
        {
            if (registryPath != null)
            {
                try { Registry.SetValue(registryPath, name, value, kind); }
                catch (SecurityException) { }
                catch (UnauthorizedAccessException) { }
                catch (IOException) { }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRectangle { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct MonitorInfo
        {
            public uint Size;
            public NativeRectangle Bounds, WorkArea;
            public uint Flags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string DeviceName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct DisplayDevice
        {
            public uint Size;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string DeviceName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string Name;
            public uint StateFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string DeviceId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string DeviceKey;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern bool EnumDisplayDevices(string device, uint index, ref DisplayDevice info, uint flags);
    }
}
