/*
    Spreadalonia - A spreadsheet control for Avalonia
    Copyright (C) 2023  Giorgio Bianchini, University of Bristol
 
    This library is free software; you can redistribute it and/or
    modify it under the terms of the GNU Lesser General Public
    License as published by the Free Software Foundation; either
    version 2.1 of the License, or (at your option) any later version.

    This library is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
    Lesser General Public License for more details.

    You should have received a copy of the GNU Lesser General Public
    License along with this library; if not, see <https://www.gnu.org/licenses/>.
*/

using Avalonia.Input;
using Avalonia.Media.Imaging;
using System.Reflection;

namespace Spreadalonia
{
    internal static class Cursors
    {
        private static readonly Cursor SelectColumnCursor24 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.SelectColumnCursor_24.png")), new Avalonia.PixelPoint(11, 22));
        private static readonly Cursor SelectColumnCursor16 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.SelectColumnCursor_16.png")), new Avalonia.PixelPoint(8, 14));
        private static readonly Cursor SelectColumnCursor12 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.SelectColumnCursor_12.png")), new Avalonia.PixelPoint(5, 10));

        private static readonly Cursor SelectRowCursor24 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.SelectRowCursor_24.png")), new Avalonia.PixelPoint(22, 12));
        private static readonly Cursor SelectRowCursor16 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.SelectRowCursor_16.png")), new Avalonia.PixelPoint(14, 7));
        private static readonly Cursor SelectRowCursor12 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.SelectRowCursor_12.png")), new Avalonia.PixelPoint(10, 6));

        private static readonly Cursor ResizeColumnCursor34 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.ResizeColumnCursor_34.png")), new Avalonia.PixelPoint(17, 17));
        private static readonly Cursor ResizeColumnCursor25 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.ResizeColumnCursor_25.png")), new Avalonia.PixelPoint(12, 12));
        private static readonly Cursor ResizeColumnCursor18 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.ResizeColumnCursor_18.png")), new Avalonia.PixelPoint(9, 9));

        private static readonly Cursor ResizeRowCursor34 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.ResizeRowCursor_34.png")), new Avalonia.PixelPoint(17, 17));
        private static readonly Cursor ResizeRowCursor25 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.ResizeRowCursor_25.png")), new Avalonia.PixelPoint(12, 12));
        private static readonly Cursor ResizeRowCursor18 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.ResizeRowCursor_18.png")), new Avalonia.PixelPoint(9, 9));

        private static readonly Cursor CrossCursor30 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.CrossCursor_30.png")), new Avalonia.PixelPoint(15, 15));
        private static readonly Cursor CrossCursor21 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.CrossCursor_21.png")), new Avalonia.PixelPoint(10, 10));
        private static readonly Cursor CrossCursor15 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.CrossCursor_15.png")), new Avalonia.PixelPoint(7, 7));

        private static readonly Cursor MoveCursor32 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.MoveCursor_32.png")), new Avalonia.PixelPoint(16, 16));
        private static readonly Cursor MoveCursor24 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.MoveCursor_24.png")), new Avalonia.PixelPoint(12, 12));
        private static readonly Cursor MoveCursor17 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.MoveCursor_17.png")), new Avalonia.PixelPoint(8, 8));

        private static readonly Cursor FillCursor26 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.FillCursor_26.png")), new Avalonia.PixelPoint(13, 13));
        private static readonly Cursor FillCursor21 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.FillCursor_21.png")), new Avalonia.PixelPoint(10, 10));
        private static readonly Cursor FillCursor14 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.FillCursor_14.png")), new Avalonia.PixelPoint(7, 7));

        private static readonly Cursor MoveCopyCursor34 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.MoveCopyCursor_34.png")), new Avalonia.PixelPoint(16, 18));
        private static readonly Cursor MoveCopyCursor26 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.MoveCopyCursor_26.png")), new Avalonia.PixelPoint(12, 14));
        private static readonly Cursor MoveCopyCursor18 = new Cursor(new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Cursors.MoveCopyCursor_18.png")), new Avalonia.PixelPoint(8, 9));

        public static Cursor SelectColumnCursor(double dpi)
        {
            if (dpi < 1.5)
            {
                return SelectColumnCursor12;
            }
            else if (dpi < 2 || System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
            {
                return SelectColumnCursor16;
            }
            else
            {
                return SelectColumnCursor24;
            }
        }

        public static Cursor SelectRowCursor(double dpi)
        {
            if (dpi < 1.5)
            {
                return SelectRowCursor12;
            }
            else if (dpi < 2 || System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
            {
                return SelectRowCursor16;
            }
            else
            {
                return SelectRowCursor24;
            }
        }

        public static Cursor ResizeColumnCursor(double dpi)
        {
            if (dpi < 1.5)
            {
                return ResizeColumnCursor18;
            }
            else if (dpi < 2 || System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
            {
                return ResizeColumnCursor25;
            }
            else
            {
                return ResizeColumnCursor34;
            }
        }

        public static Cursor ResizeRowCursor(double dpi)
        {
            if (dpi < 1.5)
            {
                return ResizeRowCursor18;
            }
            else if (dpi < 2 || System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
            {
                return ResizeRowCursor25;
            }
            else
            {
                return ResizeRowCursor34;
            }
        }

        public static Cursor CrossCursor(double dpi)
        {
            if (dpi < 1.5)
            {
                return CrossCursor15;
            }
            else if (dpi < 2 || System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
            {
                return CrossCursor21;
            }
            else
            {
                return CrossCursor30;
            }
        }

        public static Cursor MoveCursor(double dpi)
        {
            if (dpi < 1.5)
            {
                return MoveCursor17;
            }
            else if (dpi < 2 || System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
            {
                return MoveCursor24;
            }
            else
            {
                return MoveCursor32;
            }
        }

        public static Cursor MoveCopyCursor(double dpi)
        {
            if (dpi < 1.5)
            {
                return MoveCopyCursor18;
            }
            else if (dpi < 2 || System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
            {
                return MoveCopyCursor26;
            }
            else
            {
                return MoveCopyCursor34;
            }
        }

        public static Cursor FillCursor(double dpi)
        {
            if (dpi < 1.5)
            {
                return FillCursor14;
            }
            else if (dpi < 2 || System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
            {
                return FillCursor21;
            }
            else
            {
                return FillCursor26;
            }
        }
    }
}
