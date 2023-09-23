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

using Avalonia.Controls;
using Avalonia.Media.Imaging;
using System;
using System.Reflection;

namespace Spreadalonia
{
    internal static class Icons
    {
        private static readonly Bitmap Copy48 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Copy_48.png"));
        private static readonly Bitmap Copy72 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Copy_72.png"));
        private static readonly Bitmap Copy96 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Copy_96.png"));

        private static readonly Bitmap Cut48 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Cut_48.png"));
        private static readonly Bitmap Cut72 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Cut_72.png"));
        private static readonly Bitmap Cut96 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Cut_96.png"));

        private static readonly Bitmap Paste48 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Paste_48.png"));
        private static readonly Bitmap Paste72 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Paste_72.png"));
        private static readonly Bitmap Paste96 = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("Spreadalonia.Icons.Paste_96.png"));

        public static Bitmap GetCopy(double dpi)
        {
            if (dpi < 1.5)
            {
                return Copy48;
            }
            else if (dpi < 2)
            {
                return Copy72;
            }
            else
            {
                return Copy96;
            }
        }

        public static Bitmap GetCut(double dpi)
        {
            if (dpi < 1.5)
            {
                return Cut48;
            }
            else if (dpi < 2)
            {
                return Cut72;
            }
            else
            {
                return Cut96;
            }
        }

        public static Bitmap GetPaste(double dpi)
        {
            if (dpi < 1.5)
            {
                return Paste48;
            }
            else if (dpi < 2)
            {
                return Paste72;
            }
            else
            {
                return Paste96;
            }
        }

        public static Func<double, Image> GetIcon16(string imageName)
        {
            return scaling =>
            {
                Image image = new Image() { Width = 16, Height = 16 };

                if (scaling <= 1)
                {
                    image.Source = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream(imageName + "-16.png"));
                }
                else if (scaling <= 1.5)
                {
                    image.Source = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream(imageName + "-24.png"));
                }
                else
                {
                    image.Source = new Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream(imageName + "-32.png"));
                }

                return image;
            };
        }
    }
}
