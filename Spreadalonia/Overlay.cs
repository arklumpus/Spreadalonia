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

using Avalonia;
using Avalonia.Animation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Immutable;

namespace Spreadalonia
{
    internal class Overlay : Control
    {
        public static readonly StyledProperty<ImmutableList<SelectionRange>> SelectionProperty = AvaloniaProperty.Register<Overlay, ImmutableList<SelectionRange>>(nameof(Selection), ImmutableList.Create<SelectionRange>());
        public ImmutableList<SelectionRange> Selection
        {
            get { return GetValue(SelectionProperty); }
            set { SetValue(SelectionProperty, value); }
        }

        public static readonly StyledProperty<Vector> OffsetProperty = AvaloniaProperty.Register<Overlay, Vector>(nameof(Offset), new Vector(0, 0));
        public Vector Offset
        {
            get { return GetValue(OffsetProperty); }
            set { SetValue(OffsetProperty, value); }
        }

        public static readonly StyledProperty<Spreadsheet> ContainerProperty = AvaloniaProperty.Register<Overlay, Spreadsheet>(nameof(Container));
        public Spreadsheet Container
        {
            get { return GetValue(ContainerProperty); }
            set { SetValue(ContainerProperty, value); }
        }

        public static readonly StyledProperty<Table> OwnerProperty = AvaloniaProperty.Register<Overlay, Table>(nameof(Owner));
        public Table Owner
        {
            get { return GetValue(OwnerProperty); }
            set { SetValue(OwnerProperty, value); }
        }

        public static readonly StyledProperty<SolidColorBrush> SelectionAccentProperty = AvaloniaProperty.Register<Overlay, SolidColorBrush>(nameof(SelectionAccent), new SolidColorBrush(Color.FromRgb(0, 114, 176)));
        public SolidColorBrush SelectionAccent
        {
            get { return GetValue(SelectionAccentProperty); }
            set { SetValue(SelectionAccentProperty, value); }
        }

        public enum IconType
        {
            Copy, Cut, Paste
        }

        public static readonly StyledProperty<IconType> IconProperty = AvaloniaProperty.Register<Overlay, IconType>(nameof(Icon));
        public IconType Icon
        {
            get { return GetValue(IconProperty); }
            set { SetValue(IconProperty, value); }
        }

        static Overlay()
        {
            AffectsRender<Overlay>(SelectionAccentProperty, OffsetProperty, SelectionProperty, IconProperty);
        }

        public Overlay()
        {
            this.Transitions = CachedTransitions;
        }

        private Transitions CachedTransitions = new Avalonia.Animation.Transitions() { new DoubleTransition() { Property = OpacityProperty, Duration = TimeSpan.FromMilliseconds(250), Delay = TimeSpan.FromMilliseconds(250) } };

        public void PauseTransitions()
        {
            this.Transitions = null;
        }

        public void StopTransitions()
        {
            this.Transitions = null;
            this.CachedTransitions = null;
        }

        public void ResumeTransitions()
        {
            this.Transitions = CachedTransitions;
        }


        public override void Render(DrawingContext context)
        {
            if (this.Opacity > 0 && this.Selection != null && this.Selection.Count > 0)
            {
                Brush overlayBrush = new SolidColorBrush(SelectionAccent.Color, 0.25);

                using (context.PushPreTransform(Matrix.CreateTranslation(-1, -1)))
                {
                    (int left, double offsetX, int top, double offsetY, int width, double actualWidth, double startWidth, int height, double actualHeight, double startHeight) = Owner.GetRange(this.Offset.X, this.Offset.Y, this.Bounds.Width, this.Bounds.Height);

                    using (context.PushPreTransform(Matrix.CreateTranslation(offsetX, offsetY)))
                    {
                        double[] xs = Owner.lastDrawnXs;
                        double[] ys = Owner.lastDrawnYs;

                        using (context.PushClip(new Rect(-offsetX + 1, -offsetY + 1, this.Bounds.Width, this.Bounds.Height)))
                        {
                            for (int i = 0; i < Selection.Count; i++)
                            {
                                if (Selection[i].Left <= left + width && Selection[i].Right >= left && Selection[i].Top <= top + height && Selection[i].Bottom >= top)
                                {
                                    double x0 = Selection[i].Left - left <= 0 ? -5 : xs[Selection[i].Left - left - 1];
                                    double y0 = Selection[i].Top - top <= 0 ? -5 : ys[Selection[i].Top - top - 1];

                                    double x1 = Selection[i].Right - left < 0 ? -5 : Selection[i].Right > left + width ? xs[xs.Length - 1] + 5 : xs[Selection[i].Right - left] + 1;
                                    double y1 = Selection[i].Bottom - top < 0 ? -5 : Selection[i].Bottom > top + height ? ys[ys.Length - 1] + 5 : ys[Selection[i].Bottom - top] + 1;

                                    x0 = Math.Max(-offsetX, x0);
                                    y0 = Math.Max(-offsetY, y0);

                                    x1 = Math.Min(-offsetX + this.Bounds.Width + 1, x1);
                                    y1 = Math.Min(-offsetY + this.Bounds.Height + 1, y1);

                                    context.FillRectangle(overlayBrush, new Rect(new Point(x0, y0).SnapToDevicePixels(this, true, true), new Point(x1, y1).SnapToDevicePixels(this, true, true)));

                                    if (this.Icon == IconType.Copy)
                                    {
                                        context.DrawImage(Icons.GetCopy((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1), new Rect(new Point((x0 + x1) * 0.5 - 24, (y0 + y1) * 0.5 - 24).SnapToDevicePixels(this, true, true), new Size(48, 48)));
                                    }
                                    else if (this.Icon == IconType.Cut)
                                    {
                                        context.DrawImage(Icons.GetCut((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1), new Rect(new Point((x0 + x1) * 0.5 - 24, (y0 + y1) * 0.5 - 24).SnapToDevicePixels(this, true, true), new Size(48, 48)));
                                    }
                                    else if (this.Icon == IconType.Paste)
                                    {
                                        context.DrawImage(Icons.GetPaste((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1), new Rect(new Point((x0 + x1) * 0.5 - 24, (y0 + y1) * 0.5 - 24).SnapToDevicePixels(this, true, true), new Size(48, 48)));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
