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
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;

namespace Spreadalonia
{
    internal class DPIAwareBox : Viewbox
    {
        private Grid container;

        private Dictionary<double, Control> cachedControls;

        private double lastScaling = double.NaN;

        public DPIAwareBox()
        {
            container = new Grid();
            this.Child = container;

            cachedControls = new Dictionary<double, Control>();
        }

        public DPIAwareBox(Func<double, Control> getControlAtResolution)
        {
            container = new Grid();
            this.Child = container;

            cachedControls = new Dictionary<double, Control>();

            this.GetControlAtResolution = getControlAtResolution;
        }

        public Func<double, Control> GetControlAtResolution { get; set; }

        public override void Render(DrawingContext context)
        {
            double scaling = (this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1;

            if (scaling != lastScaling)
            {
                if (cachedControls.TryGetValue(lastScaling, out Control lastControl))
                {
                    lastControl.IsVisible = false;
                }

                if (cachedControls.TryGetValue(scaling, out Control control))
                {
                    control.IsVisible = true;
                }
                else
                {
                    control = this.GetControlAtResolution(scaling);
                    control.IsVisible = true;
                    container.Children.Add(control);

                    cachedControls[scaling] = control;
                }


                lastScaling = scaling;

                Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { this.InvalidateVisual(); });
            }

            base.Render(context);
        }
    }
}
