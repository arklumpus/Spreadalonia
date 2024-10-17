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
using System;
using System.Collections.Generic;

namespace Spreadalonia
{
    internal class DPIAwareBox : Viewbox
    {
        private TopLevel _topLevel;
        private readonly Grid _container;
        private readonly Dictionary<double, Control> _cachedControls;
        private double _lastScaling = double.NaN;

        private Func<double, Control> GetControlAtResolution { get; }

        public DPIAwareBox(Func<double, Control> getControlAtResolution)
        {
            _container = new Grid();
            Child = _container;

            _cachedControls = new Dictionary<double, Control>();

            GetControlAtResolution = getControlAtResolution;
        }

        protected override void OnInitialized()
        {
            _topLevel = TopLevel.GetTopLevel(this);
            if (_topLevel != null)
            {
                UpdateScaling();
                _topLevel.ScalingChanged += (o, e) => UpdateScaling();
            }
            base.OnInitialized();
        }

        private void UpdateScaling()
        {
            var scaling = _topLevel.RenderScaling;
            if (_cachedControls.TryGetValue(_lastScaling, out Control lastControl))
            {
                lastControl.IsVisible = false;
            }

            if (_cachedControls.TryGetValue(scaling, out Control control))
            {
                control.IsVisible = true;
            }
            else
            {
                control = GetControlAtResolution(scaling);
                control.IsVisible = true;
                _container.Children.Add(control);

                _cachedControls[scaling] = control;
            }
            
            _lastScaling = scaling;
        }
    }
}
