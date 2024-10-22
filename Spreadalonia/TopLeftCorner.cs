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
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using System.Collections.Immutable;

namespace Spreadalonia
{
    internal class TopLeftCorner : Control
    {
        public static readonly StyledProperty<Spreadsheet> ContainerProperty = AvaloniaProperty.Register<TopLeftCorner, Spreadsheet>(nameof(Container));
        public Spreadsheet Container
        {
            get { return GetValue(ContainerProperty); }
            set { SetValue(ContainerProperty, value); }
        }

        public static readonly StyledProperty<Table> OwnerProperty = AvaloniaProperty.Register<TopLeftCorner, Table>(nameof(Owner));
        public Table Owner
        {
            get { return GetValue(OwnerProperty); }
            set { SetValue(OwnerProperty, value); }
        }

        public static readonly StyledProperty<Color> HeaderBackgroundProperty = AvaloniaProperty.Register<TopLeftCorner, Color>(nameof(HeaderBackground), Color.FromRgb(240, 240, 240));
        public Color HeaderBackground
        {
            get { return GetValue(HeaderBackgroundProperty); }
            set { SetValue(HeaderBackgroundProperty, value); }
        }

        public static readonly StyledProperty<ImmutableList<SelectionRange>> SelectionProperty = AvaloniaProperty.Register<TopLeftCorner, ImmutableList<SelectionRange>>(nameof(Selection), ImmutableList.Create<SelectionRange>());
        public ImmutableList<SelectionRange> Selection
        {
            get { return GetValue(SelectionProperty); }
            set { SetValue(SelectionProperty, value); }
        }

        public static readonly StyledProperty<SolidColorBrush> SelectionAccentProperty = AvaloniaProperty.Register<TopLeftCorner, SolidColorBrush>(nameof(SelectionAccent), new SolidColorBrush(Color.FromRgb(0, 114, 176)));
        public SolidColorBrush SelectionAccent
        {
            get { return GetValue(SelectionAccentProperty); }
            set { SetValue(SelectionAccentProperty, value); }
        }



        static TopLeftCorner()
        {
            AffectsRender<TopLeftCorner>(HeaderBackgroundProperty, SelectionProperty, SelectionAccentProperty);
        }

        public TopLeftCorner()
        {
            
        }

        protected override void OnPointerEntered(PointerEventArgs e)
        {
            base.OnPointerEntered(e);

            this.Cursor = Cursors.CrossCursor((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1);
            hover = true;
            this.InvalidateVisual();
        }

        protected override void OnPointerExited(PointerEventArgs e)
        {
            base.OnPointerExited(e);
            hover = false;
            this.InvalidateVisual();
        }

        protected override void OnPointerPressed(PointerPressedEventArgs e)
        {
            base.OnPointerPressed(e);

            this.Container.Selection = ImmutableList.Create(new SelectionRange(0, 0, Owner.MaxTableWidth, Owner.MaxTableHeight));
        }

        bool hover = false;


        private static Brush triangleBrush = new SolidColorBrush(Colors.Black, 0.15);
        private static Brush triangleHoverBrush = new SolidColorBrush(Colors.Black, 0.3);

        public override void Render(DrawingContext context)
        {
            context.FillRectangle(new SolidColorBrush(HeaderBackground), new Rect(0, 0, this.Bounds.Width, this.Bounds.Height));
            Geometry triangleGeometry = new PathGeometry() { Figures = new PathFigures() { new PathFigure() { IsClosed = true, IsFilled = true, StartPoint = new Point(this.Bounds.Width - 16, this.Bounds.Height - 4), Segments = new PathSegments() { new LineSegment() { Point = new Point(this.Bounds.Width - 4, this.Bounds.Height - 4) }, new LineSegment() { Point = new Point(this.Bounds.Width - 4, this.Bounds.Height - 16) } } } } };

            bool found = false;

            for (int i= 0; i < Selection.Count; i++)
            {
                if (Selection[i].Left == 0 && Selection[i].Top == 0 && Selection[i].Right == Owner.MaxTableWidth && Selection[i].Bottom == Owner.MaxTableHeight)
                {
                    found = true;
                    break;
                }
            }
            
            if (found)
            {
                context.DrawGeometry(SelectionAccent, null, triangleGeometry);
            }
            else
            {
                if (hover)
                {
                    context.DrawGeometry(triangleHoverBrush, null, triangleGeometry);
                }
                else
                {
                    context.DrawGeometry(triangleBrush, null, triangleGeometry);
                }
            }
        }
    }
}
