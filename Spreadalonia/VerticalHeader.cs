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
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Globalization;

namespace Spreadalonia
{
    internal class VerticalHeader : Control
    {
        public static readonly StyledProperty<Spreadsheet> ContainerProperty = AvaloniaProperty.Register<VerticalHeader, Spreadsheet>(nameof(Container));
        public Spreadsheet Container
        {
            get { return GetValue(ContainerProperty); }
            set { SetValue(ContainerProperty, value); }
        }

        public static readonly StyledProperty<Table> OwnerProperty = AvaloniaProperty.Register<VerticalHeader, Table>(nameof(Owner));
        public Table Owner
        {
            get { return GetValue(OwnerProperty); }
            set { SetValue(OwnerProperty, value); }
        }

        public static readonly StyledProperty<Vector> OffsetProperty = AvaloniaProperty.Register<VerticalHeader, Vector>(nameof(Offset), new Vector(0, 0));
        public Vector Offset
        {
            get { return GetValue(OffsetProperty); }
            set { SetValue(OffsetProperty, value); }
        }

        public static readonly StyledProperty<Color> GridColorProperty = AvaloniaProperty.Register<VerticalHeader, Color>(nameof(GridColor), Color.FromRgb(220, 220, 220));
        public Color GridColor
        {
            get { return GetValue(GridColorProperty); }
            set { SetValue(GridColorProperty, value); }
        }

        public static readonly StyledProperty<Color> HeaderBackgroundProperty = AvaloniaProperty.Register<VerticalHeader, Color>(nameof(HeaderBackground), Color.FromRgb(240, 240, 240));
        public Color HeaderBackground
        {
            get { return GetValue(HeaderBackgroundProperty); }
            set { SetValue(HeaderBackgroundProperty, value); }
        }

        public static readonly StyledProperty<Pen> HeaderSeparatorProperty = AvaloniaProperty.Register<VerticalHeader, Pen>(nameof(HeaderSeparator), new Pen(new SolidColorBrush(Color.FromRgb(170, 170, 170))));
        public Pen HeaderSeparator
        {
            get { return GetValue(HeaderSeparatorProperty); }
            set { SetValue(HeaderSeparatorProperty, value); }
        }

        public static readonly StyledProperty<ImmutableList<SelectionRange>> SelectionProperty = AvaloniaProperty.Register<VerticalHeader, ImmutableList<SelectionRange>>(nameof(Selection), ImmutableList.Create<SelectionRange>());
        public ImmutableList<SelectionRange> Selection
        {
            get { return GetValue(SelectionProperty); }
            set { SetValue(SelectionProperty, value); }
        }

        public static readonly StyledProperty<SolidColorBrush> SelectionAccentProperty = AvaloniaProperty.Register<VerticalHeader, SolidColorBrush>(nameof(SelectionAccent), new SolidColorBrush(Color.FromRgb(0, 114, 176)));
        public SolidColorBrush SelectionAccent
        {
            get { return GetValue(SelectionAccentProperty); }
            set { SetValue(SelectionAccentProperty, value); }
        }

        public static readonly StyledProperty<FontFamily> FontFamilyProperty = UserControl.FontFamilyProperty.AddOwner<HorizontalHeader>();
        public FontFamily FontFamily
        {
            get { return GetValue(FontFamilyProperty); }
            set { SetValue(FontFamilyProperty, value); }
        }

        public static readonly StyledProperty<double> FontSizeProperty = UserControl.FontSizeProperty.AddOwner<HorizontalHeader>();
        public double FontSize
        {
            get { return GetValue(FontSizeProperty); }
            set { SetValue(FontSizeProperty, value); }
        }

        public static readonly StyledProperty<FontStyle> FontStyleProperty = UserControl.FontStyleProperty.AddOwner<HorizontalHeader>();
        public FontStyle FontStyle
        {
            get { return GetValue(FontStyleProperty); }
            set { SetValue(FontStyleProperty, value); }
        }

        public static readonly StyledProperty<FontWeight> FontWeightProperty = UserControl.FontWeightProperty.AddOwner<HorizontalHeader>();
        public FontWeight FontWeight
        {
            get { return GetValue(FontWeightProperty); }
            set { SetValue(FontWeightProperty, value); }
        }

        public static readonly StyledProperty<IBrush> ForegroundProperty = UserControl.ForegroundProperty.AddOwner<HorizontalHeader>();
        public IBrush Foreground
        {
            get { return GetValue(ForegroundProperty); }
            set { SetValue(ForegroundProperty, value); }
        }

        static VerticalHeader()
        {
            AffectsRender<VerticalHeader>(OffsetProperty, GridColorProperty, HeaderBackgroundProperty, SelectionProperty, SelectionAccentProperty);
            AffectsMeasure<VerticalHeader>(OffsetProperty);
        }

        public VerticalHeader()
        {
            this.Transitions = CachedTransitions;
        }

        private Transitions CachedTransitions = new Avalonia.Animation.Transitions() { new VectorTransition() { Property = OffsetProperty, Duration = TimeSpan.FromMilliseconds(100) } };

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

        protected override void OnPointerReleased(PointerReleasedEventArgs e)
        {
            base.OnPointerReleased(e);

            PointerPoint currentPoint = e.GetCurrentPoint(this);

            double yPos = currentPoint.Position.Y - lastDrawnDelta;

            if (pointerPressedAction == 2)
            {
                double newHeight = Math.Max(10, originalRowHeight + yPos - dragStartY);

                if (newHeight != originalRowHeight)
                {
                    this.Container.UndoStackRowHeight.Push(new ValueStackFrame<int, double>(this.Container.Selection, new Dictionary<int, Reference<double>>() { { selectionStartRow, wasOriginalRowHeightDefault ? null : new Reference<double>(originalRowHeight) } }, new Dictionary<int, Reference<double>>() { { selectionStartRow, newHeight } }));
                    this.Container.UndoStack.Push(null);

                    this.Container.UndoStackRowForeground.Push(null);
                    this.Container.UndoStackRowTypeface.Push(null);

                    this.Container.UndoStackColumnForeground.Push(null);
                    this.Container.UndoStackColumnWidth.Push(null);
                    this.Container.UndoStackColumnTypeface.Push(null);

                    this.Container.UndoStackCellForeground.Push(null);
                    this.Container.UndoStackCellMargin.Push(null);
                    this.Container.UndoStackCellHorizontalAlignment.Push(null);
                    this.Container.UndoStackCellTypeface.Push(null);
                    this.Container.UndoStackCellVerticalAlignment.Push(null);

                    this.Container.ClearRedoStack();

                    if (this.Container.Selection.Count > 0 && this.Container.Selection[0].Top == selectionStartRow)
                    {
                        this.Container.RaiseCellSizeChanged();
                    }
                }
            }
        }

        protected override void OnPointerMoved(PointerEventArgs e)
        {
            base.OnPointerMoved(e);

            PointerPoint currentPoint = e.GetCurrentPoint(this);

            double closest = double.MaxValue;
            int clickedRow = -1;

            double yPos = currentPoint.Position.Y - lastDrawnDelta;

            if (lastDrawnYs != null)
            {
                for (int i = 0; i < lastDrawnYs.Length; i++)
                {
                    if (lastDrawnYs[i] > yPos && yPos >= (i == 0 ? 0 : lastDrawnYs[i - 1]))
                    {
                        clickedRow = i + lastDrawnTop;
                    }

                    closest = Math.Min(closest, Math.Abs(lastDrawnYs[i] - yPos));

                    if (closest <= 5 && clickedRow >= 0)
                    {
                        break;
                    }
                }
            }

            if (clickedRow >= 0)
            {

                if (!currentPoint.Properties.IsLeftButtonPressed)
                {
                    if (closest <= 5)
                    {
                        this.Cursor = Cursors.ResizeRowCursor((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1);
                    }
                    else
                    {
                        this.Cursor = Cursors.SelectRowCursor((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1);
                    }

                    if (hoverRow != clickedRow)
                    {
                        hoverRow = clickedRow;
                        this.InvalidateVisual();
                    }
                }
                else
                {
                    if (pointerPressedAction == 1)
                    {
                        if (selectionMode == 0)
                        {
                            this.Container.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(0, Math.Min(selectionStartRow, clickedRow), Owner.MaxTableWidth, Math.Max(selectionStartRow, clickedRow)));
                        }
                        else if (selectionMode == 1)
                        {
                            this.Container.Selection = previousSelection.Difference(new SelectionRange(0, Math.Min(selectionStartRow, clickedRow), Owner.MaxTableWidth, Math.Max(selectionStartRow, clickedRow)));
                        }
                    }
                    else if (pointerPressedAction == 2)
                    {
                        double newHeight = Math.Max(10, originalRowHeight + yPos - dragStartY);

                        Owner.RowHeights[selectionStartRow] = newHeight;
                        Owner.InvalidateVisual();
                        this.InvalidateVisual();
                    }
                }
            }
        }

        int pointerPressedAction = 0;
        int selectionStartRow = -1;
        int selectionMode = 0;
        double dragStartY = 0;
        double originalRowHeight = 0;
        bool wasOriginalRowHeightDefault = false;
        ImmutableList<SelectionRange> previousSelection;

        protected override void OnPointerPressed(PointerPressedEventArgs e)
        {
            base.OnPointerPressed(e);

            double closest = double.MaxValue;
            int clickedRow = -1;

            PointerPoint currentPoint = e.GetCurrentPoint(this);

            double yPos = currentPoint.Position.Y - lastDrawnDelta;

            if (lastDrawnYs != null)
            {
                for (int i = 0; i < lastDrawnYs.Length; i++)
                {
                    if (lastDrawnYs[i] > yPos && yPos >= (i == 0 ? 0 : lastDrawnYs[i - 1]))
                    {
                        clickedRow = i + lastDrawnTop;
                    }

                    closest = Math.Min(closest, Math.Abs(lastDrawnYs[i] - yPos));

                    if (closest <= 5 && clickedRow >= 0)
                    {
                        break;
                    }
                }
            }

            if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                if (closest > 5)
                {
                    if (e.KeyModifiers == Spreadsheet.ControlModifier)
                    {
                        bool found = false;

                        for (int i = 0; i < this.Selection.Count; i++)
                        {
                            if (this.Selection[i].Left == 0 && this.Selection[i].Right == Owner.MaxTableWidth && this.Selection[i].Top <= clickedRow && this.Selection[i].Bottom >= clickedRow)
                            {
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            this.Container.Selection = this.Container.Selection.Add(new SelectionRange(0, clickedRow, Owner.MaxTableWidth, clickedRow));
                            selectionStartRow = clickedRow;
                            selectionMode = 0;
                        }
                        else
                        {
                            previousSelection = this.Container.Selection;
                            this.Container.Selection = this.Container.Selection.Difference(new SelectionRange(0, clickedRow, Owner.MaxTableWidth, clickedRow));
                            selectionStartRow = clickedRow;
                            selectionMode = 1;
                        }
                    }
                    else if (e.KeyModifiers == KeyModifiers.Shift && this.Selection.Count > 0)
                    {
                        int prevStartRow = this.Selection[this.Selection.Count - 1].Top;
                        int prevEndRow = this.Selection[this.Selection.Count - 1].Bottom;

                        if (clickedRow >= prevStartRow)
                        {
                            selectionStartRow = prevStartRow;
                        }
                        else
                        {
                            selectionStartRow = prevEndRow;
                        }

                        this.Container.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(0, Math.Min(selectionStartRow, clickedRow), Owner.MaxTableWidth, Math.Max(selectionStartRow, clickedRow)));
                        selectionMode = 0;
                    }
                    else
                    {
                        this.Container.Selection = ImmutableList.Create(new SelectionRange(0, clickedRow, Owner.MaxTableWidth, clickedRow));
                        selectionStartRow = clickedRow;
                        selectionMode = 0;
                    }

                    this.Cursor = Cursors.CrossCursor((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1);
                    pointerPressedAction = 1;
                }
                else
                {
                    if (yPos - (clickedRow > 0 ? lastDrawnYs[clickedRow - 1] : 0) <= 5)
                    {
                        clickedRow--;
                    }

                    if (e.ClickCount == 1)
                    {
                        pointerPressedAction = 2;

                        selectionStartRow = clickedRow;

                        originalRowHeight = Owner.GetHeight(clickedRow, out wasOriginalRowHeightDefault);
                        dragStartY = yPos;
                    }
                    else if (e.ClickCount == 2)
                    {
                        AutoFitHeight(clickedRow);

                        if (this.Container.Selection.Count > 0 && this.Container.Selection[0].Top == clickedRow)
                        {
                            this.Container.RaiseCellSizeChanged();
                        }
                    }
                }
            }
            else if (e.GetCurrentPoint(this).Properties.IsRightButtonPressed)
            {
                bool found = false;

                for (int i = 0; i < this.Selection.Count; i++)
                {
                    if (this.Selection[i].Left == 0 && this.Selection[i].Right == Owner.MaxTableWidth && this.Selection[i].Top <= clickedRow && this.Selection[i].Bottom >= clickedRow)
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    this.Container.Selection = ImmutableList.Create(new SelectionRange(0, clickedRow, Owner.MaxTableWidth, clickedRow));
                }
            }
        }

        private void AutoFitHeight(int clickedRow)
        {
            bool found = false;
            double maxHeight = 10;

            Typeface defaultTypeFace = new Typeface(Owner.FontFamily, Owner.FontStyle, Owner.FontWeight);

            foreach (KeyValuePair<(int, int), string> kvp in Owner.Data)
            {
                if (kvp.Key.Item2 == clickedRow)
                {
                    if (!Owner.CellTypefaces.TryGetValue(kvp.Key, out Typeface face) &&
                            !Owner.RowTypefaces.TryGetValue(kvp.Key.Item2, out face) &&
                            !Owner.ColumnTypefaces.TryGetValue(kvp.Key.Item1, out face))
                    {
                        face = defaultTypeFace;
                    }

                    if (!Owner.CellMargin.TryGetValue(kvp.Key, out Thickness margin))
                    {
                        margin = Owner.DefaultMargin;
                    }

                    FormattedText fmtText = new FormattedText(kvp.Value, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, face, this.Container.FontSize, this.Foreground);

                    found = true;
                    maxHeight = Math.Max(maxHeight, fmtText.Height + margin.Top + margin.Bottom);
                }
            }

            if (found)
            {
                double originalHeight = this.Owner.GetHeight(clickedRow, out bool wasDefault);

                if (originalHeight != maxHeight)
                {
                    Owner.RowHeights[clickedRow] = maxHeight;
                    Owner.InvalidateVisual();
                    this.InvalidateVisual();

                    this.Container.UndoStackRowHeight.Push(new ValueStackFrame<int, double>(this.Container.Selection, new Dictionary<int, Reference<double>>() { { clickedRow, wasDefault ? null : new Reference<double>(originalHeight) } }, new Dictionary<int, Reference<double>>() { { clickedRow, maxHeight } }));
                    this.Container.UndoStack.Push(null);

                    this.Container.UndoStackRowForeground.Push(null);
                    this.Container.UndoStackRowTypeface.Push(null);

                    this.Container.UndoStackColumnForeground.Push(null);
                    this.Container.UndoStackColumnWidth.Push(null);
                    this.Container.UndoStackColumnTypeface.Push(null);

                    this.Container.UndoStackCellForeground.Push(null);
                    this.Container.UndoStackCellMargin.Push(null);
                    this.Container.UndoStackCellHorizontalAlignment.Push(null);
                    this.Container.UndoStackCellTypeface.Push(null);
                    this.Container.UndoStackCellVerticalAlignment.Push(null);

                    this.Container.ClearRedoStack();
                }
            }
            else
            {
                double originalHeight = this.Owner.GetHeight(clickedRow, out bool wasDefault);

                if (!wasDefault)
                {
                    Owner.RowHeights.Remove(clickedRow);
                    Owner.InvalidateVisual();
                    this.InvalidateVisual();

                    this.Container.UndoStackRowHeight.Push(new ValueStackFrame<int, double>(this.Container.Selection, new Dictionary<int, Reference<double>>() { { clickedRow, originalHeight } }, new Dictionary<int, Reference<double>>() { { clickedRow, null } }));
                    this.Container.UndoStack.Push(null);

                    this.Container.UndoStackRowForeground.Push(null);
                    this.Container.UndoStackRowTypeface.Push(null);

                    this.Container.UndoStackColumnForeground.Push(null);
                    this.Container.UndoStackColumnWidth.Push(null);
                    this.Container.UndoStackColumnTypeface.Push(null);

                    this.Container.UndoStackCellForeground.Push(null);
                    this.Container.UndoStackCellMargin.Push(null);
                    this.Container.UndoStackCellHorizontalAlignment.Push(null);
                    this.Container.UndoStackCellTypeface.Push(null);
                    this.Container.UndoStackCellVerticalAlignment.Push(null);

                    this.Container.ClearRedoStack();
                }
            }
        }

        internal Reference<double> AutoFitHeightWithoutStack(int clickedRow)
        {
            bool found = false;
            double maxHeight = 10;

            Typeface defaultTypeFace = new Typeface(Owner.FontFamily, Owner.FontStyle, Owner.FontWeight);

            foreach (KeyValuePair<(int, int), string> kvp in Owner.Data)
            {
                if (kvp.Key.Item2 == clickedRow)
                {
                    if (!Owner.CellTypefaces.TryGetValue(kvp.Key, out Typeface face) &&
                            !Owner.RowTypefaces.TryGetValue(kvp.Key.Item2, out face) &&
                            !Owner.ColumnTypefaces.TryGetValue(kvp.Key.Item1, out face))
                    {
                        face = defaultTypeFace;
                    }

                    if (!Owner.CellMargin.TryGetValue(kvp.Key, out Thickness margin))
                    {
                        margin = Owner.DefaultMargin;
                    }

                    FormattedText fmtText = new FormattedText(kvp.Value, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, face, this.Container.FontSize, this.Foreground);

                    found = true;
                    maxHeight = Math.Max(maxHeight, fmtText.Height + margin.Top + margin.Bottom);
                }
            }

            if (found)
            {
                double originalHeight = this.Owner.GetHeight(clickedRow, out _);

                if (originalHeight != maxHeight)
                {
                    Owner.RowHeights[clickedRow] = maxHeight;
                }

                return maxHeight;
            }
            else
            {
                this.Owner.GetHeight(clickedRow, out bool wasDefault);

                if (!wasDefault)
                {
                    Owner.RowHeights.Remove(clickedRow);
                }

                return null;
            }
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            if (double.IsInfinity(availableSize.Height))
            {
                return new Size(100, 1000);
            }

            (int top, double topDelta, int height, double actualHeight, double startHeight) = Owner.GetRangeY(this.Offset.Y, availableSize.Height);

            string max = new string('9', (int)Math.Ceiling(Math.Log10(top + height)));

            FormattedText fmt = new FormattedText(max, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(this.FontFamily), this.FontSize, this.Foreground);

            if (this.FontFamily != null)
            {
                return new Size(fmt.Width + 10, availableSize.Height);
            }
            else
            {
                return new Size(max.Length * this.FontSize, availableSize.Height);
            }
        }

        protected override void OnPointerExited(PointerEventArgs e)
        {
            base.OnPointerExited(e);

            if (hoverRow >= 0)
            {
                hoverRow = -1;
                this.InvalidateVisual();
            }
        }

        private double[] lastDrawnYs = null;
        private int lastDrawnTop = -1;
        private int hoverRow = -1;
        private double lastDrawnDelta = 0;

        public override void Render(DrawingContext context)
        {
            if (this.FontFamily == null)
            {
                return;
            }

            context.FillRectangle(new SolidColorBrush(this.HeaderBackground), new Rect(0, 0, this.Bounds.Width, this.Bounds.Height));

            (int left, double offsetX, int top, double offsetY, int width, double actualWidth, double startWidth, int height, double actualHeight, double startHeight) = Owner.GetRange(Owner.Offset.X, Owner.Offset.Y, this.Bounds.Width, this.Bounds.Height);

            lastDrawnTop = top;
            lastDrawnDelta = offsetY;

            Pen gridPen = new Pen(new LinearGradientBrush() { StartPoint = new RelativePoint(0, 0, RelativeUnit.Absolute), EndPoint = new RelativePoint(this.Bounds.Width, 0, RelativeUnit.Absolute), GradientStops = new GradientStops() { new GradientStop(this.HeaderBackground, 0), new GradientStop(this.GridColor, 1) } });
            IBrush brs = this.Foreground;
            Brush selectionHighlightBrush = new SolidColorBrush(Color.FromArgb(16, 0, 0, 0));
            Brush fullSelectionHighlightBrush = new SolidColorBrush(SelectionAccent.Color, 0.2);
            Brush hoverHighlightBrush = new SolidColorBrush(SelectionAccent.Color, 0.35);
            Pen selectionPen = new Pen(SelectionAccent, 2);

            Typeface typeFace = new Typeface(this.FontFamily, this.FontStyle, this.FontWeight);

            using (context.PushTransform(Matrix.CreateTranslation(0, offsetY)))
            {
                double[] ys = new double[height + 1];

                int[] selected = new int[height + 1];
                bool?[] shouldBeAccented = new bool?[height + 1];

                for (int i = 0; i < Selection.Count; i++)
                {
                    for (int j = Math.Max(Selection[i].Top, top); j <= Math.Min(Selection[i].Bottom, top + height); j++)
                    {
                        if (Selection[i].Left == 0 && Selection[i].Right == Owner.MaxTableWidth)
                        {
                            selected[j - top] = Math.Max(selected[j - top], 2);
                        }
                        else
                        {
                            selected[j - top] = Math.Max(selected[j - top], 1);
                        }

                        if (Selection.Count == 1 && Selection[i].Left < left && Selection[i].Right >= left)
                        {
                            shouldBeAccented[j - top] = false;
                        }
                        else if (shouldBeAccented[j - top] == null)
                        {
                            shouldBeAccented[j - top] = true;
                        }
                    }
                }

                double currY = 0;
                for (int y = 0; y <= height; y++)
                {
                    context.DrawLine(gridPen, new Point(0, currY).SnapToDevicePixels(this), new Point(this.Bounds.Width, currY).SnapToDevicePixels(this));
                    currY += Owner.GetHeight(top + y);
                    ys[y] = currY;
                }

                lastDrawnYs = ys;

                for (int y = 0; y <= height; y++)
                {
                    if (hoverRow == y + top)
                    {
                        context.FillRectangle(hoverHighlightBrush, new Rect(new Point(0, y == 0 ? 0 : ys[y - 1]).SnapToDevicePixels(this, true, true), new Point(this.Bounds.Width, ys[y]).SnapToDevicePixels(this, true, true)));
                    }
                    else
                    {
                        if (selected[y] == 1)
                        {
                            context.FillRectangle(selectionHighlightBrush, new Rect(new Point(0, y == 0 ? 0 : ys[y - 1]).SnapToDevicePixels(this, true, true), new Point(this.Bounds.Width, ys[y]).SnapToDevicePixels(this, true, true)));
                        }
                        else if (selected[y] == 2)
                        {
                            context.FillRectangle(fullSelectionHighlightBrush, new Rect(new Point(0, y == 0 ? 0 : ys[y - 1]).SnapToDevicePixels(this, true, true), new Point(this.Bounds.Width, ys[y]).SnapToDevicePixels(this, true, true)));
                        }
                    }
                }

                for (int y = 0; y <= height; y++)
                {
                    string txt = (y + top + 1).ToString();

                    double realY = y == 0 ? 0 : ys[y - 1];


                    using (context.PushClip(new Rect(3, realY + 3, this.Bounds.Width, ys[y] - realY - 6)))
                    {
                        FormattedText fmtText = new FormattedText(txt, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeFace, this.FontSize, selected[y] == 0 ? brs : SelectionAccent);
                        double realX = (this.Bounds.Width - fmtText.Width) * 0.5;
                        context.DrawText(fmtText, new Point(realX, realY + (ys[y] - realY) * 0.5 - fmtText.Height * 0.5));
                    }
                }

                context.DrawLine(this.HeaderSeparator, new Point(this.Bounds.Width - 1, -offsetY).SnapToDevicePixels(this), new Point(this.Bounds.Width - 1, this.Bounds.Height - offsetY).SnapToDevicePixels(this));

                for (int y = 0; y <= height; y++)
                {
                    if (selected[y] > 0 && shouldBeAccented[y] == true)
                    {
                        context.DrawLine(selectionPen, new Point(this.Bounds.Width - 1, (y == 0 ? 0 : ys[y - 1])).SnapToDevicePixels(this, true, true), new Point(this.Bounds.Width - 1, ys[y] + 1).SnapToDevicePixels(this, true, true));
                    }
                }
            }
        }
    }
}
