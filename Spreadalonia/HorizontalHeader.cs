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
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace Spreadalonia
{
    internal class HorizontalHeader : Control
    {
        public static readonly StyledProperty<Spreadsheet> ContainerProperty = AvaloniaProperty.Register<HorizontalHeader, Spreadsheet>(nameof(Container));
        public Spreadsheet Container
        {
            get { return GetValue(ContainerProperty); }
            set { SetValue(ContainerProperty, value); }
        }

        public static readonly StyledProperty<Table> OwnerProperty = AvaloniaProperty.Register<HorizontalHeader, Table>(nameof(Owner));
        public Table Owner
        {
            get { return GetValue(OwnerProperty); }
            set { SetValue(OwnerProperty, value); }
        }

        public static readonly StyledProperty<Vector> OffsetProperty = AvaloniaProperty.Register<HorizontalHeader, Vector>(nameof(Offset), new Vector(0, 0));
        public Vector Offset
        {
            get { return GetValue(OffsetProperty); }
            set { SetValue(OffsetProperty, value); }
        }

        public static readonly StyledProperty<Color> GridColorProperty = AvaloniaProperty.Register<HorizontalHeader, Color>(nameof(GridColor), Color.FromRgb(220, 220, 220));
        public Color GridColor
        {
            get { return GetValue(GridColorProperty); }
            set { SetValue(GridColorProperty, value); }
        }

        public static readonly StyledProperty<Color> HeaderBackgroundProperty = AvaloniaProperty.Register<HorizontalHeader, Color>(nameof(HeaderBackground), Color.FromRgb(240, 240, 240));
        public Color HeaderBackground
        {
            get { return GetValue(HeaderBackgroundProperty); }
            set { SetValue(HeaderBackgroundProperty, value); }
        }

        public static readonly StyledProperty<Pen> HeaderSeparatorProperty = AvaloniaProperty.Register<HorizontalHeader, Pen>(nameof(HeaderSeparator), new Pen(new SolidColorBrush(Color.FromRgb(170, 170, 170))));
        public Pen HeaderSeparator
        {
            get { return GetValue(HeaderSeparatorProperty); }
            set { SetValue(HeaderSeparatorProperty, value); }
        }

        public static readonly StyledProperty<ImmutableList<SelectionRange>> SelectionProperty = AvaloniaProperty.Register<HorizontalHeader, ImmutableList<SelectionRange>>(nameof(Selection), ImmutableList.Create<SelectionRange>());
        public ImmutableList<SelectionRange> Selection
        {
            get { return GetValue(SelectionProperty); }
            set { SetValue(SelectionProperty, value); }
        }

        public static readonly StyledProperty<SolidColorBrush> SelectionAccentProperty = AvaloniaProperty.Register<HorizontalHeader, SolidColorBrush>(nameof(SelectionAccent), new SolidColorBrush(Color.FromRgb(0, 114, 176)));
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


        static HorizontalHeader()
        {
            AffectsRender<HorizontalHeader>(OffsetProperty, GridColorProperty, HeaderBackgroundProperty, SelectionProperty, SelectionAccentProperty);
        }

        public HorizontalHeader()
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

        protected override Size MeasureOverride(Size availableSize)
        {
            FormattedText fmt = new FormattedText("ABCDEFGHIJKLMNOPQRSTUVWXYZ", CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(this.FontFamily), this.FontSize, Foreground);

            if (this.FontFamily != null)
            {
                return new Size(availableSize.Width, fmt.Height + 7);
            }
            else
            {
                return new Size(availableSize.Width, this.FontSize + 7);
            }
        }

        protected override void OnPointerExited(PointerEventArgs e)
        {
            base.OnPointerExited(e);

            if (hoverColumn >= 0)
            {
                hoverColumn = -1;
                this.InvalidateVisual();
            }
        }

        protected override void OnPointerReleased(PointerReleasedEventArgs e)
        {
            base.OnPointerReleased(e);

            PointerPoint currentPoint = e.GetCurrentPoint(this);

            double xPos = currentPoint.Position.X - lastDrawnDelta;

            if (pointerPressedAction == 2)
            {
                double newWidth = Math.Max(10, originalColumnWidth + xPos - dragStartX);

                if (newWidth != originalColumnWidth)
                {
                    this.Container.UndoStackColumnWidth.Push(new ValueStackFrame<int, double>(this.Container.Selection, new Dictionary<int, Reference<double>>() { { selectionStartColumn, wasOriginalColumnWidthDefault ? null : new Reference<double>(originalColumnWidth) } }, new Dictionary<int, Reference<double>>() { { selectionStartColumn, newWidth } }));
                    this.Container.UndoStack.Push(null);

                    this.Container.UndoStackRowForeground.Push(null);
                    this.Container.UndoStackRowHeight.Push(null);
                    this.Container.UndoStackRowTypeface.Push(null);

                    this.Container.UndoStackColumnForeground.Push(null);
                    this.Container.UndoStackColumnTypeface.Push(null);

                    this.Container.UndoStackCellForeground.Push(null);
                    this.Container.UndoStackCellMargin.Push(null);
                    this.Container.UndoStackCellHorizontalAlignment.Push(null);
                    this.Container.UndoStackCellTypeface.Push(null);
                    this.Container.UndoStackCellVerticalAlignment.Push(null);

                    this.Container.ClearRedoStack();

                    if (this.Container.Selection.Count > 0 && this.Container.Selection[0].Left == selectionStartColumn)
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
            int clickedColumn = -1;

            double xPos = currentPoint.Position.X - lastDrawnDelta;

            if (lastDrawnXs != null)
            {
                for (int i = 0; i < lastDrawnXs.Length; i++)
                {
                    if (lastDrawnXs[i] > xPos && xPos >= (i == 0 ? 0 : lastDrawnXs[i - 1]))
                    {
                        clickedColumn = i + lastDrawnLeft;
                    }

                    closest = Math.Min(closest, Math.Abs(lastDrawnXs[i] - xPos));

                    if (closest <= 5 && clickedColumn >= 0)
                    {
                        break;
                    }
                }
            }

            if (clickedColumn >= 0)
            {
                if (!currentPoint.Properties.IsLeftButtonPressed)
                {
                    if (closest <= 5)
                    {
                        this.Cursor = Cursors.ResizeColumnCursor((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1);
                    }
                    else
                    {
                        this.Cursor = Cursors.SelectColumnCursor((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1);
                    }

                    if (hoverColumn != clickedColumn)
                    {
                        hoverColumn = clickedColumn;
                        this.InvalidateVisual();
                    }

                }
                else
                {
                    if (pointerPressedAction == 1)
                    {
                        if (selectionMode == 0)
                        {
                            this.Container.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(Math.Min(selectionStartColumn, clickedColumn), 0, Math.Max(selectionStartColumn, clickedColumn), Owner.MaxTableHeight));
                        }
                        else if (selectionMode == 1)
                        {
                            this.Container.Selection = previousSelection.Difference(new SelectionRange(Math.Min(selectionStartColumn, clickedColumn), 0, Math.Max(selectionStartColumn, clickedColumn), Owner.MaxTableHeight));
                        }
                    }
                    else if (pointerPressedAction == 2)
                    {
                        double newWidth = Math.Max(10, originalColumnWidth + xPos - dragStartX);

                        Owner.ColumnWidths[selectionStartColumn] = newWidth;
                        Owner.InvalidateVisual();
                        this.InvalidateVisual();
                    }
                }
            }
        }

        int pointerPressedAction = 0;
        int selectionStartColumn = -1;
        int selectionMode = 0;
        double dragStartX = 0;
        double originalColumnWidth = 0;
        bool wasOriginalColumnWidthDefault = false;
        ImmutableList<SelectionRange> previousSelection;

        protected override void OnPointerPressed(PointerPressedEventArgs e)
        {
            base.OnPointerPressed(e);

            double closest = double.MaxValue;
            int clickedColumn = -1;

            PointerPoint currentPoint = e.GetCurrentPoint(this);

            double xPos = currentPoint.Position.X - lastDrawnDelta;

            if (lastDrawnXs != null)
            {
                for (int i = 0; i < lastDrawnXs.Length; i++)
                {
                    if (lastDrawnXs[i] > xPos && xPos >= (i == 0 ? 0 : lastDrawnXs[i - 1]))
                    {
                        clickedColumn = i + lastDrawnLeft;
                    }

                    closest = Math.Min(closest, Math.Abs(lastDrawnXs[i] - xPos));

                    if (closest <= 5 && clickedColumn >= 0)
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
                            if (this.Selection[i].Top == 0 && this.Selection[i].Bottom == Owner.MaxTableHeight && this.Selection[i].Left <= clickedColumn && this.Selection[i].Right >= clickedColumn)
                            {
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            this.Container.Selection = this.Container.Selection.Add(new SelectionRange(clickedColumn, 0, clickedColumn, Owner.MaxTableHeight));
                            selectionStartColumn = clickedColumn;
                            selectionMode = 0;
                        }
                        else
                        {
                            previousSelection = this.Container.Selection;
                            this.Container.Selection = this.Container.Selection.Difference(new SelectionRange(clickedColumn, 0, clickedColumn, Owner.MaxTableHeight));
                            selectionStartColumn = clickedColumn;
                            selectionMode = 1;
                        }
                    }
                    else if (e.KeyModifiers == KeyModifiers.Shift && this.Selection.Count > 0)
                    {
                        int prevStartColumn = this.Selection[this.Selection.Count - 1].Left;
                        int prevEndColumn = this.Selection[this.Selection.Count - 1].Right;

                        if (clickedColumn >= prevStartColumn)
                        {
                            selectionStartColumn = prevStartColumn;
                        }
                        else
                        {
                            selectionStartColumn = prevEndColumn;
                        }

                        this.Container.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(Math.Min(selectionStartColumn, clickedColumn), 0, Math.Max(selectionStartColumn, clickedColumn), Owner.MaxTableHeight));
                        selectionMode = 0;
                    }
                    else
                    {
                        this.Container.Selection = ImmutableList.Create(new SelectionRange(clickedColumn, 0, clickedColumn, Owner.MaxTableHeight));
                        selectionStartColumn = clickedColumn;
                        selectionMode = 0;
                    }

                    this.Cursor = Cursors.CrossCursor((this.VisualRoot as ILayoutRoot)?.LayoutScaling ?? 1);
                    pointerPressedAction = 1;
                }
                else
                {
                    if (xPos - (clickedColumn > 0 ? lastDrawnXs[clickedColumn - 1] : 0) <= 5)
                    {
                        clickedColumn--;
                    }

                    if (e.ClickCount == 1)
                    {
                        pointerPressedAction = 2;

                        selectionStartColumn = clickedColumn;

                        originalColumnWidth = Owner.GetWidth(clickedColumn, out wasOriginalColumnWidthDefault);
                        dragStartX = xPos;
                    }
                    else if (e.ClickCount == 2)
                    {
                        AutoFitWidth(clickedColumn);
                        if (this.Container.Selection.Count > 0 && this.Container.Selection[0].Left == clickedColumn)
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
                    if (this.Selection[i].Top == 0 && this.Selection[i].Bottom == Owner.MaxTableHeight && this.Selection[i].Left <= clickedColumn && this.Selection[i].Right >= clickedColumn)
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    this.Container.Selection = ImmutableList.Create(new SelectionRange(clickedColumn, 0, clickedColumn, Owner.MaxTableHeight));
                }
            }
        }

        private void AutoFitWidth(int column)
        {
            bool found = false;
            double maxWidth = 10;

            Typeface defaultTypeFace = new Typeface(Owner.FontFamily, Owner.FontStyle, Owner.FontWeight);

            foreach (KeyValuePair<(int, int), string> kvp in Owner.Data)
            {
                if (kvp.Key.Item1 == column)
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

                    FormattedText fmtText = new FormattedText(kvp.Value, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, face, this.Container.FontSize, Foreground);

                    double textWidth = fmtText.Width;

                    if (this.Container.ShowColorPreview && kvp.Value.StartsWith("#") && (kvp.Value.Length == 7 || kvp.Value.Length == 9))
                    {
                        try
                        {
                            Brush.Parse(kvp.Value.Length == 7 ? kvp.Value : ("#" + kvp.Value.Substring(7, 2) + kvp.Value.Substring(1, 6)));
                            textWidth += this.Owner.FontSize + 3;
                        }
                        catch { }
                    }

                    found = true;
                    maxWidth = Math.Max(maxWidth, textWidth + margin.Left + margin.Right);
                }
            }

            if (found)
            {
                double originalWidth = this.Owner.GetWidth(column, out bool wasDefault);

                if (originalWidth != maxWidth)
                {
                    Owner.ColumnWidths[column] = maxWidth;
                    Owner.InvalidateVisual();
                    this.InvalidateVisual();

                    this.Container.UndoStackColumnWidth.Push(new ValueStackFrame<int, double>(this.Container.Selection, new Dictionary<int, Reference<double>>() { { column, wasDefault ? null : new Reference<double>(originalWidth) } }, new Dictionary<int, Reference<double>>() { { column, maxWidth } }));
                    this.Container.UndoStack.Push(null);

                    this.Container.UndoStackRowForeground.Push(null);
                    this.Container.UndoStackRowHeight.Push(null);
                    this.Container.UndoStackRowTypeface.Push(null);

                    this.Container.UndoStackColumnForeground.Push(null);
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
                double originalWidth = this.Owner.GetWidth(column, out bool wasDefault);

                if (!wasDefault)
                {
                    Owner.ColumnWidths.Remove(column);
                    Owner.InvalidateVisual();
                    this.InvalidateVisual();

                    this.Container.UndoStackColumnWidth.Push(new ValueStackFrame<int, double>(this.Container.Selection, new Dictionary<int, Reference<double>>() { { column, originalWidth } }, new Dictionary<int, Reference<double>>() { { column, null } }));
                    this.Container.UndoStack.Push(null);

                    this.Container.UndoStackRowForeground.Push(null);
                    this.Container.UndoStackRowHeight.Push(null);
                    this.Container.UndoStackRowTypeface.Push(null);

                    this.Container.UndoStackColumnForeground.Push(null);
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

        internal Reference<double> AutoFitWidthWithoutStack(int column)
        {
            bool found = false;
            double maxWidth = 10;

            Typeface defaultTypeFace = new Typeface(Owner.FontFamily, Owner.FontStyle, Owner.FontWeight);

            foreach (KeyValuePair<(int, int), string> kvp in Owner.Data)
            {
                if (kvp.Key.Item1 == column)
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

                    FormattedText fmtText = new FormattedText(kvp.Value, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, face, this.Container.FontSize, Foreground);

                    double textWidth = fmtText.Width;

                    if (this.Container.ShowColorPreview && kvp.Value.StartsWith("#") && (kvp.Value.Length == 7 || kvp.Value.Length == 9))
                    {
                        try
                        {
                            Brush.Parse(kvp.Value.Length == 7 ? kvp.Value : ("#" + kvp.Value.Substring(7, 2) + kvp.Value.Substring(1, 6)));
                            textWidth += this.Owner.FontSize + 3;
                        }
                        catch { }
                    }

                    found = true;
                    maxWidth = Math.Max(maxWidth, textWidth + margin.Left + margin.Right);
                }
            }

            if (found)
            {
                double originalWidth = this.Owner.GetWidth(column, out bool wasDefault);

                if (originalWidth != maxWidth)
                {
                    Owner.ColumnWidths[column] = maxWidth;
                }

                return maxWidth;
            }
            else
            {
                this.Owner.GetWidth(column, out bool wasDefault);

                if (!wasDefault)
                {
                    Owner.ColumnWidths.Remove(column);
                }

                return null;
            }
        }

        internal static string GetLetter(int x)
        {
            if (x < 26)
            {
                return ((char)(x + 65)).ToString();
            }
            else
            {
                StringBuilder tbr = new StringBuilder();

                while (x >= 26)
                {
                    tbr.Insert(0, (char)((x % 26) + 65));
                    x = x / 26;
                }

                tbr.Insert(0, (char)(x + 64));

                return tbr.ToString();
            }
        }

        internal static int FromLetter(string letter)
        {
            int tbr = 0;

            for (int i = 0; i < letter.Length; i++)
            {
                int v = ((int)letter[letter.Length - i - 1] - 65) + 1;

                for (int j = 0; j < i; j++)
                {
                    v *= 26;
                }

                tbr += v;
            }

            return tbr;
        }

        private double[] lastDrawnXs = null;
        private int lastDrawnLeft = -1;
        private int hoverColumn = -1;
        double lastDrawnDelta = 0;

        public override void Render(DrawingContext context)
        {
            if (this.FontFamily == null)
            {
                return;
            }

            context.FillRectangle(new SolidColorBrush(this.HeaderBackground), new Rect(0, 0, this.Bounds.Width, this.Bounds.Height));

            (int left, double offsetX, int top, double offsetY, int width, double actualWidth, double startWidth, int height, double actualHeight, double startHeight) = Owner.GetRange(Owner.Offset.X, Owner.Offset.Y, this.Bounds.Width, this.Bounds.Height);

            lastDrawnLeft = left;
            lastDrawnDelta = offsetX;

            Pen gridPen = new Pen(new LinearGradientBrush() { StartPoint = new RelativePoint(0, 0, RelativeUnit.Absolute), EndPoint = new RelativePoint(0, this.Bounds.Height, RelativeUnit.Absolute), GradientStops = new GradientStops() { new GradientStop(this.HeaderBackground, 0), new GradientStop(this.GridColor, 1) } });
            IBrush brs = this.Foreground;
            Brush selectionHighlightBrush = new SolidColorBrush(Color.FromArgb(16, 0, 0, 0));
            Brush fullSelectionHighlightBrush = new SolidColorBrush(SelectionAccent.Color, 0.2);
            Brush hoverHighlightBrush = new SolidColorBrush(SelectionAccent.Color, 0.35);
            Pen selectionPen = new Pen(SelectionAccent, 2);

            Typeface typeFace = new Typeface(this.FontFamily, this.FontStyle, this.FontWeight);

            using (context.PushPreTransform(Matrix.CreateTranslation(offsetX, 0)))
            {
                double[] xs = new double[width + 1];
                int[] selected = new int[width + 1];
                bool?[] shouldBeAccented = new bool?[width + 1];

                for (int i = 0; i < Selection.Count; i++)
                {
                    for (int j = Math.Max(Selection[i].Left, left); j <= Math.Min(Selection[i].Right, left + width); j++)
                    {
                        if (Selection[i].Top == 0 && Selection[i].Bottom == Owner.MaxTableHeight)
                        {
                            selected[j - left] = Math.Max(selected[j - left], 2);
                        }
                        else
                        {
                            selected[j - left] = Math.Max(selected[j - left], 1);
                        }


                        if (Selection.Count == 1 && Selection[i].Top < top && Selection[i].Bottom >= top)
                        {
                            shouldBeAccented[j - left] = false;
                        }
                        else if (shouldBeAccented[j - left] == null)
                        {
                            shouldBeAccented[j - left] = true;
                        }
                    }
                }

                double currX = 0;
                for (int x = 0; x <= width; x++)
                {
                    context.DrawLine(gridPen, new Point(currX, 0).SnapToDevicePixels(this), new Point(currX, this.Bounds.Height).SnapToDevicePixels(this));

                    currX += Owner.GetWidth(left + x);
                    xs[x] = currX;
                }

                lastDrawnXs = xs;

                for (int x = 0; x <= width; x++)
                {
                    if (hoverColumn == x + left)
                    {
                        context.FillRectangle(hoverHighlightBrush, new Rect(new Point(x == 0 ? 0 : xs[x - 1], 0).SnapToDevicePixels(this, true, true), new Point(xs[x], this.Bounds.Height).SnapToDevicePixels(this, true, true)));
                    }
                    else
                    {
                        if (selected[x] == 1)
                        {
                            context.FillRectangle(selectionHighlightBrush, new Rect(new Point(x == 0 ? 0 : xs[x - 1], 0).SnapToDevicePixels(this, true, true), new Point(xs[x], this.Bounds.Height).SnapToDevicePixels(this, true, true)));
                        }
                        else if (selected[x] == 2)
                        {
                            context.FillRectangle(fullSelectionHighlightBrush, new Rect(new Point(x == 0 ? 0 : xs[x - 1], 0).SnapToDevicePixels(this, true, true), new Point(xs[x], this.Bounds.Height).SnapToDevicePixels(this, true, true)));
                        }
                    }
                }


                for (int x = 0; x <= width; x++)
                {
                    string txt = GetLetter(x + left);

                    double realX = x == 0 ? 0 : xs[x - 1];

                    

                    using (context.PushClip(new Rect(realX + 3, 0, xs[x] - realX - 6, this.Bounds.Height)))
                    { 
                        FormattedText fmtText = new FormattedText(txt, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeFace, this.FontSize, selected[x] == 0 ? brs : SelectionAccent);
                        
                        double realY = this.Bounds.Height - 4 - fmtText.Height;

                        context.DrawText(fmtText, new Point(realX + (xs[x] - realX) * 0.5 - fmtText.Width * 0.5, realY));
                    }
                }

                context.DrawLine(this.HeaderSeparator, new Point(-offsetX, this.Bounds.Height - 1).SnapToDevicePixels(this), new Point(this.Bounds.Width - offsetX, this.Bounds.Height - 1).SnapToDevicePixels(this));

                for (int x = 0; x <= width; x++)
                {
                    if (selected[x] > 0 && shouldBeAccented[x] == true)
                    {
                        context.DrawLine(selectionPen, new Point((x == 0 ? 0 : xs[x - 1]), this.Bounds.Height - 1).SnapToDevicePixels(this, true, true), new Point(xs[x] + 1, this.Bounds.Height - 1).SnapToDevicePixels(this, true, true));
                    }
                }
            }


        }
    }
}
