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
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Fare;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Spreadalonia
{
    /// <summary>
    /// A simple spreadsheet control for Avalonia.
    /// </summary>
    public partial class Spreadsheet : UserControl
    {
        internal static KeyModifiers ControlModifier = System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX) ? KeyModifiers.Meta : KeyModifiers.Control;

        /// <summary>
        /// Raised when the size of a cell changes.
        /// </summary>
        public event EventHandler<CellSizeChangedEventArgs> CellSizeChanged;

        internal void RaiseCellSizeChanged()
        {
            if (this.Selection.Count > 0)
            {
                int x = this.Selection[0].Left;
                int y = this.Selection[0].Top;
                (double w, double h) = this.GetCellSize(x, y);

                CellSizeChanged?.Invoke(this, new CellSizeChangedEventArgs(x, y, w, h));
            }
        }

        /// <summary>
        /// Raised when the user double taps on a colour preview.
        /// </summary>
        public event EventHandler<ColorDoubleTappedEventArgs> ColorDoubleTapped;

        internal bool RaiseColorDoubleTapped((int, int) cell, Color color)
        {
            ColorDoubleTappedEventArgs e = new ColorDoubleTappedEventArgs(cell.Item1, cell.Item2, color);
            ColorDoubleTapped?.Invoke(this, e);
            return e.Handled;
        }

        /// <summary>
        /// Defines the <see cref="Data"/> property.
        /// </summary>
        public static readonly DirectProperty<Spreadsheet, Dictionary<(int, int), string>> DataProperty = AvaloniaProperty.RegisterDirect<Spreadsheet, Dictionary<(int, int), string>>(nameof(HeaderFontFamily), o => o.Data);

        /// <summary>
        /// The data contained in the spreadsheet. Please do not change this directly, or you will mess up the Undo/Redo stack.
        /// </summary>
        public Dictionary<(int, int), string> Data
        {
            get
            {
                if (this.IsInitialized)
                {
                    return this.FindControl<Table>("ContentTable").Data;
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Defines the <see cref="DefaultTextAlignment"/> property.
        /// </summary>
        public static readonly StyledProperty<TextAlignment> DefaultTextAlignmentProperty = AvaloniaProperty.Register<Spreadsheet, TextAlignment>(nameof(DefaultTextAlignment), TextAlignment.Left);

        /// <summary>
        /// The default horizontal text alignment for spreadsheet cells.
        /// </summary>
        public TextAlignment DefaultTextAlignment
        {
            get { return GetValue(DefaultTextAlignmentProperty); }
            set { SetValue(DefaultTextAlignmentProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="DefaultVerticalAlignment"/> property.
        /// </summary>
        public static readonly StyledProperty<VerticalAlignment> DefaultVerticalAlignmentProperty = AvaloniaProperty.Register<Spreadsheet, VerticalAlignment>(nameof(DefaultVerticalAlignment), VerticalAlignment.Center);

        /// <summary>
        /// The default vertical text alignment for spreadsheet cells.
        /// </summary>
        public VerticalAlignment DefaultVerticalAlignment
        {
            get { return GetValue(DefaultVerticalAlignmentProperty); }
            set { SetValue(DefaultVerticalAlignmentProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="DefaultRowHeight"/> property.
        /// </summary>
        public static readonly StyledProperty<double> DefaultRowHeightProperty = AvaloniaProperty.Register<Spreadsheet, double>(nameof(DefaultRowHeight), 23);

        /// <summary>
        /// The default height for spreadsheet rows.
        /// </summary>
        public double DefaultRowHeight
        {
            get { return GetValue(DefaultRowHeightProperty); }
            set { SetValue(DefaultRowHeightProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="DefaultColumnWidth"/> property.
        /// </summary>
        public static readonly StyledProperty<double> DefaultColumnWidthProperty = AvaloniaProperty.Register<Spreadsheet, double>(nameof(DefaultColumnWidth), 65);

        /// <summary>
        /// The default width for spreadsheet columns.
        /// </summary>
        public double DefaultColumnWidth
        {
            get { return GetValue(DefaultColumnWidthProperty); }
            set { SetValue(DefaultColumnWidthProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="DefaultColumnWidth"/> property.
        /// </summary>
        public static readonly StyledProperty<Thickness> DefaultCellMarginProperty = AvaloniaProperty.Register<Spreadsheet, Thickness>(nameof(DefaultCellMargin), new Thickness(3));

        /// <summary>
        /// The default width for spreadsheet columns.
        /// </summary>
        public Thickness DefaultCellMargin
        {
            get { return GetValue(DefaultCellMarginProperty); }
            set { SetValue(DefaultCellMarginProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="HeaderFontFamily"/> property.
        /// </summary>
        public static readonly StyledProperty<FontFamily> HeaderFontFamilyProperty = AvaloniaProperty.Register<Spreadsheet, FontFamily>(nameof(HeaderFontFamily), FontFamily.Default);

        /// <summary>
        /// The font family used for the row and column headers.
        /// </summary>
        public FontFamily HeaderFontFamily
        {
            get { return GetValue(HeaderFontFamilyProperty); }
            set { SetValue(HeaderFontFamilyProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="HeaderFontSize"/> property.
        /// </summary>
        public static readonly StyledProperty<double> HeaderFontSizeProperty = AvaloniaProperty.Register<Spreadsheet, double>(nameof(HeaderFontSize), 14);

        /// <summary>
        /// The font size used for the row and column headers.
        /// </summary>
        public double HeaderFontSize
        {
            get { return GetValue(HeaderFontSizeProperty); }
            set { SetValue(HeaderFontSizeProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="HeaderForeground"/> property.
        /// </summary>
        public static readonly StyledProperty<IBrush> HeaderForegroundProperty = AvaloniaProperty.Register<Spreadsheet, IBrush>(nameof(HeaderForeground), Brushes.Black);

        /// <summary>
        /// The text colour for the row and column headers.
        /// </summary>
        public IBrush HeaderForeground
        {
            get { return GetValue(HeaderForegroundProperty); }
            set { SetValue(HeaderForegroundProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="HeaderBackground"/> property.
        /// </summary>
        public static readonly StyledProperty<Color> HeaderBackgroundProperty = AvaloniaProperty.Register<Spreadsheet, Color>(nameof(HeaderBackground), Color.FromRgb(240, 240, 240));

        /// <summary>
        /// The background of the row and column headers.
        /// </summary>
        public Color HeaderBackground
        {
            get { return GetValue(HeaderBackgroundProperty); }
            set { SetValue(HeaderBackgroundProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="GridColor"/> property.
        /// </summary>
        public static readonly StyledProperty<Color> GridColorProperty = AvaloniaProperty.Register<Spreadsheet, Color>(nameof(GridColor), Color.FromRgb(220, 220, 220));

        /// <summary>
        /// The colour used to draw the grid lines.
        /// </summary>
        public Color GridColor
        {
            get { return GetValue(GridColorProperty); }
            set { SetValue(GridColorProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="Selection"/> property.
        /// </summary>
        public static readonly StyledProperty<ImmutableList<SelectionRange>> SelectionProperty = AvaloniaProperty.Register<Spreadsheet, ImmutableList<SelectionRange>>(nameof(Selection), ImmutableList.Create<SelectionRange>(new SelectionRange(0, 0)));

        /// <summary>
        /// The currently selected cells.
        /// </summary>
        public ImmutableList<SelectionRange> Selection
        {
            get { return GetValue(SelectionProperty); }
            set { SetValue(SelectionProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="SpreadsheetBackground"/> property.
        /// </summary>
        public static readonly StyledProperty<SolidColorBrush> GridBackgroundProperty = AvaloniaProperty.Register<Spreadsheet, SolidColorBrush>(nameof(SpreadsheetBackground), new SolidColorBrush(Colors.White));

        /// <summary>
        /// The background colour of the spreadsheet.
        /// </summary>
        public SolidColorBrush SpreadsheetBackground
        {
            get { return GetValue(GridBackgroundProperty); }
            set { SetValue(GridBackgroundProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="SelectionAccent"/> property.
        /// </summary>
        public static readonly StyledProperty<SolidColorBrush> SelectionAccentProperty = AvaloniaProperty.Register<Spreadsheet, SolidColorBrush>(nameof(SelectionAccent), new SolidColorBrush(Color.FromRgb(0, 114, 176)));

        /// <summary>
        /// The colour used to highlight the selected cells.
        /// </summary>
        public SolidColorBrush SelectionAccent
        {
            get { return GetValue(SelectionAccentProperty); }
            set { SetValue(SelectionAccentProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="IsEditing"/> property.
        /// </summary>
        public static readonly StyledProperty<bool> IsEditingProperty = AvaloniaProperty.Register<Spreadsheet, bool>(nameof(IsEditing), false);

        /// <summary>
        /// Returns <see langword="true"/> if the user is editing the value of a cell, <see langword="false"/> otherwise.
        /// </summary>
        public bool IsEditing
        {
            get { return GetValue(IsEditingProperty); }
            internal set { SetValue(IsEditingProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="ColumnSeparator"/> property.
        /// </summary>
        public static readonly StyledProperty<Regex> ColumnSeparatorProperty = AvaloniaProperty.Register<Spreadsheet, Regex>(nameof(ColumnSeparator), new Regex("\\t", RegexOptions.Compiled));

        /// <summary>
        /// The character used to separate columns (e.g., tab or comma).
        /// </summary>
        public Regex ColumnSeparator
        {
            get { return GetValue(ColumnSeparatorProperty); }
            set { SetValue(ColumnSeparatorProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="RowSeparator"/> property.
        /// </summary>
        public static readonly StyledProperty<Regex> RowSeparatorProperty = AvaloniaProperty.Register<Spreadsheet, Regex>(nameof(RowSeparator), new Regex("\n", RegexOptions.Compiled));

        /// <summary>
        /// The character used to separate rows (e.g., the newline character).
        /// </summary>
        public Regex RowSeparator
        {
            get { return GetValue(RowSeparatorProperty); }
            set { SetValue(RowSeparatorProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="QuoteSymbol"/> property.
        /// </summary>
        public static readonly StyledProperty<string> QuoteSymbolProperty = AvaloniaProperty.Register<Spreadsheet, string>(nameof(QuoteSymbol), "\"");

        /// <summary>
        /// The character used to quote cells containing literal values corresponding to the row or column separators (e.g., single or double quotes).
        /// </summary>
        public string QuoteSymbol
        {
            get { return GetValue(QuoteSymbolProperty); }
            set { SetValue(QuoteSymbolProperty, value); }
        }


        /// <summary>
        /// Defines the <see cref="MaxTableWidth"/> property.
        /// </summary>
        public static readonly StyledProperty<int> MaxTableWidthProperty = AvaloniaProperty.Register<Spreadsheet, int>(nameof(MaxTableWidth), int.MaxValue - 2);

        /// <summary>
        /// The maximum width of the spreadsheet.
        /// </summary>
        public int MaxTableWidth
        {
            get { return GetValue(MaxTableWidthProperty); }
            set { SetValue(MaxTableWidthProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="MaxTableHeight"/> property.
        /// </summary>
        public static readonly StyledProperty<int> MaxTableHeightProperty = AvaloniaProperty.Register<Spreadsheet, int>(nameof(MaxTableHeight), int.MaxValue - 2);

        /// <summary>
        /// The maximum width of the spreadsheet.
        /// </summary>
        public int MaxTableHeight
        {
            get { return GetValue(MaxTableHeightProperty); }
            set { SetValue(MaxTableHeightProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="ShowColorPreview"/> property.
        /// </summary>
        public static readonly StyledProperty<bool> ShowColorPreviewProperty = AvaloniaProperty.Register<Spreadsheet, bool>(nameof(ShowColorPreview), true);

        /// <summary>
        /// If this is <see langword="true"/>, a small square showing the colour preview is shown in cells containing hex colour representations (<c>#RRGGBB</c> or <c>#RRGGBBAA</c>).
        /// </summary>
        public bool ShowColorPreview
        {
            get { return GetValue(ShowColorPreviewProperty); }
            set { SetValue(ShowColorPreviewProperty, value); }
        }

        /// <summary>
        /// Defines the <see cref="CanUndo"/> property.
        /// </summary>
        public static readonly DirectProperty<Spreadsheet, bool> CanUndoProperty = AvaloniaProperty.RegisterDirect<Spreadsheet, bool>(nameof(CanUndo), o => o.CanUndo);

        private bool _canUndo = false;

        /// <summary>
        /// Determines whether it is possible to undo the last action that was performed on the spreadsheet.
        /// </summary>
        public bool CanUndo
        {
            get { return _canUndo; }
            private set { SetAndRaise(CanUndoProperty, ref _canUndo, value); }
        }

        /// <summary>
        /// Defines the <see cref="CanRedo"/> property.
        /// </summary>
        public static readonly DirectProperty<Spreadsheet, bool> CanRedoProperty = AvaloniaProperty.RegisterDirect<Spreadsheet, bool>(nameof(CanRedo), o => o.CanRedo);

        private bool _canRedo = false;

        /// <summary>
        /// Determines whether it is possible to Redo the last action that was performed on the spreadsheet.
        /// </summary>
        public bool CanRedo
        {
            get { return _canRedo; }
            private set { SetAndRaise(CanRedoProperty, ref _canRedo, value); }
        }



        internal (int, int) EditingCell { get; set; }

        internal Stack<StackFrame<(int, int), string>> UndoStack { get; } = new Stack<StackFrame<(int, int), string>>();
        internal Stack<StackFrame<(int, int), IBrush>> UndoStackCellForeground { get; } = new Stack<StackFrame<(int, int), IBrush>>();
        internal Stack<ValueStackFrame<(int, int), Thickness>> UndoStackCellMargin { get; } = new Stack<ValueStackFrame<(int, int), Thickness>>();
        internal Stack<ValueStackFrame<(int, int), TextAlignment>> UndoStackCellHorizontalAlignment { get; } = new Stack<ValueStackFrame<(int, int), TextAlignment>>();
        internal Stack<ValueStackFrame<(int, int), VerticalAlignment>> UndoStackCellVerticalAlignment { get; } = new Stack<ValueStackFrame<(int, int), VerticalAlignment>>();
        internal Stack<ValueStackFrame<(int, int), Typeface>> UndoStackCellTypeface { get; } = new Stack<ValueStackFrame<(int, int), Typeface>>();
        internal Stack<StackFrame<int, IBrush>> UndoStackRowForeground { get; } = new Stack<StackFrame<int, IBrush>>();
        internal Stack<StackFrame<int, IBrush>> UndoStackColumnForeground { get; } = new Stack<StackFrame<int, IBrush>>();
        internal Stack<ValueStackFrame<int, Typeface>> UndoStackRowTypeface { get; } = new Stack<ValueStackFrame<int, Typeface>>();
        internal Stack<ValueStackFrame<int, Typeface>> UndoStackColumnTypeface { get; } = new Stack<ValueStackFrame<int, Typeface>>();
        internal Stack<ValueStackFrame<int, double>> UndoStackRowHeight { get; } = new Stack<ValueStackFrame<int, double>>();
        internal Stack<ValueStackFrame<int, double>> UndoStackColumnWidth { get; } = new Stack<ValueStackFrame<int, double>>();


        internal Stack<StackFrame<(int, int), string>> RedoStack { get; } = new Stack<StackFrame<(int, int), string>>();
        internal Stack<StackFrame<(int, int), IBrush>> RedoStackCellForeground { get; } = new Stack<StackFrame<(int, int), IBrush>>();
        internal Stack<ValueStackFrame<(int, int), Thickness>> RedoStackCellMargin { get; } = new Stack<ValueStackFrame<(int, int), Thickness>>();
        internal Stack<ValueStackFrame<(int, int), TextAlignment>> RedoStackCellHorizontalAlignment { get; } = new Stack<ValueStackFrame<(int, int), TextAlignment>>();
        internal Stack<ValueStackFrame<(int, int), VerticalAlignment>> RedoStackCellVerticalAlignment { get; } = new Stack<ValueStackFrame<(int, int), VerticalAlignment>>();
        internal Stack<ValueStackFrame<(int, int), Typeface>> RedoStackCellTypeface { get; } = new Stack<ValueStackFrame<(int, int), Typeface>>();
        internal Stack<StackFrame<int, IBrush>> RedoStackRowForeground { get; } = new Stack<StackFrame<int, IBrush>>();
        internal Stack<StackFrame<int, IBrush>> RedoStackColumnForeground { get; } = new Stack<StackFrame<int, IBrush>>();
        internal Stack<ValueStackFrame<int, Typeface>> RedoStackRowTypeface { get; } = new Stack<ValueStackFrame<int, Typeface>>();
        internal Stack<ValueStackFrame<int, Typeface>> RedoStackColumnTypeface { get; } = new Stack<ValueStackFrame<int, Typeface>>();
        internal Stack<ValueStackFrame<int, double>> RedoStackRowHeight { get; } = new Stack<ValueStackFrame<int, double>>();
        internal Stack<ValueStackFrame<int, double>> RedoStackColumnWidth { get; } = new Stack<ValueStackFrame<int, double>>();

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        /// <summary>
        /// Create a new <see cref="Spreadsheet"/> instance.
        /// </summary>
        public Spreadsheet()
        {
            InitializeComponent();

            this.FindControl<HorizontalHeader>("HorizontalHeaderControl").FontFamily = this.HeaderFontFamily;
            this.FindControl<VerticalHeader>("VerticalHeaderControl").FontFamily = this.HeaderFontFamily;

            this.FindControl<HorizontalHeader>("HorizontalHeaderControl").Owner = this.FindControl<Table>("ContentTable");
            this.FindControl<VerticalHeader>("VerticalHeaderControl").Owner = this.FindControl<Table>("ContentTable");
            this.FindControl<TopLeftCorner>("TopLeftCornerControl").Owner = this.FindControl<Table>("ContentTable");
            this.FindControl<Overlay>("IconOverlay").Owner = this.FindControl<Table>("ContentTable");

            this.FindControl<ScrollBar>("HorizontalScrollBar").SmallChange = this.FindControl<Table>("ContentTable").DefaultColumnWidth;
            this.FindControl<ScrollBar>("HorizontalScrollBar").LargeChange = this.FindControl<Table>("ContentTable").DefaultColumnWidth * 3;
            this.FindControl<ScrollBar>("VerticalScrollBar").SmallChange = this.FindControl<Table>("ContentTable").DefaultRowHeight;
            this.FindControl<ScrollBar>("VerticalScrollBar").LargeChange = this.FindControl<Table>("ContentTable").DefaultRowHeight * 3;

            this.FindControl<Canvas>("CutIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.Cut")));
            this.FindControl<Canvas>("CopyIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.CopyDoc")));
            this.FindControl<Canvas>("PasteIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.PasteDoc")));

            this.FindControl<Canvas>("InsertColIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.InsertColumn")));
            this.FindControl<Canvas>("DeleteColIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.DeleteColumn")));
            this.FindControl<Canvas>("InsertRowIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.InsertRow")));
            this.FindControl<Canvas>("DeleteRowIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.DeleteRow")));

            this.FindControl<Canvas>("ClearContentsIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.ClearContents")));
            this.FindControl<Canvas>("ClearFormatIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.ClearFormat")));

            this.FindControl<Canvas>("AutoWidthIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.AutoWidth")));
            this.FindControl<Canvas>("AutoHeightIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.AutoHeight")));

            this.FindControl<Canvas>("ResetWidthIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.ResetWidth")));
            this.FindControl<Canvas>("ResetHeightIcon").Children.Add(new DPIAwareBox(Icons.GetIcon16("Spreadalonia.Icons.ResetHeight")));


            this.FindControl<ScrollBar>("HorizontalScrollBar").PropertyChanged += (s, e) =>
            {
                if (e.Property == ScrollBar.ValueProperty)
                {
                    Vector newOffset = new Avalonia.Vector(this.FindControl<ScrollBar>("HorizontalScrollBar").Value, this.FindControl<ScrollBar>("VerticalScrollBar").Value);

                    if (scrollBarXBusy)
                    {
                        this.FindControl<Table>("ContentTable").Offset = newOffset;
                        this.FindControl<HorizontalHeader>("HorizontalHeaderControl").Offset = newOffset;
                        this.FindControl<VerticalHeader>("VerticalHeaderControl").Offset = newOffset;
                        this.FindControl<Overlay>("IconOverlay").Offset = newOffset;
                    }
                    else
                    {
                        this.FindControl<Table>("ContentTable").PauseTransitions();
                        this.FindControl<HorizontalHeader>("HorizontalHeaderControl").PauseTransitions();
                        this.FindControl<VerticalHeader>("VerticalHeaderControl").PauseTransitions();
                        this.FindControl<Table>("ContentTable").Offset = newOffset;
                        this.FindControl<HorizontalHeader>("HorizontalHeaderControl").Offset = newOffset;
                        this.FindControl<VerticalHeader>("VerticalHeaderControl").Offset = newOffset;
                        this.FindControl<Overlay>("IconOverlay").Offset = newOffset;
                        this.FindControl<Table>("ContentTable").ResumeTransitions();
                        this.FindControl<HorizontalHeader>("HorizontalHeaderControl").ResumeTransitions();
                        this.FindControl<VerticalHeader>("VerticalHeaderControl").ResumeTransitions();
                    }
                }
            };

            this.FindControl<ScrollBar>("VerticalScrollBar").PropertyChanged += (s, e) =>
            {
                if (e.Property == ScrollBar.ValueProperty)
                {
                    Vector newOffset = new Avalonia.Vector(this.FindControl<ScrollBar>("HorizontalScrollBar").Value, this.FindControl<ScrollBar>("VerticalScrollBar").Value);

                    if (scrollBarYBusy)
                    {
                        this.FindControl<Table>("ContentTable").Offset = newOffset;
                        this.FindControl<HorizontalHeader>("HorizontalHeaderControl").Offset = newOffset;
                        this.FindControl<VerticalHeader>("VerticalHeaderControl").Offset = newOffset;
                        this.FindControl<Overlay>("IconOverlay").Offset = newOffset;
                    }
                    else
                    {
                        this.FindControl<Table>("ContentTable").PauseTransitions();
                        this.FindControl<HorizontalHeader>("HorizontalHeaderControl").PauseTransitions();
                        this.FindControl<VerticalHeader>("VerticalHeaderControl").PauseTransitions();
                        this.FindControl<Table>("ContentTable").Offset = newOffset;
                        this.FindControl<HorizontalHeader>("HorizontalHeaderControl").Offset = newOffset;
                        this.FindControl<VerticalHeader>("VerticalHeaderControl").Offset = newOffset;
                        this.FindControl<Overlay>("IconOverlay").Offset = newOffset;
                        this.FindControl<Table>("ContentTable").ResumeTransitions();
                        this.FindControl<HorizontalHeader>("HorizontalHeaderControl").ResumeTransitions();
                        this.FindControl<VerticalHeader>("VerticalHeaderControl").ResumeTransitions();
                    }
                }
            };

            this.FindControl<ScrollBar>("HorizontalScrollBar").AddHandler(PointerPressedEvent, (s, e) =>
            {
                scrollBarXBusy = true;

                if ((e.Pointer.Captured as Visual).Name == "ThumbVisual")
                {
                    this.FindControl<Table>("ContentTable").PauseTransitions();
                    this.FindControl<HorizontalHeader>("HorizontalHeaderControl").PauseTransitions();
                    this.FindControl<VerticalHeader>("VerticalHeaderControl").PauseTransitions();
                }

            }, handledEventsToo: true);

            this.FindControl<ScrollBar>("VerticalScrollBar").AddHandler(PointerPressedEvent, (s, e) =>
            {
                scrollBarYBusy = true;

                if ((e.Pointer.Captured as Visual).Name == "ThumbVisual")
                {
                    this.FindControl<Table>("ContentTable").PauseTransitions();
                    this.FindControl<HorizontalHeader>("HorizontalHeaderControl").PauseTransitions();
                    this.FindControl<VerticalHeader>("VerticalHeaderControl").PauseTransitions();
                }
            }, handledEventsToo: true);

            this.FindControl<ScrollBar>("HorizontalScrollBar").AddHandler(PointerReleasedEvent, (s, e) =>
            {
                scrollBarXBusy = false;

                if ((e.Pointer.Captured as Visual).Name == "ThumbVisual")
                {
                    this.FindControl<Table>("ContentTable").ResumeTransitions();
                    this.FindControl<HorizontalHeader>("HorizontalHeaderControl").ResumeTransitions();
                    this.FindControl<VerticalHeader>("VerticalHeaderControl").ResumeTransitions();
                }

                if (delayedScrollBarMaximum != null)
                {
                    (double, double) scrbMax = delayedScrollBarMaximum.Value;
                    delayedScrollBarMaximum = null;
                    SetScrollbarMaximum(scrbMax.Item1, scrbMax.Item2);
                }
            }, handledEventsToo: true);

            this.FindControl<ScrollBar>("VerticalScrollBar").AddHandler(PointerReleasedEvent, (s, e) =>
            {
                scrollBarYBusy = false;

                if ((e.Pointer.Captured as Visual).Name == "ThumbVisual")
                {
                    this.FindControl<Table>("ContentTable").ResumeTransitions();
                    this.FindControl<HorizontalHeader>("HorizontalHeaderControl").ResumeTransitions();
                    this.FindControl<VerticalHeader>("VerticalHeaderControl").ResumeTransitions();
                }

                if (delayedScrollBarMaximum != null)
                {
                    (double, double) scrbMax = delayedScrollBarMaximum.Value;
                    delayedScrollBarMaximum = null;
                    SetScrollbarMaximum(scrbMax.Item1, scrbMax.Item2);
                }
            }, handledEventsToo: true);

            this.FindControl<TextBox>("EditingBox").PropertyChanged += (s, e) =>
            {
                if (e.Property == TextBox.TextProperty)
                {
                    FormattedText txt = new FormattedText(this.FindControl<TextBox>("EditingBox").Text,
                        new Typeface(this.FindControl<TextBox>("EditingBox").FontFamily, this.FindControl<TextBox>("EditingBox").FontStyle, this.FindControl<TextBox>("EditingBox").FontWeight),
                        this.FindControl<TextBox>("EditingBox").FontSize, this.FindControl<TextBox>("EditingBox").TextAlignment, this.FindControl<TextBox>("EditingBox").TextWrapping, new Size(double.PositiveInfinity, double.PositiveInfinity));

                    this.FindControl<TextBox>("EditingBox").Width = txt.Bounds.Width + 1;
                    this.FindControl<TextBox>("EditingBox").Height = txt.Bounds.Height;
                }
            };

            this.FindControl<TextBox>("EditingBox").LostFocus += (s, e) =>
            {
                if (this.IsEditing)
                {
                    string newText = this.FindControl<TextBox>("EditingBox").Text;

                    bool present = this.FindControl<Table>("ContentTable").Data.TryGetValue(this.EditingCell, out string prevVal);

                    if ((present && prevVal != newText) || (!present && !string.IsNullOrEmpty(newText)))
                    {

                        if (!string.IsNullOrEmpty(newText))
                        {
                            this.UndoStack.Push(new StackFrame<(int, int), string>(ImmutableList.Create(new SelectionRange(this.EditingCell)), new Dictionary<(int, int), string>() { { this.EditingCell, prevVal } }, new Dictionary<(int, int), string>() { { this.EditingCell, newText } }));
                            this.FindControl<Table>("ContentTable").Data[this.EditingCell] = newText;
                        }
                        else
                        {
                            this.UndoStack.Push(new StackFrame<(int, int), string>(ImmutableList.Create(new SelectionRange(this.EditingCell)), new Dictionary<(int, int), string>() { { this.EditingCell, prevVal } }, new Dictionary<(int, int), string>() { { this.EditingCell, null } }));
                            this.FindControl<Table>("ContentTable").Data.Remove(this.EditingCell);
                        }

                        this.PushNonDataStackNull();
                        this.ClearRedoStack();
                    }

                    this.IsEditing = false;
                    this.FindControl<Table>("ContentTable").InvalidateVisual();
                }
            };

            this.FindControl<TextBox>("EditingBox").KeyDown += (s, e) =>
            {
                if (e.Key == Key.Return)
                {
                    if (e.KeyModifiers == KeyModifiers.Alt)
                    {
                        int selectionStart = this.FindControl<TextBox>("EditingBox").SelectionStart;
                        int selectionEnd = this.FindControl<TextBox>("EditingBox").SelectionEnd;
                        string text = this.FindControl<TextBox>("EditingBox").Text;

                        this.FindControl<TextBox>("EditingBox").Text = text.Substring(0, selectionStart) + "\n" + text.Substring(selectionEnd);
                        this.FindControl<TextBox>("EditingBox").SelectionStart = selectionStart + 1;
                        this.FindControl<TextBox>("EditingBox").SelectionEnd = selectionStart + 1;
                        e.Handled = true;
                    }
                    else if (e.KeyModifiers == KeyModifiers.None)
                    {
                        this.Selection = ImmutableList.Create(new SelectionRange(EditingCell.Item1, EditingCell.Item2 + 1));
                        this.FindControl<Table>("ContentTable").Focus();
                        e.Handled = true;
                    }
                    else if (e.KeyModifiers == KeyModifiers.Shift)
                    {
                        this.Selection = ImmutableList.Create(new SelectionRange(EditingCell.Item1, Math.Max(0, EditingCell.Item2 - 1)));
                        this.FindControl<Table>("ContentTable").Focus();
                        e.Handled = true;
                    }
                }
                else if (e.Key == Key.Tab)
                {
                    if (e.KeyModifiers == KeyModifiers.Shift)
                    {
                        this.FindControl<Table>("ContentTable").Focus();
                        this.Selection = ImmutableList.Create(new SelectionRange(Math.Max(0, EditingCell.Item1 - 1), EditingCell.Item2, Math.Max(0, EditingCell.Item1 - 1), EditingCell.Item2));
                        e.Handled = true;
                    }
                    else if (e.KeyModifiers == KeyModifiers.None)
                    {
                        this.FindControl<Table>("ContentTable").Focus();
                        this.Selection = ImmutableList.Create(new SelectionRange(EditingCell.Item1 + 1, EditingCell.Item2, EditingCell.Item1 + 1, EditingCell.Item2));
                        e.Handled = true;
                    }

                }
                else if (e.Key == Key.Escape)
                {
                    if (!this.FindControl<Table>("ContentTable").Data.TryGetValue(this.EditingCell, out string txt))
                    {
                        txt = null;
                    }

                    this.FindControl<TextBox>("EditingBox").Text = txt;
                    this.FindControl<Table>("ContentTable").Focus();
                    e.Handled = true;
                }
            };

            this.FindControl<ContextMenu>("TableContextMenu").ContextMenuOpening += async (s, e) =>
            {
                this.FindControl<MenuItem>("CopyMenuItem").IsEnabled = this.Selection.Count > 0;
                this.FindControl<MenuItem>("CutMenuItem").IsEnabled = this.Selection.Count > 0;
                this.FindControl<MenuItem>("PasteMenuItem").IsEnabled = this.Selection.Count == 1 && (await Application.Current.Clipboard.GetFormatsAsync()).Contains("Text");
                this.FindControl<MenuItem>("PasteSkipBlanksMenuItem").IsEnabled = this.Selection.Count == 1 && (await Application.Current.Clipboard.GetFormatsAsync()).Contains("Text");

                this.FindControl<MenuItem>("InsertColMenuItem").IsVisible = this.Selection.Count == 1 && (this.Selection[0].IsColumns(this.FindControl<Table>("ContentTable")) && !this.Selection[0].IsRows(this.FindControl<Table>("ContentTable")));
                this.FindControl<MenuItem>("DeleteColMenuItem").IsVisible = this.FindControl<MenuItem>("InsertColMenuItem").IsVisible;

                this.FindControl<MenuItem>("InsertRowMenuItem").IsVisible = this.Selection.Count == 1 && (this.Selection[0].IsRows(this.FindControl<Table>("ContentTable")) && !this.Selection[0].IsColumns(this.FindControl<Table>("ContentTable")));
                this.FindControl<MenuItem>("DeleteRowMenuItem").IsVisible = this.FindControl<MenuItem>("InsertRowMenuItem").IsVisible;

                this.FindControl<MenuItem>("ClearMenuItem").IsEnabled = this.Selection.Count > 0;
                this.FindControl<MenuItem>("ResetFormatMenuItem").IsEnabled = this.Selection.Count > 0;

                this.FindControl<MenuItem>("AutoHeightMenuItem").IsVisible = this.Selection.Count > 0 && this.Selection.All(x => x.IsRows(this.FindControl<Table>("ContentTable")) && !x.IsColumns(this.FindControl<Table>("ContentTable")));
                this.FindControl<MenuItem>("ResetHeightMenuItem").IsVisible = this.FindControl<MenuItem>("AutoHeightMenuItem").IsVisible;

                this.FindControl<MenuItem>("AutoWidthMenuItem").IsVisible = this.Selection.Count > 0 && this.Selection.All(x => x.IsColumns(this.FindControl<Table>("ContentTable")) && !x.IsRows(this.FindControl<Table>("ContentTable")));
                this.FindControl<MenuItem>("ResetWidthMenuItem").IsVisible = this.FindControl<MenuItem>("AutoWidthMenuItem").IsVisible;

                this.FindControl<MenuItem>("LastSeparator").IsVisible = this.FindControl<MenuItem>("AutoWidthMenuItem").IsVisible || this.FindControl<MenuItem>("AutoHeightMenuItem").IsVisible;
            };

            this.FindControl<MenuItem>("CopyMenuItem").Click += (s, e) => Copy();
            this.FindControl<MenuItem>("CutMenuItem").Click += (s, e) => Cut();
            this.FindControl<MenuItem>("PasteMenuItem").Click += async (s, e) => await Paste(true);
            this.FindControl<MenuItem>("PasteSkipBlanksMenuItem").Click += async (s, e) => await Paste(false);

            this.FindControl<MenuItem>("ClearMenuItem").Click += (s, e) => ClearContents();

            this.FindControl<MenuItem>("InsertColMenuItem").Click += (s, e) =>
            {
                InsertColumns();
            };

            this.FindControl<MenuItem>("DeleteColMenuItem").Click += (s, e) =>
            {
                DeleteColumns();
            };

            this.FindControl<MenuItem>("InsertRowMenuItem").Click += (s, e) =>
            {
                InsertRows();
            };

            this.FindControl<MenuItem>("DeleteRowMenuItem").Click += (s, e) =>
            {
                DeleteRows();
            };

            this.FindControl<MenuItem>("ResetFormatMenuItem").Click += (s, e) =>
            {
                ResetFormat();
            };

            this.FindControl<MenuItem>("AutoWidthMenuItem").Click += (s, e) =>
            {
                AutoFitWidth();
            };

            this.FindControl<MenuItem>("ResetWidthMenuItem").Click += (s, e) =>
            {
                ResetWidth();
            };

            this.FindControl<MenuItem>("AutoHeightMenuItem").Click += (s, e) =>
            {
                AutoFitHeight();
            };

            this.FindControl<MenuItem>("ResetHeightMenuItem").Click += (s, e) =>
            {
                ResetHeight();
            };
        }

        /// <inheritdoc/>
        protected override void OnPropertyChanged<T>(AvaloniaPropertyChangedEventArgs<T> change)
        {
            base.OnPropertyChanged(change);

            if (change.Property == IsEditingProperty)
            {
                if (this.IsEditing)
                {
                    this.FindControl<ScrollBar>("HorizontalScrollBar").IsEnabled = false;
                    this.FindControl<ScrollBar>("VerticalScrollBar").IsEnabled = false;

                    Table table = this.FindControl<Table>("ContentTable");

                    double left = EditingCell.Item1 - table.lastDrawnLeft > 0 ? table.lastDrawnXs[EditingCell.Item1 - table.lastDrawnLeft - 1] : 0;
                    double top = EditingCell.Item2 - table.lastDrawnTop > 0 ? table.lastDrawnYs[EditingCell.Item2 - table.lastDrawnTop - 1] : 0;

                    double width = table.lastDrawnXs[EditingCell.Item1 - table.lastDrawnLeft] - left;
                    double height = table.lastDrawnYs[EditingCell.Item2 - table.lastDrawnTop] - top - 1;

                    left += table.lastDrawnDeltaX;
                    top += table.lastDrawnDeltaY;

                    if (!table.Data.TryGetValue((EditingCell.Item1, EditingCell.Item2), out string txt))
                    {
                        txt = "";
                    }

                    if (!table.CellTypefaces.TryGetValue(EditingCell, out Typeface face) &&
                                        !table.RowTypefaces.TryGetValue(EditingCell.Item2, out face) &&
                                        !table.ColumnTypefaces.TryGetValue(EditingCell.Item1, out face))
                    {
                        face = new Typeface(table.FontFamily, table.FontStyle, table.FontWeight); ;
                    }

                    if (!table.CellForeground.TryGetValue(EditingCell, out IBrush brs) &&
                        !table.RowForeground.TryGetValue(EditingCell.Item2, out brs) &&
                        !table.ColumnForeground.TryGetValue(EditingCell.Item1, out brs))
                    {
                        brs = table.Foreground;
                    }

                    if (!table.CellTextAlignment.TryGetValue(EditingCell, out TextAlignment hor))
                    {
                        hor = table.DefaultTextAlignment;
                    }

                    if (!table.CellVerticalAlignment.TryGetValue(EditingCell, out VerticalAlignment ver))
                    {
                        ver = table.DefaultVerticalAlignment;
                    }

                    if (!table.CellMargin.TryGetValue(EditingCell, out Thickness margin))
                    {
                        margin = table.DefaultMargin;
                    }

                    this.FindControl<TextBox>("EditingBox").FontFamily = face.FontFamily;
                    this.FindControl<TextBox>("EditingBox").FontStyle = face.Style;
                    this.FindControl<TextBox>("EditingBox").FontWeight = face.Weight;
                    this.FindControl<TextBox>("EditingBox").FontSize = this.FontSize;

                    this.FindControl<TextBox>("EditingBox").Foreground = brs;

                    this.FindControl<TextBox>("EditingBox").TextAlignment = hor;

                    this.FindControl<TextBox>("EditingBox").VerticalContentAlignment = ver;

                    Point pt = new Point(left + margin.Left - 1, top + margin.Top - 1);

                    this.FindControl<TextBox>("EditingBox").Margin = new Thickness(pt.X, pt.Y, 0, 0);

                    this.FindControl<TextBox>("EditingBox").MinWidth = width - margin.Left - margin.Right;
                    this.FindControl<TextBox>("EditingBox").MinHeight = height - margin.Top - margin.Bottom;

                    this.FindControl<TextBox>("EditingBox").Text = txt;
                    this.FindControl<TextBox>("EditingBox").SelectionStart = txt.Length;
                    this.FindControl<TextBox>("EditingBox").SelectionEnd = txt.Length;
                    this.FindControl<TextBox>("EditingBox").Focus();
                    this.FindControl<TextBox>("EditingBox").IsVisible = true;
                }
                else
                {
                    this.FindControl<ScrollBar>("HorizontalScrollBar").IsEnabled = true;
                    this.FindControl<ScrollBar>("VerticalScrollBar").IsEnabled = true;
                    this.FindControl<TextBox>("EditingBox").IsVisible = false;
                }
            }
            else if (change.Property == SelectionProperty)
            {
                ImmutableList<SelectionRange> oldValue = change.OldValue.GetValueOrDefault<ImmutableList<SelectionRange>>();
                ImmutableList<SelectionRange> newValue = change.NewValue.GetValueOrDefault<ImmutableList<SelectionRange>>();

                if (newValue != null && newValue.Count > 0)
                {
                    SelectionRange newRange = newValue.Merge();

                    Table table = this.FindControl<Table>("ContentTable");

                    Rect topLeft = table.GetCoordinates(newRange.Left, newRange.Top);
                    Rect bottomRight = topLeft;

                    if (newRange.IsFinite(table))
                    {
                        bottomRight = table.GetCoordinates(newRange.Right, newRange.Bottom);
                    }
                    else if (newRange.IsColumns(table) && !newRange.IsRows(table))
                    {
                        bottomRight = table.GetCoordinates(newRange.Right, newRange.Top);
                    }
                    else if (!newRange.IsColumns(table) && newRange.IsRows(table))
                    {
                        bottomRight = table.GetCoordinates(newRange.Left, newRange.Bottom);
                    }

                    if (oldValue != null && oldValue.Count > 0)
                    {
                        SelectionRange oldRange = oldValue.Merge();

                        if (newRange.Left != oldRange.Left)
                        {
                            if (table.lastDrawnLeft > newRange.Left - 1 || table.lastDrawnLeft + table.lastDrawnWidth < newRange.Left + 3)
                            {
                                if (table.lastDrawnLeft > newRange.Left - 1)
                                {
                                    SetScrollbarMaximum(Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, topLeft.Right - table.Bounds.Width + 20), this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
                                    this.FindControl<ScrollBar>("HorizontalScrollBar").Value = topLeft.Left;
                                }
                                else
                                {
                                    SetScrollbarMaximum(Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, topLeft.Right - table.Bounds.Width + 20), this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
                                    this.FindControl<ScrollBar>("HorizontalScrollBar").Value = topLeft.Right - table.Bounds.Width + 20;
                                }
                            }
                        }
                        else if (newRange.Right != oldRange.Right && (newRange.IsFinite(table) || (newRange.IsColumns(table) && !newRange.IsRows(table))))
                        {
                            if (table.lastDrawnLeft > newRange.Right - 1 || table.lastDrawnLeft + table.lastDrawnWidth < newRange.Right + 3)
                            {
                                if (table.lastDrawnLeft > newRange.Right - 1)
                                {
                                    SetScrollbarMaximum(Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, bottomRight.Right - table.Bounds.Width + 20), this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
                                    this.FindControl<ScrollBar>("HorizontalScrollBar").Value = bottomRight.Left;
                                }
                                else
                                {
                                    SetScrollbarMaximum(Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, bottomRight.Right - table.Bounds.Width + 20), this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
                                    this.FindControl<ScrollBar>("HorizontalScrollBar").Value = bottomRight.Right - table.Bounds.Width + 20;
                                }
                            }
                        }

                        if (newRange.Top != oldRange.Top)
                        {
                            if (table.lastDrawnTop > newRange.Top - 1 || table.lastDrawnTop + table.lastDrawnHeight < newRange.Top + 3)
                            {
                                if (table.lastDrawnTop > newRange.Top - 1)
                                {
                                    SetScrollbarMaximum(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Maximum, topLeft.Bottom - table.Bounds.Height));
                                    this.FindControl<ScrollBar>("VerticalScrollBar").Value = topLeft.Top;
                                }
                                else
                                {
                                    SetScrollbarMaximum(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Maximum, topLeft.Bottom - table.Bounds.Height));
                                    this.FindControl<ScrollBar>("VerticalScrollBar").Value = topLeft.Bottom - table.Bounds.Height;
                                }
                            }
                        }
                        else if (newRange.Bottom != oldRange.Bottom && (newRange.IsFinite(table) || (!newRange.IsColumns(table) && newRange.IsRows(table))))
                        {
                            if (table.lastDrawnTop > newRange.Bottom - 1 || table.lastDrawnTop + table.lastDrawnHeight < newRange.Bottom + 3)
                            {
                                if (table.lastDrawnTop > newRange.Bottom - 1)
                                {
                                    SetScrollbarMaximum(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Maximum, bottomRight.Bottom - table.Bounds.Height));
                                    this.FindControl<ScrollBar>("VerticalScrollBar").Value = bottomRight.Top;
                                }
                                else
                                {
                                    SetScrollbarMaximum(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Maximum, bottomRight.Bottom - table.Bounds.Height));
                                    this.FindControl<ScrollBar>("VerticalScrollBar").Value = bottomRight.Bottom - table.Bounds.Height;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (newRange.IsFinite(table) || (newRange.IsColumns(table) && !newRange.IsRows(table)))
                        {
                            if (table.lastDrawnLeft > newRange.Right - 1 || table.lastDrawnLeft + table.lastDrawnWidth < newRange.Right + 3)
                            {
                                if (table.lastDrawnLeft > newRange.Right - 1)
                                {
                                    SetScrollbarMaximum(Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, bottomRight.Right - table.Bounds.Width + 20), this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
                                    this.FindControl<ScrollBar>("HorizontalScrollBar").Value = bottomRight.Left;
                                }
                                else
                                {
                                    SetScrollbarMaximum(Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, bottomRight.Right - table.Bounds.Width + 20), this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
                                    this.FindControl<ScrollBar>("HorizontalScrollBar").Value = bottomRight.Right - table.Bounds.Width + 20;
                                }
                            }
                        }

                        if (newRange.IsFinite(table) || (!newRange.IsColumns(table) && newRange.IsRows(table)))
                        {
                            if (table.lastDrawnTop > newRange.Bottom - 1 || table.lastDrawnTop + table.lastDrawnHeight < newRange.Bottom + 3)
                            {
                                if (table.lastDrawnTop > newRange.Bottom - 1)
                                {
                                    SetScrollbarMaximum(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Maximum, bottomRight.Bottom - table.Bounds.Height));
                                    this.FindControl<ScrollBar>("VerticalScrollBar").Value = bottomRight.Top;
                                }
                                else
                                {
                                    SetScrollbarMaximum(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Maximum, bottomRight.Bottom - table.Bounds.Height));
                                    this.FindControl<ScrollBar>("VerticalScrollBar").Value = bottomRight.Bottom - table.Bounds.Height;
                                }
                            }
                        }

                        if (table.lastDrawnLeft > newRange.Left - 1 || table.lastDrawnLeft + table.lastDrawnWidth < newRange.Left + 3)
                        {
                            if (table.lastDrawnLeft > newRange.Left - 1)
                            {
                                SetScrollbarMaximum(Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, topLeft.Right - table.Bounds.Width + 20), this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
                                this.FindControl<ScrollBar>("HorizontalScrollBar").Value = topLeft.Left;
                            }
                            else
                            {
                                SetScrollbarMaximum(Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, topLeft.Right - table.Bounds.Width + 20), this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
                                this.FindControl<ScrollBar>("HorizontalScrollBar").Value = topLeft.Right - table.Bounds.Width + 20;
                            }
                        }

                        if (table.lastDrawnTop > newRange.Top - 1 || table.lastDrawnTop + table.lastDrawnHeight < newRange.Top + 3)
                        {
                            if (table.lastDrawnTop > newRange.Top - 1)
                            {
                                SetScrollbarMaximum(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Maximum, topLeft.Bottom - table.Bounds.Height));
                                this.FindControl<ScrollBar>("VerticalScrollBar").Value = topLeft.Top;
                            }
                            else
                            {
                                SetScrollbarMaximum(this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum, Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Maximum, topLeft.Bottom - table.Bounds.Height));
                                this.FindControl<ScrollBar>("VerticalScrollBar").Value = topLeft.Bottom - table.Bounds.Height;
                            }
                        }
                    }
                }
            }
            else if (change.Property == HeaderFontFamilyProperty)
            {
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").FontFamily = this.HeaderFontFamily;
                this.FindControl<VerticalHeader>("VerticalHeaderControl").FontFamily = this.HeaderFontFamily;
            }
            else if (change.Property == MaxTableWidthProperty)
            {
                this.FindControl<Table>("ContentTable").MaxTableWidth = this.MaxTableWidth;
            }
            else if (change.Property == MaxTableHeightProperty)
            {
                this.FindControl<Table>("ContentTable").MaxTableHeight = this.MaxTableHeight;
            }
            else if (change.Property == DefaultRowHeightProperty || change.Property == DefaultColumnWidthProperty)
            {
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
                this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            }
            else if (change.Property == ShowColorPreviewProperty)
            {
                this.FindControl<Table>("ContentTable").InvalidateVisual();
            }
        }

        /// <inheritdoc/>
        protected override void OnTextInput(TextInputEventArgs e)
        {
            base.OnTextInput(e);

            if (this.Selection.Count > 0 && !string.IsNullOrEmpty(e.Text) && e.Text != "\u007f" && !this.FindControl<TextBox>("EditingBox").IsFocused)
            {
                SelectionRange startSelection = this.Selection[this.Selection.Count - 1];
                this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, startSelection.Top));
                this.EditingCell = (startSelection.Left, startSelection.Top);
                this.IsEditing = true;
                this.FindControl<TextBox>("EditingBox").Text = e.Text;
                this.FindControl<TextBox>("EditingBox").SelectionStart = e.Text.Length;
                this.FindControl<TextBox>("EditingBox").SelectionEnd = e.Text.Length;
            }
        }

        /// <summary>
        /// Obtains a text representation of the selected cells, concatenating them with the <see cref="ColumnSeparator"/> and the <see cref="RowSeparator"/>.
        /// </summary>
        /// <param name="selection">The cells whose contents should be included.</param>
        /// <returns>A text representation of the selected cells.</returns>
        public string GetTextRepresentation(IReadOnlyList<SelectionRange> selection)
        {
            Table table = this.FindControl<Table>("ContentTable");

            Dictionary<(int, int), string> data = table.Data;

            int minX = int.MaxValue;
            int minY = int.MaxValue;
            int maxX = int.MinValue;
            int maxY = int.MinValue;

            for (int i = 0; i < selection.Count; i++)
            {
                minX = Math.Min(minX, selection[i].Left);
                minY = Math.Min(minY, selection[i].Top);
                maxX = Math.Max(maxX, selection[i].Right);
                maxY = Math.Max(maxY, selection[i].Bottom);
            }

            SelectionRange dataRange = data.SelectAll(table);

            if (maxX == table.MaxTableWidth)
            {
                maxX = Math.Min(maxX, dataRange.Right);
            }

            if (maxY == table.MaxTableHeight)
            {
                maxY = Math.Min(maxY, dataRange.Bottom);
            }

            bool multipleSelection = selection.Count > 1;

            bool[] anyRow = null;
            bool[] anyColumn = null;

            if (multipleSelection)
            {
                anyRow = new bool[maxY - minY + 1];
                anyColumn = new bool[maxX - minX + 1];
            }

            for (int i = 0; i < selection.Count; i++)
            {

                if (multipleSelection)
                {
                    if (selection[i].IsFinite(table))
                    {
                        for (int x = selection[i].Left; x <= selection[i].Right; x++)
                        {
                            anyColumn[x - minX] = true;
                        }

                        for (int y = selection[i].Top; y <= selection[i].Bottom; y++)
                        {
                            anyRow[y - minY] = true;
                        }
                    }
                    else if (selection[i].IsRows(table) && !selection[i].IsColumns(table))
                    {
                        for (int y = selection[i].Top; y <= selection[i].Bottom; y++)
                        {
                            anyRow[y - minY] = true;
                        }

                        int rightMost = data.GetRight(selection[i].Top, selection[i].Bottom);

                        for (int x = minX; x <= Math.Min(rightMost, maxX); x++)
                        {
                            anyColumn[x - minX] = true;
                        }
                    }
                    else if (selection[i].IsColumns(table) && !selection[i].IsRows(table))
                    {
                        for (int x = selection[i].Left; x <= selection[i].Right; x++)
                        {
                            anyColumn[x - minX] = true;
                        }

                        int bottom = data.GetBottom(selection[i].Left, selection[i].Right);

                        for (int y = minY; y <= Math.Min(bottom, maxY); y++)
                        {
                            anyRow[y - minY] = true;
                        }
                    }
                }
            }

            int lastColumn = maxX - minX;
            int lastRow = maxY - minY;

            if (multipleSelection)
            {
                while (!anyColumn[lastColumn] && lastColumn > 0)
                {
                    lastColumn--;
                }

                while (!anyRow[lastRow] && lastRow > 0)
                {
                    lastRow--;
                }
            }

            string[,] outputMatrix = new string[lastColumn + 1, lastRow + 1];


            foreach (SelectionRange range in selection)
            {
                if (range.IsFinite(table))
                {
                    for (int x = range.Left; x <= range.Right; x++)
                    {
                        for (int y = range.Top; y <= range.Bottom; y++)
                        {
                            data.TryGetValue((x, y), out outputMatrix[x - minX, y - minY]);
                        }
                    }
                }
                else
                {
                    foreach (KeyValuePair<(int, int), string> kvp in data)
                    {
                        int x = kvp.Key.Item1;
                        int y = kvp.Key.Item2;

                        if (x >= range.Left && x <= range.Right && y >= range.Top && y <= range.Bottom)
                        {
                            outputMatrix[x - minX, y - minY] = kvp.Value;
                        }
                    }
                }
            }

            string colSep = new Xeger(Regex.Unescape(this.ColumnSeparator.ToString()), new Random(20230926)).Generate();
            string rowSep = new Xeger(Regex.Unescape(this.RowSeparator.ToString()), new Random(20230926)).Generate();

            StringBuilder tbr = new StringBuilder();

            for (int y = 0; y < lastRow + 1; y++)
            {
                if (!multipleSelection || anyRow[y])
                {
                    for (int x = 0; x < lastColumn + 1; x++)
                    {
                        if (!multipleSelection || anyColumn[x])
                        {
                            string cell = outputMatrix[x, y];

                            if (string.IsNullOrEmpty(cell) || (!this.ColumnSeparator.IsMatch(cell) && !this.RowSeparator.IsMatch(cell)))
                            {
                                tbr.Append(outputMatrix[x, y]);
                            }
                            else
                            {
                                tbr.Append(this.QuoteSymbol);
                                tbr.Append(outputMatrix[x, y]);
                                tbr.Append(this.QuoteSymbol);
                            }

                            if (x < lastColumn)
                            {
                                tbr.Append(colSep);
                            }
                        }
                    }

                    if (y < lastRow)
                    {
                        tbr.Append(rowSep);
                    }
                }
            }

            return tbr.ToString();
        }

        /// <summary>
        /// Returns the data in the currently selected cells.
        /// </summary>
        /// <param name="coordinates">When this method returns, this array will contain the coordinates of the cells corresponding to the values in the return value.</param>
        /// <returns>An array containing the data in the currently selected cells.</returns>
        public string[,] GetSelectedData(out (int, int)[,] coordinates)
        {
            IReadOnlyList<SelectionRange> selection = this.Selection;

            Table table = this.FindControl<Table>("ContentTable");

            Dictionary<(int, int), string> data = table.Data;

            int minX = int.MaxValue;
            int minY = int.MaxValue;
            int maxX = int.MinValue;
            int maxY = int.MinValue;

            for (int i = 0; i < selection.Count; i++)
            {
                minX = Math.Min(minX, selection[i].Left);
                minY = Math.Min(minY, selection[i].Top);
                maxX = Math.Max(maxX, selection[i].Right);
                maxY = Math.Max(maxY, selection[i].Bottom);
            }

            SelectionRange dataRange = data.SelectAll(table);

            if (maxX == table.MaxTableWidth)
            {
                maxX = Math.Min(maxX, dataRange.Right);
            }

            if (maxY == table.MaxTableHeight)
            {
                maxY = Math.Min(maxY, dataRange.Bottom);
            }

            bool multipleSelection = selection.Count > 1;

            bool[] anyRow = null;
            bool[] anyColumn = null;

            if (multipleSelection)
            {
                anyRow = new bool[maxY - minY + 1];
                anyColumn = new bool[maxX - minX + 1];
            }

            for (int i = 0; i < selection.Count; i++)
            {

                if (multipleSelection)
                {
                    if (selection[i].IsFinite(table))
                    {
                        for (int x = selection[i].Left; x <= selection[i].Right; x++)
                        {
                            anyColumn[x - minX] = true;
                        }

                        for (int y = selection[i].Top; y <= selection[i].Bottom; y++)
                        {
                            anyRow[y - minY] = true;
                        }
                    }
                    else if (selection[i].IsRows(table) && !selection[i].IsColumns(table))
                    {
                        for (int y = selection[i].Top; y <= selection[i].Bottom; y++)
                        {
                            anyRow[y - minY] = true;
                        }

                        int rightMost = data.GetRight(selection[i].Top, selection[i].Bottom);

                        for (int x = minX; x <= Math.Min(rightMost, maxX); x++)
                        {
                            anyColumn[x - minX] = true;
                        }
                    }
                    else if (selection[i].IsColumns(table) && !selection[i].IsRows(table))
                    {
                        for (int x = selection[i].Left; x <= selection[i].Right; x++)
                        {
                            anyColumn[x - minX] = true;
                        }

                        int bottom = data.GetBottom(selection[i].Left, selection[i].Right);

                        for (int y = minY; y <= Math.Min(bottom, maxY); y++)
                        {
                            anyRow[y - minY] = true;
                        }
                    }
                }
            }

            int lastColumn = maxX - minX;
            int lastRow = maxY - minY;

            if (multipleSelection)
            {
                while (!anyColumn[lastColumn] && lastColumn > 0)
                {
                    lastColumn--;
                }

                while (!anyRow[lastRow] && lastRow > 0)
                {
                    lastRow--;
                }
            }

            string[,] outputMatrix = new string[lastColumn + 1, lastRow + 1];
            coordinates = new (int, int)[lastColumn + 1, lastRow + 1];

            for (int x = 0; x <= lastColumn; x++)
            {
                for (int y = 0; y <= lastRow; y++)
                {
                    coordinates[x, y] = (-1, -1);
                }
            }

            foreach (SelectionRange range in selection)
            {
                if (range.IsFinite(table))
                {
                    for (int x = range.Left; x <= range.Right; x++)
                    {
                        for (int y = range.Top; y <= range.Bottom; y++)
                        {
                            data.TryGetValue((x, y), out outputMatrix[x - minX, y - minY]);
                            coordinates[x - minX, y - minY] = (x, y);
                        }
                    }
                }
                else
                {
                    foreach (KeyValuePair<(int, int), string> kvp in data)
                    {
                        int x = kvp.Key.Item1;
                        int y = kvp.Key.Item2;

                        if (x >= range.Left && x <= range.Right && y >= range.Top && y <= range.Bottom)
                        {
                            outputMatrix[x - minX, y - minY] = kvp.Value;
                            coordinates[x - minX, y - minY] = (x, y);
                        }
                    }
                }
            }

            return outputMatrix;
        }

        /// <summary>
        /// Copy the selected cells to the clipboard.
        /// </summary>
        public void Copy()
        {
            if (this.Selection.Count > 0)
            {
                string txt = this.GetTextRepresentation(this.Selection);
                _ = Application.Current.Clipboard.SetTextAsync(txt);
                this.FindControl<Overlay>("IconOverlay").Selection = this.Selection.RemoveDuplicates().ToImmutableList();
                this.FindControl<Overlay>("IconOverlay").Icon = Overlay.IconType.Copy;
                this.FindControl<Overlay>("IconOverlay").PauseTransitions();
                this.FindControl<Overlay>("IconOverlay").Opacity = 1;
                this.FindControl<Overlay>("IconOverlay").ResumeTransitions();
                this.FindControl<Overlay>("IconOverlay").Opacity = 0;
            }
        }

        /// <summary>
        /// Cut the selected text out of the spreadsheed and into the clipboard.
        /// </summary>
        public void Cut()
        {
            if (this.Selection.Count > 0)
            {
                Copy();
                this.FindControl<Overlay>("IconOverlay").Icon = Overlay.IconType.Cut;
                ClearContents();
            }
        }

        /// <summary>
        /// Clear the contents of the selected cells.
        /// </summary>
        public void ClearContents()
        {
            if (this.Selection.Count > 0)
            {
                this.FindControl<Table>("ContentTable").Data = this.FindControl<Table>("ContentTable").Data.Remove(this.Selection, this.UndoStack);
                PushNonDataStackNull();
                this.ClearRedoStack();

                this.FindControl<Table>("ContentTable").InvalidateVisual();
            }
        }

        /// <summary>
        /// Paste text or cells from the clipboard onto the spreadsheet. The text is parsed using the <see cref="ColumnSeparator"/> and the <see cref="RowSeparator"/>.
        /// </summary>
        /// <param name="overwriteEmpty">If this is <see langword="true"/>, the contents of cells that are empty in the pasted text is removed; otherwise, empty cells in the text being pasted do not affect the spreadsheet.</param>
        /// <returns>A <see cref="Task"/> that should be awaited.</returns>
        public async Task Paste(bool overwriteEmpty)
        {
            if (this.Selection.Count == 1)
            {
                string text = await Application.Current.Clipboard.GetTextAsync();

                Paste(text, overwriteEmpty);
            }
        }

        /// <summary>
        /// Paste the specified text onto the spreadsheet. The text is parsed using the specified <paramref name="rowSeparator"/> and <paramref name="columnSeparator"/>.
        /// </summary>
        /// <param name="text">The text to parse and paste.</param>
        /// <param name="overwriteEmpty">If this is <see langword="true"/>, the contents of cells that are empty in the pasted text is removed; otherwise, empty cells in the text being pasted do not affect the spreadsheet.</param>
        /// <param name="rowSeparator">The row separator character. If this is <see langword="null"/>, the current <see cref="RowSeparator"/> is used.</param>
        /// <param name="columnSeparator">The column separator character. If this is <see langword="null"/>, the current <see cref="ColumnSeparator"/> is used.</param>
        public void Paste(string text, bool overwriteEmpty, Regex rowSeparator = null, Regex columnSeparator = null)
        {
            if (this.Selection.Count == 1)
            {
                if (!string.IsNullOrEmpty(text))
                {
                    if (rowSeparator == null)
                    {
                        rowSeparator = RowSeparator;
                    }

                    if (columnSeparator == null)
                    {
                        columnSeparator = ColumnSeparator;
                    }

                    string[][] cells = SplitData(text, rowSeparator, columnSeparator, QuoteSymbol, out int width);
                    int height = cells.Length;

                    Table table = this.FindControl<Table>("ContentTable");
                    Dictionary<(int, int), string> data = table.Data;

                    SelectionRange selection = this.Selection[0];

                    if (!selection.IsFinite(table))
                    {
                        if (selection.IsColumns(table) && !selection.IsRows(table))
                        {
                            selection = new SelectionRange(selection.Left, selection.Top, selection.Right, selection.Top + height - 1);
                        }
                        else if (selection.IsRows(table) && !selection.IsColumns(table))
                        {
                            selection = new SelectionRange(selection.Left, selection.Top, selection.Left + width - 1, selection.Bottom);
                        }
                        else
                        {
                            selection = data.SelectAll(table);
                        }
                    }

                    Dictionary<(int, int), string> changedData = new Dictionary<(int, int), string>();
                    Dictionary<(int, int), string> newData = new Dictionary<(int, int), string>();

                    if (selection.Left == selection.Right && selection.Top == selection.Bottom)
                    {
                        for (int y = 0; y < height; y++)
                        {
                            for (int x = 0; x < width; x++)
                            {
                                if (x < cells[y].Length && !string.IsNullOrEmpty(cells[y][x]))
                                {
                                    if (data.TryGetValue((x + selection.Left, y + selection.Top), out string val))
                                    {
                                        changedData[(x + selection.Left, y + selection.Top)] = val;
                                    }
                                    else
                                    {
                                        changedData[(x + selection.Left, y + selection.Top)] = null;
                                    }

                                    data[(x + selection.Left, y + selection.Top)] = cells[y][x];
                                    newData[(x + selection.Left, y + selection.Top)] = cells[y][x];
                                }
                                else if (overwriteEmpty)
                                {
                                    if (data.TryGetValue((x + selection.Left, y + selection.Top), out string val))
                                    {
                                        changedData[(x + selection.Left, y + selection.Top)] = val;
                                        data.Remove((x + selection.Left, y + selection.Top));
                                        newData[(x + selection.Left, y + selection.Top)] = null;
                                    }
                                }
                            }
                        }

                        this.UndoStack.Push(new StackFrame<(int, int), string>(this.Selection, changedData, newData));
                        this.PushNonDataStackNull();
                        this.ClearRedoStack();

                        this.Selection = ImmutableList.Create(new SelectionRange(selection.Left, selection.Top, selection.Left + width - 1, selection.Top + height - 1));
                    }
                    else
                    {
                        for (int y = 0; y < selection.Bottom - selection.Top + 1; y++)
                        {
                            for (int x = 0; x < selection.Right - selection.Left + 1; x++)
                            {
                                int realX = x % width;
                                int realY = y % height;

                                if (realX < cells[realY].Length && !string.IsNullOrEmpty(cells[realY][realX]))
                                {
                                    if (data.TryGetValue((x + selection.Left, y + selection.Top), out string val))
                                    {
                                        changedData[(x + selection.Left, y + selection.Top)] = val;
                                    }
                                    else
                                    {
                                        changedData[(x + selection.Left, y + selection.Top)] = null;
                                    }

                                    data[(x + selection.Left, y + selection.Top)] = cells[realY][realX];
                                    newData[(x + selection.Left, y + selection.Top)] = cells[realY][realX];
                                }
                                else if (overwriteEmpty)
                                {
                                    if (data.TryGetValue((x + selection.Left, y + selection.Top), out string val))
                                    {
                                        changedData[(x + selection.Left, y + selection.Top)] = val;
                                        data.Remove((x + selection.Left, y + selection.Top));
                                        newData[(x + selection.Left, y + selection.Top)] = null;
                                    }
                                }
                            }
                        }

                        this.UndoStack.Push(new StackFrame<(int, int), string>(this.Selection, changedData, newData));
                        this.PushNonDataStackNull();
                        this.ClearRedoStack();

                        this.Selection = ImmutableList.Create(selection);
                    }

                    this.FindControl<Overlay>("IconOverlay").Selection = this.Selection;
                    this.FindControl<Overlay>("IconOverlay").Icon = Overlay.IconType.Paste;
                    this.FindControl<Overlay>("IconOverlay").PauseTransitions();
                    this.FindControl<Overlay>("IconOverlay").Opacity = 1;
                    this.FindControl<Overlay>("IconOverlay").ResumeTransitions();
                    this.FindControl<Overlay>("IconOverlay").Opacity = 0;
                }
            }
        }

        private (double, double)? delayedScrollBarMaximum = null;

        private bool scrollBarXBusy = false;
        private bool scrollBarYBusy = false;

        internal void SetScrollbarMaximum(double maxX, double maxY)
        {
            if (!scrollBarXBusy && !scrollBarYBusy)
            {
                Table table = this.FindControl<Table>("ContentTable");

                (int x, int y) bottomRight = table.Data.BottomRight();

                Rect maxBounds = table.GetCoordinates(bottomRight.x, bottomRight.y);

                maxX = Math.Max(maxX, maxBounds.Right - table.Bounds.Width + 20);
                maxY = Math.Max(maxY, maxBounds.Bottom - table.Bounds.Height);

                this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum = maxX;
                this.FindControl<ScrollBar>("HorizontalScrollBar").ViewportSize = table.Bounds.Width;
                this.FindControl<ScrollBar>("VerticalScrollBar").Maximum = maxY;
                this.FindControl<ScrollBar>("VerticalScrollBar").ViewportSize = table.Bounds.Height;
            }
            else
            {
                delayedScrollBarMaximum = (maxX, maxY);
            }
        }

        /// <summary>
        /// Undoes the last action performed on the spreadsheet.
        /// </summary>
        public void Undo()
        {
            if (this.UndoStack != null && this.UndoStack.Count > 0)
            {
                Table table = this.FindControl<Table>("ContentTable");

                ImmutableList<SelectionRange> selection = null;

                this.RedoStack.Push(this.UndoStack.Pop(table.Data, true, ref selection));
                this.RedoStackCellForeground.Push(this.UndoStackCellForeground.Pop(table.CellForeground, true, ref selection));
                this.RedoStackCellMargin.Push(this.UndoStackCellMargin.Pop(table.CellMargin, true, ref selection));
                this.RedoStackCellHorizontalAlignment.Push(this.UndoStackCellHorizontalAlignment.Pop(table.CellTextAlignment, true, ref selection));
                this.RedoStackCellVerticalAlignment.Push(this.UndoStackCellVerticalAlignment.Pop(table.CellVerticalAlignment, true, ref selection));
                this.RedoStackCellTypeface.Push(this.UndoStackCellTypeface.Pop(table.CellTypefaces, true, ref selection));
                this.RedoStackRowForeground.Push(this.UndoStackRowForeground.Pop(table.RowForeground, true, ref selection));
                this.RedoStackColumnForeground.Push(this.UndoStackColumnForeground.Pop(table.ColumnForeground, true, ref selection));
                this.RedoStackRowTypeface.Push(this.UndoStackRowTypeface.Pop(table.RowTypefaces, true, ref selection));
                this.RedoStackColumnTypeface.Push(this.UndoStackColumnTypeface.Pop(table.ColumnTypefaces, true, ref selection));
                this.RedoStackRowHeight.Push(this.UndoStackRowHeight.Pop(table.RowHeights, true, ref selection));
                this.RedoStackColumnWidth.Push(this.UndoStackColumnWidth.Pop(table.ColumnWidths, true, ref selection));

                if (selection != null)
                {
                    this.Selection = selection;
                }

                table.InvalidateVisual();
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
                this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            }

            this.CanUndo = this.UndoStack.Count > 0;
            this.CanRedo = this.RedoStack.Count > 0;
        }

        /// <summary>
        /// Redoes the last undone action.
        /// </summary>
        public void Redo()
        {
            if (this.RedoStack.Count > 0)
            {
                Table table = this.FindControl<Table>("ContentTable");

                ImmutableList<SelectionRange> selection = null;

                this.UndoStack.Push(this.RedoStack.Pop(table.Data, false, ref selection));
                this.UndoStackCellForeground.Push(this.RedoStackCellForeground.Pop(table.CellForeground, false, ref selection));
                this.UndoStackCellMargin.Push(this.RedoStackCellMargin.Pop(table.CellMargin, false, ref selection));
                this.UndoStackCellHorizontalAlignment.Push(this.RedoStackCellHorizontalAlignment.Pop(table.CellTextAlignment, false, ref selection));
                this.UndoStackCellVerticalAlignment.Push(this.RedoStackCellVerticalAlignment.Pop(table.CellVerticalAlignment, false, ref selection));
                this.UndoStackCellTypeface.Push(this.RedoStackCellTypeface.Pop(table.CellTypefaces, false, ref selection));
                this.UndoStackRowForeground.Push(this.RedoStackRowForeground.Pop(table.RowForeground, false, ref selection));
                this.UndoStackColumnForeground.Push(this.RedoStackColumnForeground.Pop(table.ColumnForeground, false, ref selection));
                this.UndoStackRowTypeface.Push(this.RedoStackRowTypeface.Pop(table.RowTypefaces, false, ref selection));
                this.UndoStackColumnTypeface.Push(this.RedoStackColumnTypeface.Pop(table.ColumnTypefaces, false, ref selection));
                this.UndoStackRowHeight.Push(this.RedoStackRowHeight.Pop(table.RowHeights, false, ref selection));
                this.UndoStackColumnWidth.Push(this.RedoStackColumnWidth.Pop(table.ColumnWidths, false, ref selection));

                if (selection != null)
                {
                    this.Selection = selection;
                }

                table.InvalidateVisual();
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
                this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            }

            this.CanUndo = this.UndoStack.Count > 0;
            this.CanRedo = this.RedoStack.Count > 0;
        }

        internal void PushNonDataStackNull()
        {
            this.UndoStackRowForeground.Push(null);
            this.UndoStackRowHeight.Push(null);
            this.UndoStackRowTypeface.Push(null);

            this.UndoStackColumnForeground.Push(null);
            this.UndoStackColumnWidth.Push(null);
            this.UndoStackColumnTypeface.Push(null);

            this.UndoStackCellForeground.Push(null);
            this.UndoStackCellMargin.Push(null);
            this.UndoStackCellHorizontalAlignment.Push(null);
            this.UndoStackCellTypeface.Push(null);
            this.UndoStackCellVerticalAlignment.Push(null);
        }

        internal void ClearRedoStack()
        {
            this.RedoStack.Clear();

            this.RedoStackRowForeground.Clear();
            this.RedoStackRowHeight.Clear();
            this.RedoStackRowTypeface.Clear();

            this.RedoStackColumnForeground.Clear();
            this.RedoStackColumnWidth.Clear();
            this.RedoStackColumnTypeface.Clear();

            this.RedoStackCellForeground.Clear();
            this.RedoStackCellMargin.Clear();
            this.RedoStackCellHorizontalAlignment.Clear();
            this.RedoStackCellTypeface.Clear();
            this.RedoStackCellVerticalAlignment.Clear();

            this.CanUndo = this.UndoStack.Count > 0;
            this.CanRedo = this.RedoStack.Count > 0;
        }

        internal void ClearUndoStack()
        {
            this.UndoStack.Clear();

            this.UndoStackRowForeground.Clear();
            this.UndoStackRowHeight.Clear();
            this.UndoStackRowTypeface.Clear();

            this.UndoStackColumnForeground.Clear();
            this.UndoStackColumnWidth.Clear();
            this.UndoStackColumnTypeface.Clear();

            this.UndoStackCellForeground.Clear();
            this.UndoStackCellMargin.Clear();
            this.UndoStackCellHorizontalAlignment.Clear();
            this.UndoStackCellTypeface.Clear();
            this.UndoStackCellVerticalAlignment.Clear();

            this.CanUndo = this.UndoStack.Count > 0;
            this.CanRedo = this.RedoStack.Count > 0;
        }

        /// <summary>
        /// Inserts a number of columns equal to the number of selected columns, just before the first selected column.
        /// </summary>
        public void InsertColumns()
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (this.Selection.Count == 1 && this.Selection[0].IsColumns(table) && !this.Selection[0].IsRows(table))
            {
                table.Data = table.Data.InsertColumns(this.Selection[0], this.UndoStack);

                table.CellForeground = table.CellForeground.InsertColumns(this.Selection[0], this.UndoStackCellForeground);
                table.CellMargin = table.CellMargin.InsertColumns(this.Selection[0], this.UndoStackCellMargin);
                table.CellTextAlignment = table.CellTextAlignment.InsertColumns(this.Selection[0], this.UndoStackCellHorizontalAlignment);
                table.CellTypefaces = table.CellTypefaces.InsertColumns(this.Selection[0], this.UndoStackCellTypeface);
                table.CellVerticalAlignment = table.CellVerticalAlignment.InsertColumns(this.Selection[0], this.UndoStackCellVerticalAlignment);

                table.ColumnForeground = table.ColumnForeground.Insert(this.Selection[0].Left, this.Selection[0].Width, this.Selection[0], this.UndoStackColumnForeground);
                table.ColumnTypefaces = table.ColumnTypefaces.Insert(this.Selection[0].Left, this.Selection[0].Width, this.Selection[0], this.UndoStackColumnTypeface);
                table.ColumnWidths = table.ColumnWidths.Insert(this.Selection[0].Left, this.Selection[0].Width, this.Selection[0], this.UndoStackColumnWidth);

                this.UndoStackRowForeground.Push(null);
                this.UndoStackRowHeight.Push(null);
                this.UndoStackRowTypeface.Push(null);

                ClearRedoStack();

                table.InvalidateVisual();
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
                this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            }
        }

        /// <summary>
        /// Deletes the selected columns.
        /// </summary>
        public void DeleteColumns()
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (this.Selection.Count == 1 && this.Selection[0].IsColumns(table) && !this.Selection[0].IsRows(table))
            {
                table.Data = table.Data.DeleteColumns(this.Selection[0], this.UndoStack);

                table.CellForeground = table.CellForeground.DeleteColumns(this.Selection[0], this.UndoStackCellForeground);
                table.CellMargin = table.CellMargin.DeleteColumns(this.Selection[0], this.UndoStackCellMargin);
                table.CellTextAlignment = table.CellTextAlignment.DeleteColumns(this.Selection[0], this.UndoStackCellHorizontalAlignment);
                table.CellTypefaces = table.CellTypefaces.DeleteColumns(this.Selection[0], this.UndoStackCellTypeface);
                table.CellVerticalAlignment = table.CellVerticalAlignment.DeleteColumns(this.Selection[0], this.UndoStackCellVerticalAlignment);

                table.ColumnForeground = table.ColumnForeground.Delete(this.Selection[0].Left, this.Selection[0].Width, this.Selection[0], this.UndoStackColumnForeground);
                table.ColumnTypefaces = table.ColumnTypefaces.Delete(this.Selection[0].Left, this.Selection[0].Width, this.Selection[0], this.UndoStackColumnTypeface);
                table.ColumnWidths = table.ColumnWidths.Delete(this.Selection[0].Left, this.Selection[0].Width, this.Selection[0], this.UndoStackColumnWidth);

                this.UndoStackRowForeground.Push(null);
                this.UndoStackRowHeight.Push(null);
                this.UndoStackRowTypeface.Push(null);

                ClearRedoStack();

                table.InvalidateVisual();
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
                this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            }
        }

        /// <summary>
        /// Inserts a number of rows equal to the number of selected rows, just before the first selected rows.
        /// </summary>
        public void InsertRows()
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (this.Selection.Count == 1 && !this.Selection[0].IsColumns(table) && this.Selection[0].IsRows(table))
            {
                table.Data = table.Data.InsertRows(this.Selection[0], this.UndoStack);

                table.CellForeground = table.CellForeground.InsertRows(this.Selection[0], this.UndoStackCellForeground);
                table.CellMargin = table.CellMargin.InsertRows(this.Selection[0], this.UndoStackCellMargin);
                table.CellTextAlignment = table.CellTextAlignment.InsertRows(this.Selection[0], this.UndoStackCellHorizontalAlignment);
                table.CellTypefaces = table.CellTypefaces.InsertRows(this.Selection[0], this.UndoStackCellTypeface);
                table.CellVerticalAlignment = table.CellVerticalAlignment.InsertRows(this.Selection[0], this.UndoStackCellVerticalAlignment);

                table.RowForeground = table.RowForeground.Insert(this.Selection[0].Top, this.Selection[0].Height, this.Selection[0], this.UndoStackRowForeground);
                table.RowTypefaces = table.RowTypefaces.Insert(this.Selection[0].Top, this.Selection[0].Height, this.Selection[0], this.UndoStackRowTypeface);
                table.RowHeights = table.RowHeights.Insert(this.Selection[0].Top, this.Selection[0].Height, this.Selection[0], this.UndoStackRowHeight);

                this.UndoStackColumnForeground.Push(null);
                this.UndoStackColumnWidth.Push(null);
                this.UndoStackColumnTypeface.Push(null);

                ClearRedoStack();

                table.InvalidateVisual();
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
                this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            }
        }

        /// <summary>
        /// Deletes the selected rows.
        /// </summary>
        public void DeleteRows()
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (this.Selection.Count == 1 && !this.Selection[0].IsColumns(table) && this.Selection[0].IsRows(table))
            {
                table.Data = table.Data.DeleteRows(this.Selection[0], this.UndoStack);

                table.CellForeground = table.CellForeground.DeleteRows(this.Selection[0], this.UndoStackCellForeground);
                table.CellMargin = table.CellMargin.DeleteRows(this.Selection[0], this.UndoStackCellMargin);
                table.CellTextAlignment = table.CellTextAlignment.DeleteRows(this.Selection[0], this.UndoStackCellHorizontalAlignment);
                table.CellTypefaces = table.CellTypefaces.DeleteRows(this.Selection[0], this.UndoStackCellTypeface);
                table.CellVerticalAlignment = table.CellVerticalAlignment.DeleteRows(this.Selection[0], this.UndoStackCellVerticalAlignment);

                table.RowForeground = table.RowForeground.Delete(this.Selection[0].Top, this.Selection[0].Height, this.Selection[0], this.UndoStackRowForeground);
                table.RowTypefaces = table.RowTypefaces.Delete(this.Selection[0].Top, this.Selection[0].Height, this.Selection[0], this.UndoStackRowTypeface);
                table.RowHeights = table.RowHeights.Delete(this.Selection[0].Top, this.Selection[0].Height, this.Selection[0], this.UndoStackRowHeight);

                this.UndoStackColumnForeground.Push(null);
                this.UndoStackColumnWidth.Push(null);
                this.UndoStackColumnTypeface.Push(null);

                ClearRedoStack();

                table.InvalidateVisual();
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
                this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            }
        }

        /// <summary>
        /// Resets the formatting of the selected cells/rows/columns.
        /// </summary>
        public void ResetFormat()
        {
            if (this.Selection.Count > 0)
            {
                Table table = this.FindControl<Table>("ContentTable");

                table.CellTypefaces = table.CellTypefaces.Remove(this.Selection, this.UndoStackCellTypeface);
                table.CellForeground = table.CellForeground.Remove(this.Selection, this.UndoStackCellForeground);
                table.CellMargin = table.CellMargin.Remove(this.Selection, this.UndoStackCellMargin);
                table.CellTextAlignment = table.CellTextAlignment.Remove(this.Selection, this.UndoStackCellHorizontalAlignment);
                table.CellVerticalAlignment = table.CellVerticalAlignment.Remove(this.Selection, this.UndoStackCellVerticalAlignment);
                this.UndoStack.Push(null);

                bool anyRow = false;
                bool anyColumn = false;

                for (int i = 0; i < this.Selection.Count; i++)
                {
                    bool isRows = this.Selection[i].IsRows(table);
                    bool isColumns = this.Selection[i].IsColumns(table);

                    if (isRows)
                    {
                        anyRow = true;
                    }

                    if (isColumns)
                    {
                        anyColumn = true;
                    }

                    if (anyRow && anyColumn)
                    {
                        break;
                    }
                }

                if (anyRow)
                {
                    table.RowForeground = table.RowForeground.Remove(this.Selection.Where(x => x.IsRows(table)).Select(x => (x.Top, x.Bottom)), table.MaxTableWidth, true, this.UndoStackRowForeground);
                    table.RowHeights = table.RowHeights.Remove(this.Selection.Where(x => x.IsRows(table)).Select(x => (x.Top, x.Bottom)), table.MaxTableWidth, true, this.UndoStackRowHeight);
                    table.RowTypefaces = table.RowTypefaces.Remove(this.Selection.Where(x => x.IsRows(table)).Select(x => (x.Top, x.Bottom)), table.MaxTableWidth, true, this.UndoStackRowTypeface);
                }
                else
                {
                    this.UndoStackRowForeground.Push(null);
                    this.UndoStackRowHeight.Push(null);
                    this.UndoStackRowTypeface.Push(null);
                }

                if (anyColumn)
                {
                    table.ColumnForeground = table.ColumnForeground.Remove(this.Selection.Where(x => x.IsColumns(table)).Select(x => (x.Top, x.Bottom)), table.MaxTableWidth, false, this.UndoStackColumnForeground);
                    table.ColumnWidths = table.ColumnWidths.Remove(this.Selection.Where(x => x.IsColumns(table)).Select(x => (x.Top, x.Bottom)), table.MaxTableWidth, false, this.UndoStackColumnWidth);
                    table.ColumnTypefaces = table.ColumnTypefaces.Remove(this.Selection.Where(x => x.IsColumns(table)).Select(x => (x.Top, x.Bottom)), table.MaxTableWidth, false, this.UndoStackColumnTypeface);
                }
                else
                {
                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnWidth.Push(null);
                    this.UndoStackColumnTypeface.Push(null);
                }

                table.InvalidateVisual();
                this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
                this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            }
        }

        /// <summary>
        /// Automatically determines the width of the selected columns.
        /// </summary>
        public void AutoFitWidth()
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (this.Selection.Count > 0 && this.Selection.All(x => x.IsColumns(table) && !x.IsRows(table)))
            {
                HorizontalHeader horizontalHeader = this.FindControl<HorizontalHeader>("HorizontalHeaderControl");

                Dictionary<int, Reference<double>> previousWidths = new Dictionary<int, Reference<double>>();
                Dictionary<int, Reference<double>> newWidths = new Dictionary<int, Reference<double>>();

                foreach (SelectionRange selection in this.Selection)
                {
                    for (int i = selection.Left; i <= selection.Right; i++)
                    {
                        if (this.UndoStack != null && !previousWidths.ContainsKey(i))
                        {
                            if (table.ColumnWidths.TryGetValue(i, out double prevVal))
                            {
                                previousWidths[i] = prevVal;
                            }
                            else
                            {
                                previousWidths[i] = null;
                            }
                        }

                        Reference<double> newWidth = horizontalHeader.AutoFitWidthWithoutStack(i);

                        if (this.UndoStack != null)
                        {
                            newWidths[i] = newWidth;
                        }
                    }
                }

                if (this.UndoStack != null)
                {
                    this.UndoStackColumnWidth.Push(new ValueStackFrame<int, double>(this.Selection, previousWidths, newWidths));

                    this.UndoStack.Push(null);

                    this.UndoStackRowForeground.Push(null);
                    this.UndoStackRowHeight.Push(null);
                    this.UndoStackRowTypeface.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnTypeface.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);
                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);

                    this.ClearRedoStack();
                }

                table.InvalidateVisual();
                horizontalHeader.InvalidateVisual();
            }
        }

        /// <summary>
        /// Automatically determines the height of the selected rows.
        /// </summary>
        public void AutoFitHeight()
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (this.Selection.Count > 0 && this.Selection.All(x => !x.IsColumns(table) && x.IsRows(table)))
            {
                VerticalHeader verticalHeader = this.FindControl<VerticalHeader>("VerticalHeaderControl");

                Dictionary<int, Reference<double>> previousHeights = new Dictionary<int, Reference<double>>();
                Dictionary<int, Reference<double>> newHeights = new Dictionary<int, Reference<double>>();

                foreach (SelectionRange selection in this.Selection)
                {
                    for (int i = selection.Top; i <= selection.Bottom; i++)
                    {
                        if (this.UndoStack != null && !previousHeights.ContainsKey(i))
                        {
                            if (table.RowHeights.TryGetValue(i, out double prevVal))
                            {
                                previousHeights[i] = prevVal;
                            }
                            else
                            {
                                previousHeights[i] = null;
                            }
                        }

                        Reference<double> newHeight = verticalHeader.AutoFitHeightWithoutStack(i);

                        if (this.UndoStack != null)
                        {
                            newHeights[i] = newHeight;
                        }
                    }
                }

                if (this.UndoStack != null)
                {
                    this.UndoStackRowHeight.Push(new ValueStackFrame<int, double>(this.Selection, previousHeights, newHeights));

                    this.UndoStack.Push(null);

                    this.UndoStackRowForeground.Push(null);
                    this.UndoStackRowTypeface.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnTypeface.Push(null);
                    this.UndoStackColumnWidth.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);
                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);

                    this.ClearRedoStack();
                }

                table.InvalidateVisual();
                verticalHeader.InvalidateVisual();
            }
        }

        /// <summary>
        /// Resets the width of the selected columns to the default value.
        /// </summary>
        public void ResetWidth()
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (this.Selection.Count > 0 && this.Selection.Any(x => x.IsRows(table) && x.IsColumns(table)))
            {
                HorizontalHeader horizontalHeader = this.FindControl<HorizontalHeader>("HorizontalHeaderControl");

                Dictionary<int, Reference<double>> previousWidths = new Dictionary<int, Reference<double>>();
                Dictionary<int, Reference<double>> newWidths = new Dictionary<int, Reference<double>>();

                if (this.UndoStack != null)
                {
                    foreach (KeyValuePair<int, double> kvp in table.ColumnWidths)
                    {
                        previousWidths[kvp.Key] = kvp.Value;
                        newWidths[kvp.Key] = kvp.Value;
                    }
                }

                table.ColumnWidths.Clear();

                if (this.UndoStack != null)
                {
                    this.UndoStackColumnWidth.Push(new ValueStackFrame<int, double>(this.Selection, previousWidths, newWidths));

                    this.UndoStack.Push(null);

                    this.UndoStackRowForeground.Push(null);
                    this.UndoStackRowHeight.Push(null);
                    this.UndoStackRowTypeface.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnTypeface.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);
                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);

                    this.ClearRedoStack();
                }

                table.InvalidateVisual();
                horizontalHeader.InvalidateVisual();
            }
            else if (this.Selection.Count > 0 && this.Selection.All(x => x.IsColumns(table) && !x.IsRows(table)))
            {
                HorizontalHeader horizontalHeader = this.FindControl<HorizontalHeader>("HorizontalHeaderControl");

                Dictionary<int, Reference<double>> previousWidths = new Dictionary<int, Reference<double>>();
                Dictionary<int, Reference<double>> newWidths = new Dictionary<int, Reference<double>>();

                foreach (SelectionRange selection in this.Selection)
                {
                    for (int i = selection.Left; i <= selection.Right; i++)
                    {
                        if (this.UndoStack != null && !previousWidths.ContainsKey(i))
                        {
                            if (table.ColumnWidths.TryGetValue(i, out double prevVal))
                            {
                                previousWidths[i] = prevVal;
                            }
                            else
                            {
                                previousWidths[i] = null;
                            }
                        }

                        table.ColumnWidths.Remove(i);

                        if (this.UndoStack != null)
                        {
                            newWidths[i] = null;
                        }
                    }
                }

                if (this.UndoStack != null)
                {
                    this.UndoStackColumnWidth.Push(new ValueStackFrame<int, double>(this.Selection, previousWidths, newWidths));

                    this.UndoStack.Push(null);

                    this.UndoStackRowForeground.Push(null);
                    this.UndoStackRowHeight.Push(null);
                    this.UndoStackRowTypeface.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnTypeface.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);
                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);

                    this.ClearRedoStack();
                }

                table.InvalidateVisual();
                horizontalHeader.InvalidateVisual();
            }
        }

        /// <summary>
        /// Resets the height of the selected columns to the default value.
        /// </summary>
        public void ResetHeight()
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (this.Selection.Count > 0 && this.Selection.Any(x => x.IsColumns(table) && x.IsRows(table)))
            {
                VerticalHeader verticalHeader = this.FindControl<VerticalHeader>("VerticalHeaderControl");

                Dictionary<int, Reference<double>> previousHeights = new Dictionary<int, Reference<double>>();
                Dictionary<int, Reference<double>> newHeights = new Dictionary<int, Reference<double>>();

                if (this.UndoStack != null)
                {
                    foreach (KeyValuePair<int, double> kvp in table.RowHeights)
                    {
                        previousHeights[kvp.Key] = kvp.Value;
                        newHeights[kvp.Key] = kvp.Value;
                    }
                }

                table.RowHeights.Clear();

                if (this.UndoStack != null)
                {
                    this.UndoStackRowHeight.Push(new ValueStackFrame<int, double>(this.Selection, previousHeights, newHeights));

                    this.UndoStack.Push(null);
                    this.UndoStackRowForeground.Push(null);

                    this.UndoStackRowTypeface.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnTypeface.Push(null);
                    this.UndoStackColumnWidth.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);
                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);

                    this.ClearRedoStack();
                }

                table.InvalidateVisual();
                verticalHeader.InvalidateVisual();
            }
            else if (this.Selection.Count > 0 && this.Selection.All(x => !x.IsColumns(table) && x.IsRows(table)))
            {
                VerticalHeader verticalHeader = this.FindControl<VerticalHeader>("VerticalHeaderControl");

                Dictionary<int, Reference<double>> previousHeights = new Dictionary<int, Reference<double>>();
                Dictionary<int, Reference<double>> newHeights = new Dictionary<int, Reference<double>>();

                foreach (SelectionRange selection in this.Selection)
                {
                    for (int i = selection.Top; i <= selection.Bottom; i++)
                    {
                        if (this.UndoStack != null && !previousHeights.ContainsKey(i))
                        {
                            if (table.RowHeights.TryGetValue(i, out double prevVal))
                            {
                                previousHeights[i] = prevVal;
                            }
                            else
                            {
                                previousHeights[i] = null;
                            }
                        }

                        table.RowHeights.Remove(i);

                        if (this.UndoStack != null)
                        {
                            newHeights[i] = null;
                        }
                    }
                }

                if (this.UndoStack != null)
                {
                    this.UndoStackRowHeight.Push(new ValueStackFrame<int, double>(this.Selection, previousHeights, newHeights));

                    this.UndoStack.Push(null);
                    this.UndoStackRowForeground.Push(null);

                    this.UndoStackRowTypeface.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnTypeface.Push(null);
                    this.UndoStackColumnWidth.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);
                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);

                    this.ClearRedoStack();
                }

                table.InvalidateVisual();
                verticalHeader.InvalidateVisual();
            }
        }

        /// <summary>
        /// Set the width of the specified columns.
        /// </summary>
        /// <param name="columnWidths">The column widths to set.</param>
        public void SetWidth(Dictionary<int, double> columnWidths)
        {
            Table table = this.FindControl<Table>("ContentTable");
            HorizontalHeader horizontalHeader = this.FindControl<HorizontalHeader>("HorizontalHeaderControl");

            Dictionary<int, Reference<double>> previousWidths = new Dictionary<int, Reference<double>>();
            Dictionary<int, Reference<double>> newWidths = new Dictionary<int, Reference<double>>();

            foreach (KeyValuePair<int, double> kvp in columnWidths)
            {
                if (this.UndoStack != null && !previousWidths.ContainsKey(kvp.Key))
                {
                    if (table.ColumnWidths.TryGetValue(kvp.Key, out double prevVal))
                    {
                        previousWidths[kvp.Key] = prevVal;
                    }
                    else
                    {
                        previousWidths[kvp.Key] = null;
                    }
                }

                table.ColumnWidths[kvp.Key] = kvp.Value;

                if (this.UndoStack != null)
                {
                    newWidths[kvp.Key] = kvp.Value;
                }
            }

            if (this.UndoStack != null)
            {
                this.UndoStackColumnWidth.Push(new ValueStackFrame<int, double>(this.Selection, previousWidths, newWidths));

                this.UndoStack.Push(null);

                this.UndoStackRowForeground.Push(null);
                this.UndoStackRowHeight.Push(null);
                this.UndoStackRowTypeface.Push(null);

                this.UndoStackColumnForeground.Push(null);
                this.UndoStackColumnTypeface.Push(null);

                this.UndoStackCellForeground.Push(null);
                this.UndoStackCellMargin.Push(null);
                this.UndoStackCellHorizontalAlignment.Push(null);
                this.UndoStackCellTypeface.Push(null);
                this.UndoStackCellVerticalAlignment.Push(null);

                this.ClearRedoStack();
            }

            table.InvalidateVisual();
            horizontalHeader.InvalidateVisual();
        }

        /// <summary>
        /// Set the height of the specified rows.
        /// </summary>
        /// <param name="rowHeights">The row heights to set.</param>
        public void SetHeight(Dictionary<int, double> rowHeights)
        {
            Table table = this.FindControl<Table>("ContentTable");

            VerticalHeader verticalHeader = this.FindControl<VerticalHeader>("VerticalHeaderControl");

            Dictionary<int, Reference<double>> previousHeights = new Dictionary<int, Reference<double>>();
            Dictionary<int, Reference<double>> newHeights = new Dictionary<int, Reference<double>>();

            foreach (KeyValuePair<int, double> kvp in rowHeights)
            {
                if (this.UndoStack != null && !previousHeights.ContainsKey(kvp.Key))
                {
                    if (table.RowHeights.TryGetValue(kvp.Key, out double prevVal))
                    {
                        previousHeights[kvp.Key] = prevVal;
                    }
                    else
                    {
                        previousHeights[kvp.Key] = null;
                    }
                }

                table.RowHeights[kvp.Key] = kvp.Value;

                if (this.UndoStack != null)
                {
                    newHeights[kvp.Key] = kvp.Value;
                }
            }

            if (this.UndoStack != null)
            {
                this.UndoStackRowHeight.Push(new ValueStackFrame<int, double>(this.Selection, previousHeights, newHeights));

                this.UndoStack.Push(null);

                this.UndoStackRowForeground.Push(null);
                this.UndoStackRowTypeface.Push(null);

                this.UndoStackColumnForeground.Push(null);
                this.UndoStackColumnTypeface.Push(null);
                this.UndoStackColumnWidth.Push(null);

                this.UndoStackCellForeground.Push(null);
                this.UndoStackCellMargin.Push(null);
                this.UndoStackCellHorizontalAlignment.Push(null);
                this.UndoStackCellTypeface.Push(null);
                this.UndoStackCellVerticalAlignment.Push(null);

                this.ClearRedoStack();
            }

            table.InvalidateVisual();
            verticalHeader.InvalidateVisual();
        }

        /// <summary>
        /// Sets the typeface of the current selection.
        /// </summary>
        /// <param name="typeface">The <see cref="Typeface"/> to set.</param>
        public void SetTypeface(Typeface typeface)
        {
            if (this.Selection.Count > 0)
            {
                Table table = this.FindControl<Table>("ContentTable");

                bool anyRows = false;
                bool anyColumns = false;
                bool anyAll = false;
                bool anyFinite = false;

                for (int i = 0; i < this.Selection.Count; i++)
                {
                    bool isRows = this.Selection[i].IsRows(table);
                    bool isColumns = this.Selection[i].IsColumns(table);

                    if (isRows && isColumns)
                    {
                        anyAll = true;
                    }
                    else if (isRows && !isColumns)
                    {
                        anyRows = true;
                    }
                    else if (isColumns && !isRows)
                    {
                        anyColumns = true;
                    }
                    else
                    {
                        anyFinite = true;
                    }
                }

                if (anyAll)
                {
                    this.FontFamily = typeface.FontFamily;
                    this.FontStyle = typeface.Style;
                }
                else
                {
                    if (anyFinite)
                    {
                        table.CellTypefaces.SetValue(Selection.Where(x => x.IsFinite(table)), typeface, this.UndoStackCellTypeface);
                    }
                    else
                    {
                        this.UndoStackCellTypeface.Push(null);
                    }

                    if (anyRows)
                    {
                        table.RowTypefaces.SetValue(Selection.Where(x => x.IsRows(table)).Select(x => (x.Top, x.Bottom)), typeface, table.MaxTableWidth, true, this.UndoStackRowTypeface);
                    }
                    else
                    {
                        this.UndoStackRowTypeface.Push(null);
                    }

                    if (anyColumns)
                    {
                        table.ColumnTypefaces.SetValue(Selection.Where(x => x.IsColumns(table)).Select(x => (x.Left, x.Right)), typeface, table.MaxTableHeight, false, this.UndoStackColumnTypeface);
                    }
                    else
                    {
                        this.UndoStackColumnTypeface.Push(null);
                    }

                    this.UndoStack.Push(null);

                    this.UndoStackRowForeground.Push(null);
                    this.UndoStackRowHeight.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnWidth.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);
                    this.RedoStack.Clear();
                    table.InvalidateVisual();
                }
            }
        }

        /// <summary>
        /// Sets the foreground colour of the current selection.
        /// </summary>
        /// <param name="foreground">The foreground colour to set.</param>
        public void SetForeground(IBrush foreground)
        {
            if (this.Selection.Count > 0)
            {
                Table table = this.FindControl<Table>("ContentTable");

                bool anyRows = false;
                bool anyColumns = false;
                bool anyAll = false;
                bool anyFinite = false;

                for (int i = 0; i < this.Selection.Count; i++)
                {
                    bool isRows = this.Selection[i].IsRows(table);
                    bool isColumns = this.Selection[i].IsColumns(table);

                    if (isRows && isColumns)
                    {
                        anyAll = true;
                    }
                    else if (isRows && !isColumns)
                    {
                        anyRows = true;
                    }
                    else if (isColumns && !isRows)
                    {
                        anyColumns = true;
                    }
                    else
                    {
                        anyFinite = true;
                    }
                }

                if (anyAll)
                {
                    this.Foreground = foreground;
                    table.InvalidateVisual();
                }
                else
                {
                    if (anyFinite)
                    {
                        table.CellForeground.SetValue(Selection.Where(x => x.IsFinite(table)), foreground, this.UndoStackCellForeground);
                    }
                    else
                    {
                        this.UndoStackCellForeground.Push(null);
                    }

                    if (anyRows)
                    {
                        table.RowForeground.SetValue(Selection.Where(x => x.IsRows(table)).Select(x => (x.Top, x.Bottom)), foreground, table.MaxTableWidth, true, this.UndoStackRowForeground);
                    }
                    else
                    {
                        this.UndoStackRowForeground.Push(null);
                    }

                    if (anyColumns)
                    {
                        table.ColumnForeground.SetValue(Selection.Where(x => x.IsColumns(table)).Select(x => (x.Left, x.Right)), foreground, table.MaxTableHeight, false, this.UndoStackColumnForeground);
                    }
                    else
                    {
                        this.UndoStackColumnForeground.Push(null);
                    }

                    this.UndoStack.Push(null);

                    this.UndoStackRowTypeface.Push(null);
                    this.UndoStackRowHeight.Push(null);

                    this.UndoStackColumnTypeface.Push(null);
                    this.UndoStackColumnWidth.Push(null);

                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);
                    this.RedoStack.Clear();
                    table.InvalidateVisual();
                }
            }
        }

        /// <summary>
        /// Returns the typeface for the specified cell.
        /// </summary>
        /// <param name="left">The horizontal coordinate of the cell.</param>
        /// <param name="top">The vertical coordinate of the cell.</param>
        /// <returns>The typeface associated to the specified cell.</returns>
        public Typeface GetTypeface(int left, int top)
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (!(table.CellTypefaces.TryGetValue((left, top), out Typeface tbr) || table.RowTypefaces.TryGetValue(top, out tbr) || table.ColumnTypefaces.TryGetValue(left, out tbr)))
            {
                tbr = new Typeface(table.FontFamily, table.FontStyle, table.FontWeight);
            }

            return tbr;
        }

        /// <summary>
        /// Returns the column width and row height for the specified cell.
        /// </summary>
        /// <param name="left">The horizontal coordinate of the cell.</param>
        /// <param name="top">The vertical coordinate of the cell.</param>
        /// <returns>The column width and row height of the specified cell.</returns>
        public (double width, double height) GetCellSize(int left, int top)
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (!table.ColumnWidths.TryGetValue(left, out double width))
            {
                width = table.DefaultColumnWidth;
            }

            if (!table.RowHeights.TryGetValue(top, out double height))
            {
                height = table.DefaultRowHeight;
            }

            return (width, height);
        }

        /// <summary>
        /// Returns the text alignment for the specified cell.
        /// </summary>
        /// <param name="left">The horizontal coordinate of the cell.</param>
        /// <param name="top">The vertical coordinate of the cell.</param>
        /// <returns>The the text alignment associated to the specified cell.</returns>
        public (TextAlignment, VerticalAlignment) GetAlignment(int left, int top)
        {
            Table table = this.FindControl<Table>("ContentTable");

            if (!table.CellTextAlignment.TryGetValue((left, top), out TextAlignment tbrHor))
            {
                tbrHor = table.DefaultTextAlignment;
            }

            if (!table.CellVerticalAlignment.TryGetValue((left, top), out VerticalAlignment tbrVer))
            {
                tbrVer = table.DefaultVerticalAlignment;
            }

            return (tbrHor, tbrVer);
        }

        /// <summary>
        /// Sets the horizontal text alignment of the current selection.
        /// </summary>
        /// <param name="textAlignment">The text alignment to set.</param>
        public void SetTextAlignment(TextAlignment textAlignment)
        {
            if (this.Selection.Count > 0)
            {
                Table table = this.FindControl<Table>("ContentTable");

                bool anyAll = false;

                for (int i = 0; i < this.Selection.Count; i++)
                {
                    bool isRows = this.Selection[i].IsRows(table);
                    bool isColumns = this.Selection[i].IsColumns(table);

                    if (isRows && isColumns)
                    {
                        anyAll = true;
                    }
                }

                if (anyAll)
                {
                    this.DefaultTextAlignment = textAlignment;
                    table.InvalidateVisual();
                }
                else
                {
                    SelectionRange dataRange = table.Data.SelectAll(table);
                    table.CellTextAlignment.SetValue(Selection.Select(x => x.IsFinite(table) ? x : x.Intersection(dataRange)), textAlignment, this.UndoStackCellHorizontalAlignment);

                    this.UndoStack.Push(null);

                    this.UndoStackRowForeground.Push(null);
                    this.UndoStackRowHeight.Push(null);
                    this.UndoStackRowTypeface.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnWidth.Push(null);
                    this.UndoStackColumnTypeface.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellVerticalAlignment.Push(null);

                    this.RedoStack.Clear();
                    table.InvalidateVisual();
                }
            }
        }

        /// <summary>
        /// Set the value of the specified cells.
        /// </summary>
        /// <param name="data">The cell coordinates and values to set. Elements with value <see langword="null"/> will be removed from the spreadsheet.</param>
        public void SetData(IEnumerable<KeyValuePair<(int, int), string>> data)
        {
            Dictionary<(int, int), string> newData = new Dictionary<(int, int), string>();
            Dictionary<(int, int), string> previousData = new Dictionary<(int, int), string>();

            HashSet<(int, int)> selection = new HashSet<(int, int)>();

            foreach (KeyValuePair<(int, int), string> kvp in data)
            {
                if (this.UndoStack != null && this.Data.TryGetValue(kvp.Key, out string prevVal))
                {
                    previousData[kvp.Key] = prevVal;
                    newData[kvp.Key] = kvp.Value;
                    selection.Add(kvp.Key);
                }

                if (!string.IsNullOrEmpty(kvp.Value))
                {
                    this.Data[kvp.Key] = kvp.Value;
                }
                else
                {
                    this.Data.Remove(kvp.Key);
                }
            }

            if (this.UndoStack != null)
            {
                this.UndoStack.Push(new StackFrame<(int, int), string>(selection.Select(x => new SelectionRange(x)).Consolidate().ToImmutableList(), previousData, newData));
                this.PushNonDataStackNull();
                this.ClearRedoStack();
            }

            this.FindControl<Table>("ContentTable").InvalidateVisual();
        }


        /// <summary>
        /// Sets the vertical text alignment of the current selection.
        /// </summary>
        /// <param name="verticalAlignment">The text alignment to set.</param>
        public void SetVerticalAlignment(VerticalAlignment verticalAlignment)
        {
            if (this.Selection.Count > 0)
            {
                Table table = this.FindControl<Table>("ContentTable");

                bool anyAll = false;

                for (int i = 0; i < this.Selection.Count; i++)
                {
                    bool isRows = this.Selection[i].IsRows(table);
                    bool isColumns = this.Selection[i].IsColumns(table);

                    if (isRows && isColumns)
                    {
                        anyAll = true;
                    }
                }

                if (anyAll)
                {
                    this.DefaultVerticalAlignment = verticalAlignment;
                    table.InvalidateVisual();
                }
                else
                {
                    SelectionRange dataRange = table.Data.SelectAll(table);
                    table.CellVerticalAlignment.SetValue(Selection.Select(x => x.IsFinite(table) ? x : x.Intersection(dataRange)), verticalAlignment, this.UndoStackCellVerticalAlignment);

                    this.UndoStack.Push(null);

                    this.UndoStackRowForeground.Push(null);
                    this.UndoStackRowHeight.Push(null);
                    this.UndoStackRowTypeface.Push(null);

                    this.UndoStackColumnForeground.Push(null);
                    this.UndoStackColumnWidth.Push(null);
                    this.UndoStackColumnTypeface.Push(null);

                    this.UndoStackCellForeground.Push(null);
                    this.UndoStackCellMargin.Push(null);
                    this.UndoStackCellTypeface.Push(null);
                    this.UndoStackCellHorizontalAlignment.Push(null);

                    this.RedoStack.Clear();
                    table.InvalidateVisual();
                }
            }
        }

        /// <inheritdoc/>
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);

            if (!e.Handled && !this.FindControl<TextBox>("EditingBox").IsFocused)
            {
                if (this.Selection.Count > 0)
                {
                    SelectionRange startSelection = this.Selection[this.Selection.Count - 1];

                    if (e.Key == Key.Down)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, startSelection.Top + 1));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsColumns(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, startSelection.Right, startSelection.Bottom + 1));
                            }
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsColumns(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, Math.Min(startSelection.Bottom, startSelection.Top + 1), startSelection.Right, startSelection.Bottom));
                            }
                        }
                        else if (e.KeyModifiers.HasFlag(Spreadsheet.ControlModifier))
                        {
                            int target = this.FindControl<Table>("ContentTable").Data.GetBoundaryDown(startSelection.Left, startSelection.Right, startSelection.Bottom);

                            if (target >= 0)
                            {
                                if (e.KeyModifiers == (KeyModifiers.Shift | Spreadsheet.ControlModifier))
                                {
                                    this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, startSelection.Right, target));
                                }
                                else
                                {
                                    this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, target));
                                }
                            }
                        }

                        e.Handled = true;
                    }
                    else if (e.Key == Key.Up)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, Math.Max(0, startSelection.Top - 1)));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsColumns(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, Math.Max(0, startSelection.Top - 1), startSelection.Right, startSelection.Bottom));
                            }
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsColumns(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, startSelection.Right, Math.Max(startSelection.Top, startSelection.Bottom - 1)));
                            }
                        }
                        else if (e.KeyModifiers.HasFlag(Spreadsheet.ControlModifier))
                        {
                            int target = this.FindControl<Table>("ContentTable").Data.GetBoundaryUp(startSelection.Left, startSelection.Right, startSelection.Top);

                            if (target >= 0)
                            {
                                if (e.KeyModifiers == (KeyModifiers.Shift | Spreadsheet.ControlModifier))
                                {
                                    this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, target, startSelection.Right, startSelection.Bottom));
                                }
                                else
                                {
                                    this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, target));
                                }
                            }
                        }

                        e.Handled = true;
                    }
                    if (e.Key == Key.Right)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left + 1, startSelection.Top));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsRows(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, startSelection.Right + 1, startSelection.Bottom));
                            }
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsRows(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(Math.Min(startSelection.Right, startSelection.Left + 1), startSelection.Top, startSelection.Right, startSelection.Bottom));
                            }
                        }
                        else if (e.KeyModifiers.HasFlag(Spreadsheet.ControlModifier))
                        {
                            int target = this.FindControl<Table>("ContentTable").Data.GetBoundaryRight(startSelection.Right, startSelection.Top, startSelection.Bottom);

                            if (target >= 0)
                            {
                                if (e.KeyModifiers == (KeyModifiers.Shift | Spreadsheet.ControlModifier))
                                {
                                    this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, target, startSelection.Bottom));
                                }
                                else
                                {
                                    this.Selection = ImmutableList.Create(new SelectionRange(target, startSelection.Top));
                                }
                            }
                        }

                        e.Handled = true;
                    }
                    else if (e.Key == Key.Left)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(Math.Max(0, startSelection.Left - 1), startSelection.Top));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsRows(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(Math.Max(0, startSelection.Left - 1), startSelection.Top, startSelection.Right, startSelection.Bottom));
                            }
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsRows(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, Math.Max(startSelection.Left, startSelection.Right - 1), startSelection.Bottom));
                            }
                        }
                        else if (e.KeyModifiers.HasFlag(Spreadsheet.ControlModifier))
                        {
                            int target = this.FindControl<Table>("ContentTable").Data.GetBoundaryLeft(startSelection.Left, startSelection.Top, startSelection.Bottom);

                            if (target >= 0)
                            {
                                if (e.KeyModifiers == (KeyModifiers.Shift | Spreadsheet.ControlModifier))
                                {
                                    this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(target, startSelection.Top, startSelection.Right, startSelection.Bottom));
                                }
                                else
                                {
                                    this.Selection = ImmutableList.Create(new SelectionRange(target, startSelection.Top));
                                }
                            }
                        }

                        e.Handled = true;
                    }
                    else if (e.Key == Key.Return)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, startSelection.Top + 1));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, Math.Max(0, startSelection.Top - 1)));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            ClearContents();
                        }

                        e.Handled = true;
                    }
                    else if (e.Key == Key.F2)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, startSelection.Top));
                            this.EditingCell = (startSelection.Left, startSelection.Top);
                            this.IsEditing = true;

                            e.Handled = true;
                        }
                    }
                    else if (e.Key == Key.Delete)
                    {
                        if (e.KeyModifiers == KeyModifiers.None || e.KeyModifiers == Spreadsheet.ControlModifier)
                        {
                            ClearContents();
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            Cut();
                        }

                        e.Handled = true;
                    }
                    else if (e.Key == Key.Back)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, startSelection.Top));
                            this.EditingCell = (startSelection.Left, startSelection.Top);
                            this.IsEditing = true;
                            this.FindControl<TextBox>("EditingBox").Text = "";

                            e.Handled = true;
                        }
                    }
                    else if (e.Key == Key.Home)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(0, startSelection.Top));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(0, startSelection.Top, startSelection.Right, startSelection.Bottom));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, 0));
                        }
                        else if (e.KeyModifiers == (KeyModifiers.Alt | KeyModifiers.Shift))
                        {
                            this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, 0, startSelection.Right, startSelection.Bottom));
                        }
                        else if (e.KeyModifiers == Spreadsheet.ControlModifier)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(0, 0));
                        }
                        else if (e.KeyModifiers == (Spreadsheet.ControlModifier | KeyModifiers.Shift))
                        {
                            this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(0, 0, startSelection.Right, startSelection.Bottom));
                        }

                        e.Handled = true;
                    }
                    else if (e.Key == Key.End)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            int right = this.FindControl<Table>("ContentTable").Data.GetRight(startSelection.Top, startSelection.Top);
                            this.Selection = ImmutableList.Create(new SelectionRange(right, startSelection.Top));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            int right = this.FindControl<Table>("ContentTable").Data.GetRight(startSelection.Top, startSelection.Bottom);
                            this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, right, startSelection.Bottom));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            int bottom = this.FindControl<Table>("ContentTable").Data.GetBottom(startSelection.Left, startSelection.Left);
                            this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, bottom));
                        }
                        else if (e.KeyModifiers == (KeyModifiers.Alt | KeyModifiers.Shift))
                        {
                            int bottom = this.FindControl<Table>("ContentTable").Data.GetBottom(startSelection.Left, startSelection.Right);
                            this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, startSelection.Right, bottom));
                        }
                        else if (e.KeyModifiers == Spreadsheet.ControlModifier)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(this.FindControl<Table>("ContentTable").Data.BottomRight()));
                        }
                        else if (e.KeyModifiers == (Spreadsheet.ControlModifier | KeyModifiers.Shift))
                        {
                            (int, int) bottomRight = this.FindControl<Table>("ContentTable").Data.BottomRight();

                            this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, bottomRight.Item1, bottomRight.Item2));
                        }

                        e.Handled = true;
                    }
                    else if (e.Key == Key.A && e.KeyModifiers == Spreadsheet.ControlModifier)
                    {
                        this.Selection = ImmutableList.Create(this.FindControl<Table>("ContentTable").Data.SelectAll(this.FindControl<Table>("ContentTable")));
                        e.Handled = true;
                    }
                    else if ((e.Key == Key.C && e.KeyModifiers == Spreadsheet.ControlModifier) || (e.Key == Key.Insert && e.KeyModifiers == Spreadsheet.ControlModifier))
                    {
                        Copy();
                        e.Handled = true;
                    }
                    else if ((e.Key == Key.X && e.KeyModifiers == Spreadsheet.ControlModifier))
                    {
                        Cut();
                        e.Handled = true;
                    }
                    else if ((e.Key == Key.V && e.KeyModifiers == Spreadsheet.ControlModifier) || (e.Key == Key.Insert && e.KeyModifiers == KeyModifiers.Shift))
                    {
                        _ = Paste(true);
                        e.Handled = true;
                    }
                    else if ((e.Key == Key.V && e.KeyModifiers == (Spreadsheet.ControlModifier | KeyModifiers.Shift)) || (e.Key == Key.Insert && e.KeyModifiers == (KeyModifiers.Shift | KeyModifiers.Alt)))
                    {
                        _ = Paste(false);
                        e.Handled = true;
                    }
                    else if (e.Key == Key.Z && e.KeyModifiers == Spreadsheet.ControlModifier)
                    {
                        Undo();
                        e.Handled = true;
                    }
                    else if (e.Key == Key.Y && e.KeyModifiers == Spreadsheet.ControlModifier)
                    {
                        Redo();
                        e.Handled = true;
                    }
                    else if (e.Key == Key.Tab)
                    {
                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left + 1, startSelection.Top));
                            e.Handled = true;
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(Math.Max(0, startSelection.Left - 1), startSelection.Top));
                            e.Handled = true;
                        }
                    }
                    else if (e.Key == Key.PageDown)
                    {
                        int pageSize = Math.Max(1, this.FindControl<Table>("ContentTable").lastDrawnHeight - 2);

                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, startSelection.Top + pageSize));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsColumns(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, startSelection.Right, startSelection.Bottom + pageSize));
                            }
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsColumns(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, Math.Min(startSelection.Bottom, startSelection.Top + pageSize), startSelection.Right, startSelection.Bottom));
                            }
                        }
                        else if (e.KeyModifiers.HasFlag(Spreadsheet.ControlModifier))
                        {
                            int target = this.FindControl<Table>("ContentTable").Data.GetBoundaryDown(startSelection.Left, startSelection.Right, startSelection.Bottom);

                            if (target >= 0)
                            {
                                if (e.KeyModifiers == (KeyModifiers.Shift | Spreadsheet.ControlModifier))
                                {
                                    this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, startSelection.Right, target));
                                }
                                else
                                {
                                    this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, target));
                                }
                            }
                        }

                        e.Handled = true;
                    }
                    else if (e.Key == Key.PageUp)
                    {
                        int pageSize = Math.Max(1, this.FindControl<Table>("ContentTable").lastDrawnHeight - 2);

                        if (e.KeyModifiers == KeyModifiers.None)
                        {
                            this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, Math.Max(0, startSelection.Top - pageSize)));
                        }
                        else if (e.KeyModifiers == KeyModifiers.Shift)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsColumns(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, Math.Max(0, startSelection.Top - pageSize), startSelection.Right, startSelection.Bottom));
                            }
                        }
                        else if (e.KeyModifiers == KeyModifiers.Alt)
                        {
                            if (!this.Selection[this.Selection.Count - 1].IsColumns(this.FindControl<Table>("ContentTable")))
                            {
                                this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, startSelection.Top, startSelection.Right, Math.Max(startSelection.Top, startSelection.Bottom - pageSize)));
                            }
                        }
                        else if (e.KeyModifiers.HasFlag(Spreadsheet.ControlModifier))
                        {
                            int target = this.FindControl<Table>("ContentTable").Data.GetBoundaryUp(startSelection.Left, startSelection.Right, startSelection.Top);

                            if (target >= 0)
                            {
                                if (e.KeyModifiers == (KeyModifiers.Shift | Spreadsheet.ControlModifier))
                                {
                                    this.Selection = this.Selection.SetItem(this.Selection.Count - 1, new SelectionRange(startSelection.Left, target, startSelection.Right, startSelection.Bottom));
                                }
                                else
                                {
                                    this.Selection = ImmutableList.Create(new SelectionRange(startSelection.Left, target));
                                }
                            }
                        }

                        e.Handled = true;
                    }
                }
                else if (e.Key == Key.A && e.KeyModifiers == Spreadsheet.ControlModifier)
                {
                    this.Selection = ImmutableList.Create(this.FindControl<Table>("ContentTable").Data.SelectAll(this.FindControl<Table>("ContentTable")));
                    e.Handled = true;
                }
                else if (e.Key == Key.Z && e.KeyModifiers == Spreadsheet.ControlModifier)
                {
                    Undo();
                    e.Handled = true;
                }
                else if (e.Key == Key.Y && e.KeyModifiers == Spreadsheet.ControlModifier)
                {
                    Redo();
                    e.Handled = true;
                }
            }
        }

        /// <inheritdoc/>
        protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
        {
            base.OnPointerWheelChanged(e);

            this.FindControl<ScrollBar>("VerticalScrollBar").Maximum = Math.Max(this.FindControl<ScrollBar>("VerticalScrollBar").Value - this.FindControl<ScrollBar>("VerticalScrollBar").LargeChange * e.Delta.Y, this.FindControl<ScrollBar>("VerticalScrollBar").Maximum);
            this.FindControl<ScrollBar>("VerticalScrollBar").Value = this.FindControl<ScrollBar>("VerticalScrollBar").Value - this.FindControl<ScrollBar>("VerticalScrollBar").LargeChange * e.Delta.Y;
            this.FindControl<ScrollBar>("VerticalScrollBar").ViewportSize = this.FindControl<Table>("ContentTable").Bounds.Height;

            this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum = Math.Max(this.FindControl<ScrollBar>("HorizontalScrollBar").Value - this.FindControl<ScrollBar>("HorizontalScrollBar").LargeChange * e.Delta.X, this.FindControl<ScrollBar>("HorizontalScrollBar").Maximum);
            this.FindControl<ScrollBar>("HorizontalScrollBar").Value = this.FindControl<ScrollBar>("HorizontalScrollBar").Value - this.FindControl<ScrollBar>("HorizontalScrollBar").LargeChange * e.Delta.X;
            this.FindControl<ScrollBar>("HorizontalScrollBar").ViewportSize = this.FindControl<Table>("ContentTable").Bounds.Width;
        }

        /// <summary>
        /// Obtain a string representation of all the data currently present in the spreadsheet.
        /// </summary>
        /// <returns>A string representation of all the data currently present in the spreadsheet.</returns>
        public string SerializeData()
        {
            Table table = this.FindControl<Table>("ContentTable");

            (int x, int y) bottomRight = table.Data.BottomRight();

            return this.GetTextRepresentation(new SelectionRange[] { new SelectionRange(0, 0, bottomRight.x, bottomRight.y) });
        }

        /// <summary>
        /// Obtain a string representation of all the formatting information currently present in the spreadsheet.
        /// </summary>
        /// <returns>A string representation of all the formatting information currently present in the spreadsheet.</returns>
        public string SerializeFormat()
        {
            Table table = this.FindControl<Table>("ContentTable");

            string colSep = new Xeger(Regex.Unescape(this.ColumnSeparator.ToString()), new Random(20230926)).Generate();
            string rowSep = new Xeger(Regex.Unescape(this.RowSeparator.ToString()), new Random(20230926)).Generate();

            StringBuilder sb = new StringBuilder();

            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Colour"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Family"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Size"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Style"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Weight"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Column Width"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Row Height"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Margin"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Horizontal Alignment"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Vertical Alignment"));
            sb.Append(rowSep);

            sb.Append("DEFAULT");
            sb.Append(colSep);
            sb.Append(StringifyAndQuote(table.Foreground));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote(table.FontFamily.Name));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote(table.FontSize.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote(table.FontStyle));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote(table.FontWeight));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote(table.DefaultColumnWidth.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote(table.DefaultRowHeight.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote(table.DefaultMargin));
            sb.Append(colSep);

            switch (table.DefaultTextAlignment)
            {
                case TextAlignment.Left:
                    sb.Append(StringifyAndQuote(">"));
                    break;
                case TextAlignment.Center:
                    sb.Append(StringifyAndQuote("|"));
                    break;
                case TextAlignment.Right:
                    sb.Append(StringifyAndQuote("<"));
                    break;
            }

            sb.Append(colSep);

            switch (table.DefaultVerticalAlignment)
            {
                case VerticalAlignment.Top:
                    sb.Append(StringifyAndQuote("^"));
                    break;
                case VerticalAlignment.Center:
                case VerticalAlignment.Stretch:
                    sb.Append(StringifyAndQuote("-"));
                    break;
                case VerticalAlignment.Bottom:
                    sb.Append(StringifyAndQuote("v"));
                    break;
            }
            sb.Append(rowSep);

            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Colour"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Row Height"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Family"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Style"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Weight"));
            sb.Append(rowSep);

            HashSet<int> allRows = new HashSet<int>(table.RowForeground.Keys.Concat(table.RowHeights.Keys).Concat(table.RowTypefaces.Keys));

            foreach (int i in allRows.OrderBy(i => i))
            {
                sb.Append(StringifyAndQuote("Row " + (i + 1).ToString()));

                sb.Append(colSep);

                if (table.RowForeground.TryGetValue(i, out IBrush rowForeground))
                {
                    sb.Append(StringifyAndQuote(rowForeground));
                }

                sb.Append(colSep);

                if (table.RowHeights.TryGetValue(i, out double rowHeight))
                {
                    sb.Append(StringifyAndQuote(rowHeight.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }

                sb.Append(colSep);

                if (table.RowTypefaces.TryGetValue(i, out Typeface rowFace))
                {
                    sb.Append(StringifyAndQuote(rowFace.FontFamily.Name));
                    sb.Append(colSep);
                    sb.Append(StringifyAndQuote(rowFace.Style));
                    sb.Append(colSep);
                    sb.Append(StringifyAndQuote(rowFace.Weight));
                }
                else
                {
                    sb.Append(colSep);
                    sb.Append(colSep);
                }

                sb.Append(rowSep);
            }

            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Colour"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Column Width"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Family"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Style"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Weight"));
            sb.Append(rowSep);

            HashSet<int> allColumns = new HashSet<int>(table.ColumnForeground.Keys.Concat(table.ColumnWidths.Keys).Concat(table.ColumnTypefaces.Keys));

            foreach (int i in allColumns.OrderBy(i => i))
            {
                sb.Append(StringifyAndQuote("Column " + HorizontalHeader.GetLetter(i)));

                sb.Append(colSep);

                if (table.ColumnForeground.TryGetValue(i, out IBrush columnForeground))
                {
                    sb.Append(StringifyAndQuote(columnForeground));
                }

                sb.Append(colSep);

                if (table.ColumnWidths.TryGetValue(i, out double columnWidth))
                {
                    sb.Append(StringifyAndQuote(columnWidth.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }

                sb.Append(colSep);

                if (table.ColumnTypefaces.TryGetValue(i, out Typeface columnFace))
                {
                    sb.Append(StringifyAndQuote(columnFace.FontFamily.Name));
                    sb.Append(colSep);
                    sb.Append(StringifyAndQuote(columnFace.Style));
                    sb.Append(colSep);
                    sb.Append(StringifyAndQuote(columnFace.Weight));
                }
                else
                {
                    sb.Append(colSep);
                    sb.Append(colSep);
                }

                sb.Append(rowSep);
            }

            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Colour"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Margin"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Family"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Style"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Font Weight"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Horizontal Alignment"));
            sb.Append(colSep);
            sb.Append(StringifyAndQuote("Vertical Alignment"));
            sb.Append(rowSep);

            HashSet<(int, int)> allCells = new HashSet<(int, int)>(table.CellForeground.Keys.Concat(table.CellMargin.Keys).Concat(table.CellTypefaces.Keys).Concat(table.CellTextAlignment.Keys).Concat(table.CellVerticalAlignment.Keys));

            foreach ((int x, int y) pos in allCells.OrderBy(i => i))
            {
                sb.Append(StringifyAndQuote("Cell " + HorizontalHeader.GetLetter(pos.x) + (pos.y + 1).ToString()));

                sb.Append(colSep);

                if (table.CellForeground.TryGetValue(pos, out IBrush cellForeground))
                {
                    sb.Append(StringifyAndQuote(cellForeground));
                }

                sb.Append(colSep);

                if (table.CellMargin.TryGetValue(pos, out Thickness cellMargin))
                {
                    sb.Append(StringifyAndQuote(cellMargin.ToString()));
                }

                sb.Append(colSep);

                if (table.CellTypefaces.TryGetValue(pos, out Typeface cellFace))
                {
                    sb.Append(StringifyAndQuote(cellFace.FontFamily.Name));
                    sb.Append(colSep);
                    sb.Append(StringifyAndQuote(cellFace.Style));
                    sb.Append(colSep);
                    sb.Append(StringifyAndQuote(cellFace.Weight));
                }
                else
                {
                    sb.Append(colSep);
                    sb.Append(colSep);
                }

                sb.Append(colSep);

                if (table.CellTextAlignment.TryGetValue(pos, out TextAlignment cellAlignment))
                {
                    switch (cellAlignment)
                    {
                        case TextAlignment.Left:
                            sb.Append(StringifyAndQuote(">"));
                            break;
                        case TextAlignment.Center:
                            sb.Append(StringifyAndQuote("|"));
                            break;
                        case TextAlignment.Right:
                            sb.Append(StringifyAndQuote("<"));
                            break;
                    }
                }

                sb.Append(colSep);

                if (table.CellVerticalAlignment.TryGetValue(pos, out VerticalAlignment cellVertAlignment))
                {
                    switch (cellVertAlignment)
                    {
                        case VerticalAlignment.Top:
                            sb.Append(StringifyAndQuote("^"));
                            break;
                        case VerticalAlignment.Center:
                        case VerticalAlignment.Stretch:
                            sb.Append(StringifyAndQuote("-"));
                            break;
                        case VerticalAlignment.Bottom:
                            sb.Append(StringifyAndQuote("v"));
                            break;
                    }
                }

                sb.Append(rowSep);
            }

            return sb.ToString();
        }

        private string StringifyAndQuote(object obj)
        {
            string cell = obj.ToString();
            if (string.IsNullOrEmpty(cell) || (!this.ColumnSeparator.IsMatch(cell) && !this.RowSeparator.IsMatch(cell)))
            {
                return cell;
            }
            else
            {
                return this.QuoteSymbol + cell + this.QuoteSymbol;
            }
        }

        /// <summary>
        /// Splits the specified <paramref name="text"/> according to the <paramref name="rowSeparator"/> and the <paramref name="columnSeparator"/>, with the specified <paramref name="quote"/> character.
        /// </summary>
        /// <param name="text">The text to split.</param>
        /// <param name="rowSeparator">The row separator (e.g., a newline character).</param>
        /// <param name="columnSeparator">The column separator (e.g., a tabulation character).</param>
        /// <param name="quote">The character(s) used to quote cells that contain separator characters within their value.</param>
        /// <param name="width">When this method returns, this value will contain the maximum width of the splitted text.</param>
        /// <returns>The splitted text, as a jagged array. Each element of the array represents a row. Rows may have different lengths.</returns>
        public static string[][] SplitData(string text, Regex rowSeparator, Regex columnSeparator, string quote, out int width)
        {
            List<string> lines = new List<string>(Split(text, rowSeparator, quote, columnSeparator));

            if (string.IsNullOrEmpty(lines[lines.Count - 1]))
            {
                lines.RemoveAt(lines.Count - 1);
            }

            int height = lines.Count;
            width = 0;

            string[][] cells = new string[height][];

            for (int i = 0; i < lines.Count; i++)
            {
                cells[i] = Split(lines[i] ?? "", columnSeparator, quote, rowSeparator);
                width = Math.Max(width, cells[i].Length);
            }

            Regex escapedQuote = new Regex(quote + quote, RegexOptions.Compiled);

            for (int i = 0; i < cells.Length; i++)
            {
                for (int j = 0; j < cells[i].Length; j++)
                {
                    if (cells[i][j].StartsWith(quote) && cells[i][j].EndsWith(quote))
                    {
                        cells[i][j] = cells[i][j].Substring(quote.Length, cells[i][j].Length - quote.Length * 2);
                    }

                    cells[i][j] = escapedQuote.Replace(cells[i][j], quote);
                }
            }

            return cells;
        }

        private static string[] Split(string text, Regex separator, string quote, Regex otherSeparator)
        {
            if (separator.ToString() == "\n")
            {
                text = text.Replace("\r", "");
            }

            //string startingQuotePattern = "(?<=[(?:" + separator.ToString() + ")(?:" + otherSeparator.ToString() + ")](?:" + Regex.Escape(quote) + Regex.Escape(quote) + ")*)" + Regex.Escape(quote) + "(?!" + Regex.Escape(quote) + ")";
            //string endingQuotePattern = "(?<=[^(?:" + Regex.Escape(quote) + ")](?:" + Regex.Escape(quote) + Regex.Escape(quote) + ")*)" + Regex.Escape(quote) + "(?=[(?:" + separator.ToString() + ")(?:" + otherSeparator.ToString() + ")])";

            string startingQuotePattern = "(?<=(?:" + separator.ToString() + "|" + otherSeparator.ToString() + ")(?:" + Regex.Escape(quote) + Regex.Escape(quote) + ")*)" + Regex.Escape(quote) + "(?!" + Regex.Escape(quote) + ")";
            string endingQuotePattern = "(?<=[^(?:" + Regex.Escape(quote) + ")](?:" + Regex.Escape(quote) + Regex.Escape(quote) + ")*)" + Regex.Escape(quote) + "(?=(?:" + separator.ToString() + "|" + otherSeparator.ToString() + "))";

            Regex startingQuoteRegex = new Regex(startingQuotePattern, RegexOptions.Compiled);
            Regex endingQuoteRegex = new Regex(endingQuotePattern, RegexOptions.Compiled);

            List<string> splitStartingQuote = startingQuoteRegex.Split(text).ToList();

            List<string> splitQuote = new List<string>(splitStartingQuote.Count * 2)
            {
                splitStartingQuote[0]
            };

            for (int i = 1; i < splitStartingQuote.Count; i++)
            {
                Match match = endingQuoteRegex.Match(splitStartingQuote[i]);

                if (match.Success)
                {
                    splitQuote.Add(splitStartingQuote[i].Substring(0, match.Index));
                    splitQuote.Add(splitStartingQuote[i].Substring(match.Index + match.Length));
                }
                else
                {
                    splitQuote[splitQuote.Count - 1] = splitQuote[splitQuote.Count - 1] + quote + splitStartingQuote[i];
                }
            }

            if (splitQuote.Count % 2 == 0)
            {
                splitQuote[splitQuote.Count - 2] = splitQuote[splitQuote.Count - 2] + quote + splitQuote[splitQuote.Count - 1];
                splitQuote.RemoveAt(splitQuote.Count - 1);
            }


            List<int> separators = new List<int>();
            List<int> separatorLengths = new List<int>();

            int currPos = 0;

            for (int i = 0; i < splitQuote.Count; i++)
            {
                if (i % 2 == 0)
                {
                    MatchCollection matches = separator.Matches(splitQuote[i]);

                    for (int j = 0; j < matches.Count; j++)
                    {
                        currPos += matches[j].Index - (j == 0 ? 0 : (matches[j - 1].Index + matches[j - 1].Length));
                        separators.Add(currPos);
                        separatorLengths.Add(matches[j].Length);
                        currPos += matches[j].Length;
                    }

                    if (matches.Count > 0)
                    {
                        currPos += splitQuote[i].Length - (matches[matches.Count - 1].Index + matches[matches.Count - 1].Length) + quote.Length;
                    }
                    else
                    {
                        currPos += splitQuote[i].Length + quote.Length;
                    }
                }
                else
                {
                    currPos += splitQuote[i].Length + quote.Length;
                }
            }

            string[] tbr = new string[separators.Count + 1];

            currPos = 0;

            for (int i = 0; i < separators.Count; i++)
            {
                tbr[i] = text.Substring(currPos, separators[i] - currPos);
                currPos = separators[i] + separatorLengths[i];
            }

            tbr[tbr.Length - 1] = text.Substring(currPos, text.Length - currPos);

            return tbr;
        }

        private void DeserializeFormat(string serializedFormat)
        {
            string[][] splittedFormat = SplitData(serializedFormat, this.RowSeparator, this.ColumnSeparator, this.QuoteSymbol, out int width);

            Table table = this.FindControl<Table>("ContentTable");

            for (int i = 0; i < splittedFormat.Length; i++)
            {
                if (splittedFormat[i][0] == "DEFAULT")
                {
                    try
                    {
                        table.Foreground = Brush.Parse(splittedFormat[i][1]);
                    }
                    catch { }

                    try
                    {
                        table.FontFamily = FontFamily.Parse(splittedFormat[i][2]);
                    }
                    catch { }

                    if (double.TryParse(splittedFormat[i][3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double fs))
                    {
                        this.FontSize = fs;
                    }

                    try
                    {
                        table.FontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), splittedFormat[i][4]);
                    }
                    catch { }

                    try
                    {
                        table.FontWeight = (FontWeight)Enum.Parse(typeof(FontWeight), splittedFormat[i][5]);
                    }
                    catch { }

                    if (double.TryParse(splittedFormat[i][6], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double cw))
                    {
                        table.DefaultColumnWidth = cw;
                    }

                    if (double.TryParse(splittedFormat[i][7], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double rh))
                    {
                        table.DefaultRowHeight = rh;
                    }

                    try
                    {
                        table.DefaultMargin = Thickness.Parse(splittedFormat[i][8]);
                    }
                    catch { }

                    switch (splittedFormat[i][9])
                    {
                        case ">":
                            table.DefaultTextAlignment = TextAlignment.Left;
                            break;
                        case "|":
                            table.DefaultTextAlignment = TextAlignment.Center;
                            break;
                        case "<":
                            table.DefaultTextAlignment = TextAlignment.Right;
                            break;
                    }

                    switch (splittedFormat[i][10])
                    {
                        case "^":
                            table.DefaultVerticalAlignment = VerticalAlignment.Top;
                            break;
                        case "-":
                            table.DefaultVerticalAlignment = VerticalAlignment.Center;
                            break;
                        case "v":
                            table.DefaultVerticalAlignment = VerticalAlignment.Bottom;
                            break;
                    }
                }
                else if (splittedFormat[i][0].StartsWith("Row"))
                {
                    if (int.TryParse(splittedFormat[i][0].Substring(4), out int row))
                    {
                        row--;

                        if (!string.IsNullOrEmpty(splittedFormat[i][1]))
                        {
                            try
                            {
                                table.RowForeground[row] = Brush.Parse(splittedFormat[i][1]);
                            }
                            catch { }
                        }

                        if (!string.IsNullOrEmpty(splittedFormat[i][2]) && double.TryParse(splittedFormat[i][2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double rh))
                        {
                            table.RowHeights[row] = rh;
                        }

                        if (!string.IsNullOrEmpty(splittedFormat[i][3]) && !string.IsNullOrEmpty(splittedFormat[i][4]) && !string.IsNullOrEmpty(splittedFormat[i][5]))
                        {
                            try
                            {
                                FontFamily ff = FontFamily.Parse(splittedFormat[i][3]);
                                FontStyle fs = (FontStyle)Enum.Parse(typeof(FontStyle), splittedFormat[i][4]);
                                FontWeight fw = (FontWeight)Enum.Parse(typeof(FontWeight), splittedFormat[i][5]);

                                table.RowTypefaces[row] = new Typeface(ff, fs, fw);
                            }
                            catch { }
                        }
                    }
                }
                else if (splittedFormat[i][0].StartsWith("Column"))
                {
                    try
                    {
                        int column = HorizontalHeader.FromLetter(splittedFormat[i][0].Substring(7)) - 1;

                        if (!string.IsNullOrEmpty(splittedFormat[i][1]))
                        {
                            try
                            {
                                table.ColumnForeground[column] = Brush.Parse(splittedFormat[i][1]);
                            }
                            catch { }
                        }

                        if (!string.IsNullOrEmpty(splittedFormat[i][2]) && double.TryParse(splittedFormat[i][2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double cw))
                        {
                            table.ColumnWidths[column] = cw;
                        }

                        if (!string.IsNullOrEmpty(splittedFormat[i][3]) && !string.IsNullOrEmpty(splittedFormat[i][4]) && !string.IsNullOrEmpty(splittedFormat[i][5]))
                        {
                            try
                            {
                                FontFamily ff = FontFamily.Parse(splittedFormat[i][3]);
                                FontStyle fs = (FontStyle)Enum.Parse(typeof(FontStyle), splittedFormat[i][4]);
                                FontWeight fw = (FontWeight)Enum.Parse(typeof(FontWeight), splittedFormat[i][5]);

                                table.ColumnTypefaces[column] = new Typeface(ff, fs, fw);
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
                else if (splittedFormat[i][0].StartsWith("Cell"))
                {
                    try
                    {
                        string cell = splittedFormat[i][0].Substring(5);

                        int pos = 0;
                        while (pos < cell.Length && (int)cell[pos] >= 65 && (int)cell[pos] <= 90)
                        {
                            pos++;
                        }

                        int x = HorizontalHeader.FromLetter(cell.Substring(0, pos)) - 1;
                        int y = int.Parse(cell.Substring(pos)) - 1;

                        if (!string.IsNullOrEmpty(splittedFormat[i][1]))
                        {
                            try
                            {
                                table.CellForeground[(x, y)] = Brush.Parse(splittedFormat[i][1]);
                            }
                            catch { }
                        }

                        if (!string.IsNullOrEmpty(splittedFormat[i][2]))
                        {
                            try
                            {
                                table.CellMargin[(x, y)] = Thickness.Parse(splittedFormat[i][2]);
                            }
                            catch { }
                        }

                        if (!string.IsNullOrEmpty(splittedFormat[i][3]) && !string.IsNullOrEmpty(splittedFormat[i][4]) && !string.IsNullOrEmpty(splittedFormat[i][5]))
                        {
                            try
                            {
                                FontFamily ff = FontFamily.Parse(splittedFormat[i][3]);
                                FontStyle fs = (FontStyle)Enum.Parse(typeof(FontStyle), splittedFormat[i][4]);
                                FontWeight fw = (FontWeight)Enum.Parse(typeof(FontWeight), splittedFormat[i][5]);

                                table.CellTypefaces[(x, y)] = new Typeface(ff, fs, fw);
                            }
                            catch { }
                        }

                        switch (splittedFormat[i][6])
                        {
                            case ">":
                                table.CellTextAlignment[(x, y)] = TextAlignment.Left;
                                break;
                            case "|":
                                table.CellTextAlignment[(x, y)] = TextAlignment.Center;
                                break;
                            case "<":
                                table.CellTextAlignment[(x, y)] = TextAlignment.Right;
                                break;
                        }

                        switch (splittedFormat[i][7])
                        {
                            case "^":
                                table.CellVerticalAlignment[(x, y)] = VerticalAlignment.Top;
                                break;
                            case "-":
                                table.CellVerticalAlignment[(x, y)] = VerticalAlignment.Center;
                                break;
                            case "v":
                                table.CellVerticalAlignment[(x, y)] = VerticalAlignment.Bottom;
                                break;
                        }
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Clear all the currently loaded data and load the supplied data into the spreadsheet.
        /// </summary>
        /// <param name="serializedData">The data to load (serialized according to the <see cref="ColumnSeparator"/> and the <see cref="RowSeparator"/>).</param>
        /// <param name="serializedFormat">The format for the data to load (serialized according to the <see cref="ColumnSeparator"/> and the <see cref="RowSeparator"/>).</param>
        public void Load(string serializedData, string serializedFormat)
        {
            string[][] splittedData = SplitData(serializedData, this.RowSeparator, this.ColumnSeparator, this.QuoteSymbol, out int width);
            Table table = this.FindControl<Table>("ContentTable");
            table.Data.Clear();
            table.CellForeground.Clear();
            table.CellMargin.Clear();
            table.CellTextAlignment.Clear();
            table.CellTypefaces.Clear();
            table.CellVerticalAlignment.Clear();

            table.RowForeground.Clear();
            table.RowHeights.Clear();
            table.RowTypefaces.Clear();

            table.ColumnForeground.Clear();
            table.ColumnTypefaces.Clear();
            table.ColumnWidths.Clear();

            for (int i = 0; i < splittedData.Length; i++)
            {
                for (int j = 0; j < splittedData[i].Length; j++)
                {
                    table.Data[(j, i)] = splittedData[i][j];
                }
            }

            DeserializeFormat(serializedFormat);

            this.ClearRedoStack();
            this.ClearUndoStack();

            table.InvalidateVisual();
            this.FindControl<HorizontalHeader>("HorizontalHeaderControl").InvalidateVisual();
            this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateMeasure();
            this.FindControl<VerticalHeader>("VerticalHeaderControl").InvalidateVisual();
            this.InvalidateVisual();
            this.Selection = ImmutableList.Create(new SelectionRange(0, 0));
        }

        /// <summary>
        /// Scrolls to the top left corner of the spreadsheet.
        /// </summary>
        public void ScrollTopLeft()
        {
            this.FindControl<ScrollBar>("HorizontalScrollBar").Value = 0;
            this.FindControl<ScrollBar>("VerticalScrollBar").Value = 0;
        }

        /// <summary>
        /// Returns a <see langword="string"/> that matches the current <see cref="ColumnSeparator"/>.
        /// </summary>
        /// <returns>A <see langword="string"/> that matches the current <see cref="ColumnSeparator"/>.</returns>
        public string GetColumnSeparator()
        {
            return new Xeger(Regex.Unescape(this.ColumnSeparator.ToString()), new Random(20230926)).Generate();
        }

        /// <summary>
        /// Returns a <see langword="string"/> that matches the current <see cref="RowSeparator"/>.
        /// </summary>
        /// <returns>A <see langword="string"/> that matches the current <see cref="RowSeparator"/>.</returns>
        public string GetRowSeparator()
        {
            return new Xeger(Regex.Unescape(this.RowSeparator.ToString()), new Random(20230926)).Generate();
        }
    }

    /// <summary>
    /// <see cref="EventArgs"/> for the <see cref="Spreadsheet.CellSizeChanged"/> event.
    /// </summary>
    public class CellSizeChangedEventArgs : EventArgs
    {
        /// <summary>
        /// The new width of the cell.
        /// </summary>
        public double Width { get; }

        /// <summary>
        /// The new height of the cell.
        /// </summary>
        public double Height { get; }

        /// <summary>
        /// The horizontal coordinate of the cell.
        /// </summary>
        public int Left { get; }

        /// <summary>
        /// The vertical coordinate of the cell.
        /// </summary>
        public int Top { get; }

        internal CellSizeChangedEventArgs(int left, int top, double width, double height) : base()
        {
            this.Left = left;
            this.Top = top;
            this.Width = width;
            this.Height = height;
        }
    }

    /// <summary>
    /// <see cref="EventArgs"/> for the <see cref="Spreadsheet.ColorDoubleTapped"/> event.
    /// </summary>
    public class ColorDoubleTappedEventArgs : EventArgs
    {
        /// <summary>
        /// The horizontal coordinate of the cell.
        /// </summary>
        public int Left { get; }

        /// <summary>
        /// The vertical coordinate of the cell.
        /// </summary>
        public int Top { get; }

        /// <summary>
        /// The colour contained in the cell.
        /// </summary>
        public Color Color { get; }

        /// <summary>
        /// Set this to <see langword="true"/> to signal that the event has been completely handled.
        /// </summary>
        public bool Handled { get; set; }

        internal ColorDoubleTappedEventArgs(int left, int top, Color color) : base()
        {
            Left = left;
            Top = top;
            Color = color;
            Handled = false;
        }
    }
}
