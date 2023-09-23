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

using System;
using System.Collections.Generic;
using System.Collections.Immutable;

namespace Spreadalonia
{
    /// <summary>
    /// Describes a rectangular selection area.
    /// </summary>
    public struct SelectionRange
    {
        /// <summary>
        /// The y coordinate of the top of the selection.
        /// </summary>
        public int Top { get; }

        /// <summary>
        /// The y coordinate of the bottom of the selection.
        /// </summary>
        public int Bottom { get; }

        /// <summary>
        /// The x coordinate of the left of the selection.
        /// </summary>
        public int Left { get; }

        /// <summary>
        /// The x coordinate of the right of the selection.
        /// </summary>
        public int Right { get; }

        /// <summary>
        /// The width of the selection.
        /// </summary>
        public int Width => Right - Left + 1;

        /// <summary>
        /// The height of the selection.
        /// </summary>
        public int Height => Bottom - Top + 1;

        /// <summary>
        /// Create a new <see cref="SelectionRange"/> representing the specified area.
        /// </summary>
        /// <param name="left">The x coordinate of the left of the selection.</param>
        /// <param name="top">The y coordinate of the top of the selection.</param>
        /// <param name="right">The x coordinate of the right of the selection.</param>
        /// <param name="bottom">The y coordinate of the bottom of the selection.</param>
        public SelectionRange(int left, int top, int right, int bottom)
        {
            this.Left = left;
            this.Top = top;
            this.Right = right;
            this.Bottom = bottom;
        }

        /// <summary>
        /// Create a new <see cref="SelectionRange"/> representing a single cell.
        /// </summary>
        /// <param name="x">The x coordinate of the cell.</param>
        /// <param name="y">The y coordinate of the cell.</param>
        public SelectionRange(int x, int y)
        {
            this.Left = x;
            this.Top = y;
            this.Right = x;
            this.Bottom = y;
        }

        /// <summary>
        /// Create a new <see cref="SelectionRange"/> representing a single cell.
        /// </summary>
        /// <param name="cell">The coordinates of the cell.</param>
        public SelectionRange((int x, int y) cell) : this(cell.x, cell.y) { }

        /// <summary>
        /// Checks whether this <see cref="SelectionRange"/> overlaps another <see cref="SelectionRange"/>.
        /// </summary>
        /// <param name="other">The other selection ra</param>
        /// <returns><see langword="true"/> if the two <see cref="SelectionRange"/>s overlap; <see langword="false"/> otherwise.</returns>
        public bool Overlaps(SelectionRange other)
        {
            return this.Left <= other.Right && this.Right >= other.Left && this.Top <= other.Bottom && this.Bottom >= other.Top;
        }

        /// <summary>
        /// Checks whether this <see cref="SelectionRange"/> contains the specified cell.
        /// </summary>
        /// <param name="x">The x coordinate of the cell.</param>
        /// <param name="y">The y coordinate of the cell.</param>
        /// <returns><see langword="true"/> if the <see cref="SelectionRange"/> contains the cell; <see langword="false"/> otherwise.</returns>
        public bool Contains(int x, int y)
        {
            return x >= this.Left && x <= this.Right && y >= this.Top && y <= this.Bottom;
        }

        /// <summary>
        /// Checks whether this <see cref="SelectionRange"/> contains the specified cell.
        /// </summary>
        /// <param name="cell">The coordinates of the cell.</param>
        /// <returns><see langword="true"/> if the <see cref="SelectionRange"/> contains the cell; <see langword="false"/> otherwise.</returns>
        public bool Contains((int x, int y) cell) => Contains(cell.x, cell.y);

        /// <summary>
        /// Computes the intersection of this <see cref="SelectionRange"/> with another one.
        /// </summary>
        /// <param name="other">The <see cref="SelectionRange"/> whose intersection with the current instance is sought.</param>
        /// <returns>A <see cref="SelectionRange"/> representing the intersection of the two <see cref="SelectionRange"/>s.</returns>
        public SelectionRange Intersection(SelectionRange other)
        {
            return new SelectionRange(Math.Max(this.Left, other.Left), Math.Max(this.Top, other.Top), Math.Min(this.Right, other.Right), Math.Min(this.Bottom, other.Bottom));
        }

        // Adapted from https://stackoverflow.com/questions/5144615/difference-xor-between-two-rectangles-as-rectangles
        internal IEnumerable<SelectionRange> Difference(SelectionRange other)
        {
            SelectionRange inter = this.Intersection(other);

            int a = Math.Min(this.Left, inter.Left);
            int b = Math.Max(this.Left, inter.Left);
            int c = Math.Min(this.Right, inter.Right) + 1;
            int d = Math.Max(this.Right, inter.Right) + 1;

            int e = Math.Min(this.Top, inter.Top);
            int f = Math.Max(this.Top, inter.Top);
            int g = Math.Min(this.Bottom, inter.Bottom) + 1;
            int h = Math.Max(this.Bottom, inter.Bottom) + 1;

            // X = intersection, 0-7 = possible difference areas
            // h ┌─┬─┬─┐
            // . │5│6│7│
            // g ├─┼─┼─┤
            // . │3│X│4│
            // f ├─┼─┼─┤
            // . │0│1│2│
            // e └─┴─┴─┘
            // . a b c d

            // we'll always have rectangles 1, 3, 4 and 6

            if (c - 1 - b >= 0 && f - 1 - e >= 0)
            {
                yield return new SelectionRange(b, e, c - 1, f - 1);
            }

            if (b - 1 - a >= 0 && g - 1 - f >= 0)
            {
                yield return new SelectionRange(a, f, b - 1, g - 1);
            }

            if (d - 1 - c >= 0 && g - 1 - f >= 0)
            {
                yield return new SelectionRange(c, f, d - 1, g - 1);
            }

            if (c - 1 - b >= 0 && h - 1 - g >= 0)
            {
                yield return new SelectionRange(b, g, c - 1, h - 1);
            }

            if (b - 1 - a >= 0 && f - 1 - e >= 0)
            {
                yield return new SelectionRange(a, e, b - 1, f - 1);
            }

            if (d - 1 - c >= 0 && h - 1 - g >= 0)
            {
                yield return new SelectionRange(c, g, d - 1, h - 1);
            }

            if (d - 1 - c >= 0 && f - 1 - e >= 0)
            {
                yield return new SelectionRange(c, e, d - 1, f - 1);
            }

            if (b - 1 - a >= 0 && h - 1 - g >= 0)
            {
                yield return new SelectionRange(a, g, b - 1, h - 1);
            }
        }

        /// <inheritdoc/>
        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is SelectionRange sel))
            {
                return false;
            }
            
            return sel.Left == this.Left && sel.Top == this.Top && sel.Right == this.Right && sel.Bottom == this.Bottom;
        }

        /// <inheritdoc/>
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 23;

                hash = hash * 31 + this.Left;
                hash = hash * 31 + this.Top;
                hash = hash * 31 + this.Right;
                hash = hash * 31 + this.Bottom;

                return hash;
            }
        }

        /// <inheritdoc/>
        public static bool operator ==(SelectionRange el1, SelectionRange el2)
        {
            if (el1 == null)
            {
                return el2 == null;
            }
            else
            {
                return el1.Equals(el2);
            }
        }

        /// <inheritdoc/>
        public static bool operator !=(SelectionRange el1, SelectionRange el2)
        {
            if (el1 == null)
            {
                return el2 != null;
            }
            else
            {
                return !el1.Equals(el2);
            }
        }
    }


    internal class StackFrame<TKey, TValue> where TValue : class
    {
        public ImmutableList<SelectionRange> Selection { get; internal set; }
        public Dictionary<TKey, TValue> PreviousData { get; }
        public Dictionary<TKey, TValue> NewData { get; }

        public StackFrame(ImmutableList<SelectionRange> selection, Dictionary<TKey, TValue> previousData, Dictionary<TKey, TValue> newData)
        {
            Selection = selection;
            PreviousData = previousData;
            NewData = newData;
        }
    }

    internal class Reference<T> where T : struct
    {
        public T Value { get; }
        public Reference(T value)
        {
            Value = value;
        }

        public static implicit operator T?(Reference<T> reference)
        {
            return reference?.Value;
        }

        public static implicit operator Reference<T>(T? nullable)
        {
            if (nullable == null)
            {
                return null;
            }
            else
            {
                return new Reference<T>(nullable.Value);
            }
        }
    }

    internal class ValueStackFrame<TKey, TValue> : StackFrame<TKey, Reference<TValue>> where TValue : struct
    {
        public ValueStackFrame(ImmutableList<SelectionRange> selection, Dictionary<TKey, Reference<TValue>> previousData, Dictionary<TKey, Reference<TValue>> newData) : base(selection, previousData, newData) { }
    }
}
