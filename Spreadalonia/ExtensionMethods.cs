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
using Avalonia;
using System;
using System.Collections.Generic;
using Avalonia.VisualTree;
using System.Collections.Immutable;
using System.Linq;
using System.Text.RegularExpressions;
using Avalonia.Input.Platform;
using System.Threading.Tasks;

namespace Spreadalonia
{
    internal static class ExtensionMethods
    {
        public static async Task<bool> ContainsText(this IClipboard clipboard)
        {
            string[] formats = await clipboard?.GetFormatsAsync();

            if (formats != null)
            {
                foreach (string format in formats)
                {
                    switch (format)
                    {
                        case "Text":
                        case "text":
                        case "public.text":
                        case "public.plain-text":
                        case "public.utf8-plain-text":
                        case "public.utf16-plain-text":
                        case "public.utf16-external-plain-text":
                            return true;
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        }

        // Adapted from https://github.com/AvaloniaUI/AvaloniaEdit/blob/master/src/AvaloniaEdit/Utils/ExtensionMethods.cs
        public static Point SnapToDevicePixels(this Point p, Visual targetVisual)
        {
            var root = targetVisual.GetVisualRoot();

            // Get the root control and its scaling
            var scaling = new Vector(root.RenderScaling, root.RenderScaling);

            // Create a matrix to translate from control coordinates to device coordinates.
            var m = targetVisual.TransformToVisual((Control)root) * Matrix.CreateScale(scaling);

            if (m == null)
                return p;

            // Translate the point to device coordinates.
            var devicePoint = p.Transform(m.Value);

            // Snap the coordinate to the midpoint between device pixels.
            devicePoint = new Point(((int)devicePoint.X) + 0.5, ((int)devicePoint.Y) + 0.5);

            // Translate the point back to control coordinates.
            var inv = m.Value.Invert();
            Point result = devicePoint.Transform(inv);
            return result;
        }

        public static Point SnapToDevicePixels(this Point p, Visual targetVisual, bool evenX, bool evenY, double offsetX = 0, double offsetY = 0)
        {
            var root = targetVisual.GetVisualRoot();

            // Get the root control and its scaling
            var scaling = new Vector(root.RenderScaling, root.RenderScaling);

            // Create a matrix to translate from control coordinates to device coordinates.
            var m = targetVisual.TransformToVisual((Control)root) * Matrix.CreateScale(scaling);

            if (m == null)
                return p;

            // Translate the point to device coordinates.
            var devicePoint = new Point(p.X + offsetX, p.Y + offsetY).Transform(m.Value);

            double newX;
            double newY;

            if (evenX)
            {
                newX = ((int)devicePoint.X);
            }
            else
            {
                newX = ((int)devicePoint.X) + 0.5;
            }

            if (evenY)
            {
                newY = ((int)devicePoint.Y);
            }
            else
            {
                newY = ((int)devicePoint.Y) + 0.5;
            }

            devicePoint = new Point(newX - offsetX, newY - offsetY);

            // Translate the point back to control coordinates.
            var inv = m.Value.Invert();
            Point result = devicePoint.Transform(inv);
            return result;
        }


        public static ImmutableList<SelectionRange> Difference(this ImmutableList<SelectionRange> selection, SelectionRange toRemove)
        {
            List<SelectionRange> tbr = new List<SelectionRange>();

            for (int i = 0; i < selection.Count; i++)
            {
                if (selection[i] != toRemove)
                {
                    if (selection[i].Overlaps(toRemove))
                    {
                        tbr.AddRange(selection[i].Difference(toRemove).Consolidate());
                    }
                    else
                    {
                        tbr.Add(selection[i]);
                    }
                }
            }

            return tbr.ToImmutableList();
        }

        public static List<SelectionRange> Consolidate(this IEnumerable<SelectionRange> selections)
        {
            List<SelectionRange> tbr = new List<SelectionRange>(selections);

            bool foundAny = true;

            while (foundAny)
            {
                foundAny = false;
                for (int i = 0; i < tbr.Count; i++)
                {
                    for (int j = 0; j < tbr.Count; j++)
                    {
                        if ((tbr[i].Right == tbr[j].Left - 1 && tbr[i].Top == tbr[j].Top && tbr[i].Bottom == tbr[j].Bottom) ||
                            (tbr[i].Bottom == tbr[j].Top - 1 && tbr[i].Left == tbr[j].Left && tbr[i].Right == tbr[j].Right))
                        {
                            tbr[i] = new SelectionRange(tbr[i].Left, tbr[i].Top, tbr[j].Right, tbr[j].Bottom);
                            tbr.RemoveAt(j);
                            foundAny = true;
                            break;
                        }
                    }

                    if (foundAny)
                    {
                        break;
                    }
                }
            }

            return tbr;
        }

        public static List<SelectionRange> RemoveDuplicates(this IEnumerable<SelectionRange> selections)
        {

            bool anyOverlapAtAll;

            List<SelectionRange> tbr;

            do
            {
                anyOverlapAtAll = false;
                tbr = new List<SelectionRange>();

                foreach (SelectionRange range in selections)
                {
                    bool anyOverlap = false;

                    for (int i = 0; i < tbr.Count; i++)
                    {
                        if (range.Overlaps(tbr[i]))
                        {
                            anyOverlap = true;
                            anyOverlapAtAll = true;
                            tbr.AddRange(range.Difference(tbr[i]));
                            break;
                        }
                    }

                    if (!anyOverlap)
                    {
                        tbr.Add(range);
                    }
                }

                selections = tbr;
            } while (anyOverlapAtAll);

            return tbr;
        }


        public static void Move<T>(this Dictionary<(int, int), T> data, SelectionRange dataToMove, int deltaX, int deltaY, Stack<StackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : class
        {
            Dictionary<(int, int), T> dataBeingMoved = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> originalData = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>();

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (dataToMove.Contains(kvp.Key))
                {
                    dataBeingMoved.Add(kvp.Key, kvp.Value);
                }
            }

            for (int x = dataToMove.Left; x <= dataToMove.Right; x++)
            {
                for (int y = dataToMove.Top; y <= dataToMove.Bottom; y++)
                {
                    if (data.TryGetValue((x + deltaX, y + deltaY), out T val))
                    {
                        if (undoStack != null)
                        {
                            originalData.Add((x + deltaX, y + deltaY), val);
                            newData[(x + deltaX, y + deltaY)] = null;
                        }

                        data.Remove((x + deltaX, y + deltaY));
                    }
                    else if (undoStack != null)
                    {
                        originalData.Add((x + deltaX, y + deltaY), null);
                    }

                    if (!leaveOriginal)
                    {
                        data.Remove((x, y));

                        if (undoStack != null)
                        {
                            newData[(x, y)] = null;
                        }
                    }
                }
            }

            foreach (KeyValuePair<(int, int), T> kvp in dataBeingMoved)
            {
                data.Add((kvp.Key.Item1 + deltaX, kvp.Key.Item2 + deltaY), kvp.Value);
                newData[(kvp.Key.Item1 + deltaX, kvp.Key.Item2 + deltaY)] = kvp.Value;
            }

            if (undoStack != null)
            {
                foreach (KeyValuePair<(int, int), T> kvp in originalData)
                {
                    if (!dataBeingMoved.ContainsKey(kvp.Key))
                    {
                        dataBeingMoved[kvp.Key] = kvp.Value;
                    }
                }

                StackFrame<(int, int), T> frame = new StackFrame<(int, int), T>(ImmutableList.Create(dataToMove), dataBeingMoved, newData);
                undoStack.Push(frame);
            }
        }

        public static void Move<T>(this Dictionary<(int, int), T> data, SelectionRange dataToMove, int deltaX, int deltaY, Stack<ValueStackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : struct
        {
            Dictionary<(int, int), Reference<T>> dataBeingMoved = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> originalData = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> newData = new Dictionary<(int, int), Reference<T>>();

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (dataToMove.Contains(kvp.Key))
                {
                    dataBeingMoved.Add(kvp.Key, kvp.Value);
                }
            }

            for (int x = dataToMove.Left; x <= dataToMove.Right; x++)
            {
                for (int y = dataToMove.Top; y <= dataToMove.Bottom; y++)
                {
                    if (data.TryGetValue((x + deltaX, y + deltaY), out T val))
                    {
                        if (undoStack != null)
                        {
                            originalData.Add((x + deltaX, y + deltaY), val);
                            newData[(x + deltaX, y + deltaY)] = null;
                        }

                        data.Remove((x + deltaX, y + deltaY));
                    }
                    else if (undoStack != null)
                    {
                        originalData.Add((x + deltaX, y + deltaY), null);
                    }

                    if (!leaveOriginal)
                    {
                        data.Remove((x, y));

                        if (undoStack != null)
                        {
                            newData[(x, y)] = null;
                        }
                    }
                }
            }

            foreach (KeyValuePair<(int, int), Reference<T>> kvp in dataBeingMoved)
            {
                data.Add((kvp.Key.Item1 + deltaX, kvp.Key.Item2 + deltaY), kvp.Value.Value);
                newData[(kvp.Key.Item1 + deltaX, kvp.Key.Item2 + deltaY)] = kvp.Value.Value;
            }

            if (undoStack != null)
            {
                foreach (KeyValuePair<(int, int), Reference<T>> kvp in originalData)
                {
                    if (!dataBeingMoved.ContainsKey(kvp.Key))
                    {
                        dataBeingMoved[kvp.Key] = kvp.Value;
                    }
                }
                ValueStackFrame<(int, int), T> frame = new ValueStackFrame<(int, int), T>(ImmutableList.Create(dataToMove), dataBeingMoved, newData);
                undoStack.Push(frame);
            }
        }

        public static void MoveRows<T>(this Dictionary<(int, int), T> data, int minRange, int maxRange, int deltaY, Stack<StackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : class
        {
            Dictionary<(int, int), T> dataBeingMoved = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> originalData = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>();

            int minX = int.MaxValue;
            int maxX = 0;

            List<(int, int)> itemsToRemove = new List<(int, int)>();

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 >= minRange && kvp.Key.Item2 <= maxRange)
                {
                    dataBeingMoved.Add(kvp.Key, kvp.Value);

                    if (!leaveOriginal)
                    {
                        itemsToRemove.Add(kvp.Key);
                    }

                    minX = Math.Min(minX, kvp.Key.Item1);
                    maxX = Math.Max(maxX, kvp.Key.Item1);
                }
                else if (kvp.Key.Item2 >= minRange + deltaY && kvp.Key.Item2 <= maxRange + deltaY)
                {
                    itemsToRemove.Add(kvp.Key);

                    if (undoStack != null)
                    {
                        originalData.Add(kvp.Key, kvp.Value);

                        minX = Math.Min(minX, kvp.Key.Item1);
                        maxX = Math.Max(maxX, kvp.Key.Item1);
                    }
                }
            }

            for (int i = 0; i < itemsToRemove.Count; i++)
            {
                data.Remove(itemsToRemove[i]);
                newData[itemsToRemove[i]] = null;
            }

            foreach (KeyValuePair<(int, int), T> kvp in dataBeingMoved)
            {
                data.Add((kvp.Key.Item1, kvp.Key.Item2 + deltaY), kvp.Value);
                newData[(kvp.Key.Item1, kvp.Key.Item2 + deltaY)] = kvp.Value;
            }

            if (undoStack != null)
            {
                foreach (KeyValuePair<(int, int), T> kvp in dataBeingMoved)
                {
                    if (!originalData.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 + deltaY)))
                    {
                        originalData[(kvp.Key.Item1, kvp.Key.Item2 + deltaY)] = default(T);
                    }
                }

                foreach (KeyValuePair<(int, int), T> kvp in originalData)
                {
                    if (!dataBeingMoved.ContainsKey(kvp.Key))
                    {
                        dataBeingMoved[kvp.Key] = kvp.Value;
                    }
                }

                StackFrame<(int, int), T> frame = new StackFrame<(int, int), T>(ImmutableList.Create(new SelectionRange(minX, minRange, maxX, maxRange)), dataBeingMoved, newData);
                undoStack.Push(frame);
            }
        }

        public static void MoveRows<T>(this Dictionary<(int, int), T> data, int minRange, int maxRange, int deltaY, Stack<ValueStackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : struct
        {
            Dictionary<(int, int), Reference<T>> dataBeingMoved = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> originalData = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> newData = new Dictionary<(int, int), Reference<T>>();

            int minX = int.MaxValue;
            int maxX = 0;

            List<(int, int)> itemsToRemove = new List<(int, int)>();

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 >= minRange && kvp.Key.Item2 <= maxRange)
                {
                    dataBeingMoved.Add(kvp.Key, kvp.Value);

                    if (!leaveOriginal)
                    {
                        itemsToRemove.Add(kvp.Key);
                    }

                    minX = Math.Min(minX, kvp.Key.Item1);
                    maxX = Math.Max(maxX, kvp.Key.Item1);
                }
                else if (kvp.Key.Item2 >= minRange + deltaY && kvp.Key.Item2 <= maxRange + deltaY)
                {
                    itemsToRemove.Add(kvp.Key);

                    if (undoStack != null)
                    {
                        originalData.Add(kvp.Key, kvp.Value);

                        minX = Math.Min(minX, kvp.Key.Item1);
                        maxX = Math.Max(maxX, kvp.Key.Item1);
                    }
                }
            }

            for (int i = 0; i < itemsToRemove.Count; i++)
            {
                data.Remove(itemsToRemove[i]);
                newData[itemsToRemove[i]] = null;
            }

            foreach (KeyValuePair<(int, int), Reference<T>> kvp in dataBeingMoved)
            {
                data.Add((kvp.Key.Item1, kvp.Key.Item2 + deltaY), kvp.Value.Value);
                newData[(kvp.Key.Item1, kvp.Key.Item2 + deltaY)] = kvp.Value.Value;
            }

            if (undoStack != null)
            {
                foreach (KeyValuePair<(int, int), Reference<T>> kvp in dataBeingMoved)
                {
                    if (!originalData.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 + deltaY)))
                    {
                        originalData[(kvp.Key.Item1, kvp.Key.Item2 + deltaY)] = null;
                    }
                }

                foreach (KeyValuePair<(int, int), Reference<T>> kvp in originalData)
                {
                    if (!dataBeingMoved.ContainsKey(kvp.Key))
                    {
                        dataBeingMoved[kvp.Key] = kvp.Value;
                    }
                }

                ValueStackFrame<(int, int), T> frame = new ValueStackFrame<(int, int), T>(ImmutableList.Create(new SelectionRange(minX, minRange, maxX, maxRange)), dataBeingMoved, newData);
                undoStack.Push(frame);
            }
        }

        public static void MoveColumns<T>(this Dictionary<(int, int), T> data, int minRange, int maxRange, int deltaX, Stack<StackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : class
        {
            Dictionary<(int, int), T> dataBeingMoved = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> originalData = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>();

            int minY = int.MaxValue;
            int maxY = 0;

            List<(int, int)> itemsToRemove = new List<(int, int)>();

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 >= minRange && kvp.Key.Item1 <= maxRange)
                {
                    dataBeingMoved.Add(kvp.Key, kvp.Value);

                    if (!leaveOriginal)
                    {
                        itemsToRemove.Add(kvp.Key);
                    }

                    minY = Math.Min(minY, kvp.Key.Item2);
                    maxY = Math.Max(maxY, kvp.Key.Item2);
                }
                else if (kvp.Key.Item1 >= minRange + deltaX && kvp.Key.Item1 <= maxRange + deltaX)
                {
                    itemsToRemove.Add(kvp.Key);

                    if (undoStack != null)
                    {
                        originalData.Add(kvp.Key, kvp.Value);

                        minY = Math.Min(minY, kvp.Key.Item2);
                        maxY = Math.Max(maxY, kvp.Key.Item2);
                    }
                }
            }

            for (int i = 0; i < itemsToRemove.Count; i++)
            {
                data.Remove(itemsToRemove[i]);
                newData[itemsToRemove[i]] = null;
            }

            foreach (KeyValuePair<(int, int), T> kvp in dataBeingMoved)
            {
                data.Add((kvp.Key.Item1 + deltaX, kvp.Key.Item2), kvp.Value);
                newData[(kvp.Key.Item1 + deltaX, kvp.Key.Item2)] = kvp.Value;
            }

            if (undoStack != null)
            {
                foreach (KeyValuePair<(int, int), T> kvp in dataBeingMoved)
                {
                    if (!originalData.ContainsKey((kvp.Key.Item1 + deltaX, kvp.Key.Item2)))
                    {
                        originalData[(kvp.Key.Item1 + deltaX, kvp.Key.Item2)] = default(T);
                    }
                }

                foreach (KeyValuePair<(int, int), T> kvp in originalData)
                {
                    if (!dataBeingMoved.ContainsKey(kvp.Key))
                    {
                        dataBeingMoved[kvp.Key] = kvp.Value;
                    }
                }

                StackFrame<(int, int), T> frame = new StackFrame<(int, int), T>(ImmutableList.Create(new SelectionRange(minRange, minY, maxRange, maxY)), dataBeingMoved, newData);
                undoStack.Push(frame);
            }
        }

        public static void MoveColumns<T>(this Dictionary<(int, int), T> data, int minRange, int maxRange, int deltaX, Stack<ValueStackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : struct
        {
            Dictionary<(int, int), Reference<T>> dataBeingMoved = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> originalData = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> newData = new Dictionary<(int, int), Reference<T>>();

            int minY = int.MaxValue;
            int maxY = 0;

            List<(int, int)> itemsToRemove = new List<(int, int)>();

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 >= minRange && kvp.Key.Item1 <= maxRange)
                {
                    dataBeingMoved.Add(kvp.Key, kvp.Value);

                    if (!leaveOriginal)
                    {
                        itemsToRemove.Add(kvp.Key);
                    }

                    minY = Math.Min(minY, kvp.Key.Item2);
                    maxY = Math.Max(maxY, kvp.Key.Item2);
                }
                else if (kvp.Key.Item1 >= minRange + deltaX && kvp.Key.Item1 <= maxRange + deltaX)
                {
                    itemsToRemove.Add(kvp.Key);

                    if (undoStack != null)
                    {
                        originalData.Add(kvp.Key, kvp.Value);

                        minY = Math.Min(minY, kvp.Key.Item2);
                        maxY = Math.Max(maxY, kvp.Key.Item2);
                    }
                }
            }

            for (int i = 0; i < itemsToRemove.Count; i++)
            {
                data.Remove(itemsToRemove[i]);
                newData[itemsToRemove[i]] = null;
            }

            foreach (KeyValuePair<(int, int), Reference<T>> kvp in dataBeingMoved)
            {
                data.Add((kvp.Key.Item1 + deltaX, kvp.Key.Item2), kvp.Value.Value);
                newData[(kvp.Key.Item1 + deltaX, kvp.Key.Item2)] = kvp.Value.Value;
            }

            if (undoStack != null)
            {
                foreach (KeyValuePair<(int, int), Reference<T>> kvp in dataBeingMoved)
                {
                    if (!originalData.ContainsKey((kvp.Key.Item1 + deltaX, kvp.Key.Item2)))
                    {
                        originalData[(kvp.Key.Item1 + deltaX, kvp.Key.Item2)] = default(T);
                    }
                }

                foreach (KeyValuePair<(int, int), Reference<T>> kvp in originalData)
                {
                    if (!dataBeingMoved.ContainsKey(kvp.Key))
                    {
                        dataBeingMoved[kvp.Key] = kvp.Value;
                    }
                }

                ValueStackFrame<(int, int), T> frame = new ValueStackFrame<(int, int), T>(ImmutableList.Create(new SelectionRange(minRange, minY, maxRange, maxY)), dataBeingMoved, newData);
                undoStack.Push(frame);
            }
        }


        public static void Move<T>(this Dictionary<int, T> data, int minRange, int maxRange, int delta, Stack<StackFrame<int, T>> undoStack, bool leaveOriginal) where T : class
        {
            Dictionary<int, T> dataBeingMoved = new Dictionary<int, T>();
            Dictionary<int, T> originalData = new Dictionary<int, T>();
            Dictionary<int, T> newData = new Dictionary<int, T>();

            foreach (KeyValuePair<int, T> kvp in data)
            {
                if (kvp.Key >= minRange && kvp.Key <= maxRange)
                {
                    dataBeingMoved.Add(kvp.Key, kvp.Value);
                }
            }

            for (int x = minRange; x <= maxRange; x++)
            {
                if (data.TryGetValue(x + delta, out T val))
                {
                    if (undoStack != null)
                    {
                        originalData.Add(x + delta, val);
                        newData[x + delta] = null;
                    }

                    data.Remove(x + delta);
                }
                else if (undoStack != null)
                {
                    originalData.Add(x + delta, null);
                }

                if (!leaveOriginal)
                {
                    data.Remove(x);
                    newData[x] = null;
                }
            }

            foreach (KeyValuePair<int, T> kvp in dataBeingMoved)
            {
                data.Add(kvp.Key + delta, kvp.Value);
                newData[kvp.Key + delta] = kvp.Value;
            }

            if (undoStack != null)
            {
                foreach (KeyValuePair<int, T> kvp in originalData)
                {
                    if (!dataBeingMoved.ContainsKey(kvp.Key))
                    {
                        dataBeingMoved[kvp.Key] = kvp.Value;
                    }
                }
                StackFrame<int, T> frame = new StackFrame<int, T>(ImmutableList.Create(new SelectionRange(minRange, minRange, maxRange, maxRange)), dataBeingMoved, newData);
                undoStack.Push(frame);
            }
        }

        public static void Move<T>(this Dictionary<int, T> data, int minRange, int maxRange, int delta, Stack<ValueStackFrame<int, T>> undoStack, bool leaveOriginal) where T : struct
        {
            Dictionary<int, Reference<T>> dataBeingMoved = new Dictionary<int, Reference<T>>();
            Dictionary<int, Reference<T>> originalData = new Dictionary<int, Reference<T>>();
            Dictionary<int, Reference<T>> newData = new Dictionary<int, Reference<T>>();

            foreach (KeyValuePair<int, T> kvp in data)
            {
                if (kvp.Key >= minRange && kvp.Key <= maxRange)
                {
                    dataBeingMoved.Add(kvp.Key, kvp.Value);
                }
            }

            for (int x = minRange; x <= maxRange; x++)
            {
                if (data.TryGetValue(x + delta, out T val))
                {
                    if (undoStack != null)
                    {
                        originalData.Add(x + delta, val);
                    }

                    data.Remove(x + delta);
                    newData[x + delta] = null;
                }
                else if (undoStack != null)
                {
                    originalData.Add(x + delta, null);
                }

                if (!leaveOriginal)
                {
                    data.Remove(x);
                    newData[x] = null;
                }
            }

            foreach (KeyValuePair<int, Reference<T>> kvp in dataBeingMoved)
            {
                data.Add(kvp.Key + delta, kvp.Value.Value);
                newData[kvp.Key + delta] = kvp.Value.Value;
            }

            if (undoStack != null)
            {
                foreach (KeyValuePair<int, Reference<T>> kvp in originalData)
                {
                    if (!dataBeingMoved.ContainsKey(kvp.Key))
                    {
                        dataBeingMoved[kvp.Key] = kvp.Value;
                    }
                }
                ValueStackFrame<int, T> frame = new ValueStackFrame<int, T>(ImmutableList.Create(new SelectionRange(minRange, minRange, maxRange, maxRange)), dataBeingMoved, newData);
                undoStack.Push(frame);
            }
        }

        public static bool IsFinite(this SelectionRange range, Table tab)
        {
            return (range.Bottom < tab.MaxTableHeight || range.Top > 0) && (range.Right < tab.MaxTableWidth || range.Left > 0);
        }

        public static bool IsRows(this SelectionRange range, Table tab)
        {
            return range.Left == 0 && range.Right == tab.MaxTableWidth;
        }

        public static bool IsColumns(this SelectionRange range, Table tab)
        {
            return range.Top == 0 && range.Bottom == tab.MaxTableHeight;
        }

        public static int GetBoundaryDown<T>(this Dictionary<(int, int), T> data, int minX, int maxX, int y)
        {
            int tbr = int.MaxValue;
            bool found = false;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 >= minX && kvp.Key.Item1 <= maxX && kvp.Key.Item2 > y)
                {
                    if (!data.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 + 1)) || !data.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 - 1)))
                    {
                        found = true;
                        tbr = Math.Min(tbr, kvp.Key.Item2);
                    }
                }
            }

            if (found)
            {
                return tbr;
            }
            else
            {
                return -1;
            }
        }

        public static int GetBoundaryUp<T>(this Dictionary<(int, int), T> data, int minX, int maxX, int y)
        {
            int tbr = int.MinValue;
            bool found = false;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 >= minX && kvp.Key.Item1 <= maxX && kvp.Key.Item2 < y)
                {
                    if (!data.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 + 1)) || !data.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 - 1)))
                    {
                        found = true;
                        tbr = Math.Max(tbr, kvp.Key.Item2);
                    }
                }
            }

            if (found)
            {
                return tbr;
            }
            else
            {
                return -1;
            }
        }

        public static int GetBoundaryRight<T>(this Dictionary<(int, int), T> data, int x, int minY, int maxY)
        {
            int tbr = int.MaxValue;
            bool found = false;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 >= minY && kvp.Key.Item2 <= maxY && kvp.Key.Item1 > x)
                {
                    if (!data.ContainsKey((kvp.Key.Item1 + 1, kvp.Key.Item2)) || !data.ContainsKey((kvp.Key.Item1 - 1, kvp.Key.Item2)))
                    {
                        found = true;
                        tbr = Math.Min(tbr, kvp.Key.Item1);
                    }
                }
            }

            if (found)
            {
                return tbr;
            }
            else
            {
                return -1;
            }
        }

        public static int GetBoundaryLeft<T>(this Dictionary<(int, int), T> data, int x, int minY, int maxY)
        {
            int tbr = int.MinValue;
            bool found = false;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 >= minY && kvp.Key.Item2 <= maxY && kvp.Key.Item1 < x)
                {
                    if (!data.ContainsKey((kvp.Key.Item1 + 1, kvp.Key.Item2)) || !data.ContainsKey((kvp.Key.Item1 - 1, kvp.Key.Item2)))
                    {
                        found = true;
                        tbr = Math.Max(tbr, kvp.Key.Item1);
                    }
                }
            }

            if (found)
            {
                return tbr;
            }
            else
            {
                return -1;
            }
        }

        public static Dictionary<(int, int), T> Remove<T>(this Dictionary<(int, int), T> data, SelectionRange range, Stack<StackFrame<(int, int), T>> undoStack) where T : class
        {
            Dictionary<(int, int), T> tmp = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> removedData = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>();

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (!(kvp.Key.Item1 >= range.Left && kvp.Key.Item1 <= range.Right && kvp.Key.Item2 >= range.Top && kvp.Key.Item2 <= range.Bottom))
                {
                    tmp.Add(kvp.Key, kvp.Value);
                }
                else if (undoStack != null)
                {
                    removedData.Add(kvp.Key, kvp.Value);
                    newData.Add(kvp.Key, null);
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<(int, int), T>(ImmutableList.Create(range), removedData, newData));
            }

            return tmp;
        }

        public static Dictionary<(int, int), T> Remove<T>(this Dictionary<(int, int), T> data, SelectionRange range, Stack<ValueStackFrame<(int, int), T>> undoStack) where T : struct
        {
            Dictionary<(int, int), T> tmp = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), Reference<T>> removedData = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> newData = new Dictionary<(int, int), Reference<T>>();

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (!(kvp.Key.Item1 >= range.Left && kvp.Key.Item1 <= range.Right && kvp.Key.Item2 >= range.Top && kvp.Key.Item2 <= range.Bottom))
                {
                    tmp.Add(kvp.Key, kvp.Value);
                }
                else if (undoStack != null)
                {
                    removedData.Add(kvp.Key, kvp.Value);
                    newData.Add(kvp.Key, null);
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<(int, int), T>(ImmutableList.Create(range), removedData, newData));
            }

            return tmp;
        }

        public static Dictionary<(int, int), T> Remove<T>(this Dictionary<(int, int), T> data, IEnumerable<SelectionRange> ranges, Stack<StackFrame<(int, int), T>> undoStack) where T : class
        {
            List<StackFrame<(int, int), T>> frames = new List<StackFrame<(int, int), T>>();

            foreach (SelectionRange range in ranges)
            {
                data = Remove(data, range, undoStack);

                if (undoStack != null)
                {
                    frames.Add(undoStack.Pop());
                }
            }

            if (undoStack != null)
            {
                Dictionary<(int, int), T> allRemovedData = new Dictionary<(int, int), T>();
                Dictionary<(int, int), T> allNewData = new Dictionary<(int, int), T>();
                List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

                for (int i = 0; i < frames.Count; i++)
                {
                    foreach (KeyValuePair<(int, int), T> kvp in frames[i].PreviousData)
                    {
                        allRemovedData[kvp.Key] = kvp.Value;
                    }

                    foreach (KeyValuePair<(int, int), T> kvp in frames[i].NewData)
                    {
                        allNewData[kvp.Key] = kvp.Value;
                    }

                    undoStackRanges.AddRange(frames[i].Selection);
                }

                undoStack.Push(new StackFrame<(int, int), T>(undoStackRanges.ToImmutableList(), allRemovedData, allNewData));
            }

            return data;
        }

        public static Dictionary<(int, int), T> Remove<T>(this Dictionary<(int, int), T> data, IEnumerable<SelectionRange> ranges, Stack<ValueStackFrame<(int, int), T>> undoStack) where T : struct
        {
            List<ValueStackFrame<(int, int), T>> frames = new List<ValueStackFrame<(int, int), T>>();

            foreach (SelectionRange range in ranges)
            {
                data = Remove(data, range, undoStack);

                if (undoStack != null)
                {
                    frames.Add(undoStack.Pop());
                }
            }

            if (undoStack != null)
            {
                Dictionary<(int, int), Reference<T>> allRemovedData = new Dictionary<(int, int), Reference<T>>();
                Dictionary<(int, int), Reference<T>> allNewData = new Dictionary<(int, int), Reference<T>>();
                List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

                for (int i = 0; i < frames.Count; i++)
                {
                    foreach (KeyValuePair<(int, int), Reference<T>> kvp in frames[i].PreviousData)
                    {
                        allRemovedData[kvp.Key] = kvp.Value;
                    }

                    foreach (KeyValuePair<(int, int), Reference<T>> kvp in frames[i].NewData)
                    {
                        allNewData[kvp.Key] = kvp.Value;
                    }

                    undoStackRanges.AddRange(frames[i].Selection);
                }

                undoStack.Push(new ValueStackFrame<(int, int), T>(undoStackRanges.ToImmutableList(), allRemovedData, allNewData));
            }

            return data;
        }

        public static Dictionary<int, T> Remove<T>(this Dictionary<int, T> data, int min, int max, int tableMax, bool row, Stack<StackFrame<int, T>> undoStack) where T : class
        {
            Dictionary<int, T> tmp = new Dictionary<int, T>(data.Count);
            Dictionary<int, T> removedData = new Dictionary<int, T>();
            Dictionary<int, T> newData = new Dictionary<int, T>();

            foreach (KeyValuePair<int, T> kvp in data)
            {
                if (!(kvp.Key >= min && kvp.Key <= max))
                {
                    tmp.Add(kvp.Key, kvp.Value);
                }
                else if (undoStack != null)
                {
                    removedData.Add(kvp.Key, kvp.Value);
                    newData.Add(kvp.Key, null);
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<int, T>(ImmutableList.Create(new SelectionRange(row ? 0 : min, row ? min : 0, row ? tableMax : max, row ? tableMax : max)), removedData, newData));
            }

            return tmp;
        }

        public static Dictionary<int, T> Remove<T>(this Dictionary<int, T> data, int min, int max, int tableMax, bool row, Stack<ValueStackFrame<int, T>> undoStack) where T : struct
        {
            Dictionary<int, T> tmp = new Dictionary<int, T>(data.Count);
            Dictionary<int, Reference<T>> removedData = new Dictionary<int, Reference<T>>();
            Dictionary<int, Reference<T>> newData = new Dictionary<int, Reference<T>>();

            foreach (KeyValuePair<int, T> kvp in data)
            {
                if (!(kvp.Key >= min && kvp.Key <= max))
                {
                    tmp.Add(kvp.Key, kvp.Value);
                }
                else if (undoStack != null)
                {
                    removedData.Add(kvp.Key, kvp.Value);
                    newData.Add(kvp.Key, null);
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<int, T>(ImmutableList.Create(new SelectionRange(row ? 0 : min, row ? min : 0, row ? tableMax : max, row ? tableMax : max)), removedData, newData));
            }

            return tmp;
        }

        public static Dictionary<int, T> Remove<T>(this Dictionary<int, T> data, IEnumerable<(int min, int max)> ranges, int tableMax, bool row, Stack<StackFrame<int, T>> undoStack) where T : class
        {
            List<StackFrame<int, T>> frames = new List<StackFrame<int, T>>();

            foreach ((int min, int max) range in ranges)
            {
                data = Remove(data, range.min, range.max, tableMax, row, undoStack);

                if (undoStack != null)
                {
                    frames.Add(undoStack.Pop());
                }
            }

            if (undoStack != null)
            {
                Dictionary<int, T> allRemovedData = new Dictionary<int, T>();
                Dictionary<int, T> allNewData = new Dictionary<int, T>();
                List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

                for (int i = 0; i < frames.Count; i++)
                {
                    foreach (KeyValuePair<int, T> kvp in frames[i].PreviousData)
                    {
                        allRemovedData[kvp.Key] = kvp.Value;
                    }

                    foreach (KeyValuePair<int, T> kvp in frames[i].NewData)
                    {
                        allNewData[kvp.Key] = kvp.Value;
                    }

                    undoStackRanges.AddRange(frames[i].Selection);
                }

                undoStack.Push(new StackFrame<int, T>(undoStackRanges.ToImmutableList(), allRemovedData, allNewData));
            }

            return data;
        }

        public static Dictionary<int, T> Remove<T>(this Dictionary<int, T> data, IEnumerable<(int min, int max)> ranges, int tableMax, bool row, Stack<ValueStackFrame<int, T>> undoStack) where T : struct
        {
            List<ValueStackFrame<int, T>> frames = new List<ValueStackFrame<int, T>>();

            foreach ((int min, int max) range in ranges)
            {
                data = Remove(data, range.min, range.max, tableMax, row, undoStack);

                if (undoStack != null)
                {
                    frames.Add(undoStack.Pop());
                }
            }

            if (undoStack != null)
            {
                Dictionary<int, Reference<T>> allRemovedData = new Dictionary<int, Reference<T>>();
                Dictionary<int, Reference<T>> allNewData = new Dictionary<int, Reference<T>>();
                List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

                for (int i = 0; i < frames.Count; i++)
                {
                    foreach (KeyValuePair<int, Reference<T>> kvp in frames[i].PreviousData)
                    {
                        allRemovedData[kvp.Key] = kvp.Value;
                    }

                    foreach (KeyValuePair<int, Reference<T>> kvp in frames[i].NewData)
                    {
                        allNewData[kvp.Key] = kvp.Value;
                    }

                    undoStackRanges.AddRange(frames[i].Selection);
                }

                undoStack.Push(new ValueStackFrame<int, T>(undoStackRanges.ToImmutableList(), allRemovedData, allNewData));
            }

            return data;
        }

        public static void SetValue<T>(this Dictionary<(int, int), T> data, SelectionRange range, T newValue, Stack<StackFrame<(int, int), T>> undoStack) where T : class
        {
            Dictionary<(int, int), T> previousData = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>();

            for (int y = range.Top; y <= range.Bottom; y++)
            {
                for (int x = range.Left; x <= range.Right; x++)
                {
                    if (undoStack != null)
                    {
                        if (data.TryGetValue((x, y), out T prevVal))
                        {
                            previousData.Add((x, y), prevVal);
                        }
                        else
                        {
                            previousData.Add((x, y), null);
                        }
                        newData.Add((x, y), newValue);
                    }

                    data[(x, y)] = newValue;
                }
            }


            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<(int, int), T>(ImmutableList.Create(range), previousData, newData));
            }
        }

        public static void SetValue<T>(this Dictionary<(int, int), T> data, SelectionRange range, T newValue, Stack<ValueStackFrame<(int, int), T>> undoStack) where T : struct
        {
            Dictionary<(int, int), Reference<T>> previousData = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> newData = new Dictionary<(int, int), Reference<T>>();

            for (int y = range.Top; y <= range.Bottom; y++)
            {
                for (int x = range.Left; x <= range.Right; x++)
                {
                    if (undoStack != null)
                    {
                        if (data.TryGetValue((x, y), out T prevVal))
                        {
                            previousData.Add((x, y), prevVal);
                        }
                        else
                        {
                            previousData.Add((x, y), null);
                        }
                        newData.Add((x, y), newValue);
                    }

                    data[(x, y)] = newValue;
                }
            }


            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<(int, int), T>(ImmutableList.Create(range), previousData, newData));
            }
        }

        public static void SetValue<T>(this Dictionary<(int, int), T> data, IEnumerable<SelectionRange> ranges, T newValue, Stack<StackFrame<(int, int), T>> undoStack) where T : class
        {
            List<StackFrame<(int, int), T>> frames = new List<StackFrame<(int, int), T>>();

            foreach (SelectionRange range in ranges)
            {
                data.SetValue(range, newValue, undoStack);

                if (undoStack != null)
                {
                    frames.Add(undoStack.Pop());
                }
            }

            if (undoStack != null)
            {
                Dictionary<(int, int), T> allPreviousData = new Dictionary<(int, int), T>();
                Dictionary<(int, int), T> allNewData = new Dictionary<(int, int), T>();
                List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);


                for (int i = frames.Count - 1; i >= 0; i--)
                {
                    foreach (KeyValuePair<(int, int), T> kvp in frames[i].PreviousData)
                    {
                        allPreviousData[kvp.Key] = kvp.Value;
                    }

                    undoStackRanges.AddRange(frames[i].Selection);
                }

                for (int i = 0; i < frames.Count; i++)
                {
                    foreach (KeyValuePair<(int, int), T> kvp in frames[i].NewData)
                    {
                        allNewData[kvp.Key] = kvp.Value;
                    }
                }

                undoStack.Push(new StackFrame<(int, int), T>(undoStackRanges.ToImmutableList(), allPreviousData, allNewData));
            }
        }

        public static void SetValue<T>(this Dictionary<(int, int), T> data, IEnumerable<SelectionRange> ranges, T newValue, Stack<ValueStackFrame<(int, int), T>> undoStack) where T : struct
        {
            List<ValueStackFrame<(int, int), T>> frames = new List<ValueStackFrame<(int, int), T>>();

            foreach (SelectionRange range in ranges)
            {
                data.SetValue(range, newValue, undoStack);

                if (undoStack != null)
                {
                    frames.Add(undoStack.Pop());
                }
            }

            if (undoStack != null)
            {
                Dictionary<(int, int), Reference<T>> allPreviousData = new Dictionary<(int, int), Reference<T>>();
                Dictionary<(int, int), Reference<T>> allNewData = new Dictionary<(int, int), Reference<T>>();
                List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

                for (int i = 0; i < frames.Count; i++)
                {
                    foreach (KeyValuePair<(int, int), Reference<T>> kvp in frames[i].PreviousData)
                    {
                        allPreviousData[kvp.Key] = kvp.Value;
                    }

                    foreach (KeyValuePair<(int, int), Reference<T>> kvp in frames[i].NewData)
                    {
                        allNewData[kvp.Key] = kvp.Value;
                    }

                    undoStackRanges.AddRange(frames[i].Selection);
                }

                undoStack.Push(new ValueStackFrame<(int, int), T>(undoStackRanges.ToImmutableList(), allPreviousData, allNewData));
            }
        }

        public static void SetValue<T>(this Dictionary<int, T> data, int minRange, int maxRange, T newValue, int maxTable, bool row, Stack<StackFrame<int, T>> undoStack) where T : class
        {
            Dictionary<int, T> previousData = new Dictionary<int, T>();
            Dictionary<int, T> newData = new Dictionary<int, T>();


            for (int x = minRange; x <= maxRange; x++)
            {
                if (undoStack != null)
                {
                    if (data.TryGetValue(x, out T prevVal))
                    {
                        previousData.Add(x, prevVal);
                    }
                    newData.Add(x, newValue);
                }

                data[x] = newValue;
            }


            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<int, T>(ImmutableList.Create(new SelectionRange(row ? 0 : minRange, row ? minRange : 0, row ? maxTable : maxRange, row ? maxRange : maxTable)), previousData, newData));
            }
        }

        public static void SetValue<T>(this Dictionary<int, T> data, int minRange, int maxRange, T newValue, int maxTable, bool row, Stack<ValueStackFrame<int, T>> undoStack) where T : struct
        {
            Dictionary<int, Reference<T>> previousData = new Dictionary<int, Reference<T>>();
            Dictionary<int, Reference<T>> newData = new Dictionary<int, Reference<T>>();


            for (int x = minRange; x <= maxRange; x++)
            {
                if (undoStack != null)
                {
                    if (data.TryGetValue(x, out T prevVal))
                    {
                        previousData.Add(x, prevVal);
                    }
                    newData.Add(x, newValue);
                }

                data[x] = newValue;
            }


            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<int, T>(ImmutableList.Create(new SelectionRange(row ? 0 : minRange, row ? minRange : 0, row ? maxTable : maxRange, row ? maxRange : maxTable)), previousData, newData));
            }
        }

        public static void SetValue<T>(this Dictionary<int, T> data, IEnumerable<(int min, int max)> ranges, T newValue, int maxTable, bool row, Stack<StackFrame<int, T>> undoStack) where T : class
        {
            List<StackFrame<int, T>> frames = new List<StackFrame<int, T>>();

            foreach ((int min, int max) range in ranges)
            {
                data.SetValue(range.min, range.max, newValue, maxTable, row, undoStack);

                if (undoStack != null)
                {
                    frames.Add(undoStack.Pop());
                }
            }

            if (undoStack != null)
            {
                Dictionary<int, T> allPreviousData = new Dictionary<int, T>();
                Dictionary<int, T> allNewData = new Dictionary<int, T>();
                List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

                for (int i = 0; i < frames.Count; i++)
                {
                    foreach (KeyValuePair<int, T> kvp in frames[i].PreviousData)
                    {
                        allPreviousData[kvp.Key] = kvp.Value;
                    }

                    foreach (KeyValuePair<int, T> kvp in frames[i].NewData)
                    {
                        allNewData[kvp.Key] = kvp.Value;
                    }

                    undoStackRanges.AddRange(frames[i].Selection);
                }

                undoStack.Push(new StackFrame<int, T>(undoStackRanges.ToImmutableList(), allPreviousData, allNewData));
            }
        }

        public static void SetValue<T>(this Dictionary<int, T> data, IEnumerable<(int min, int max)> ranges, T newValue, int maxTable, bool row, Stack<ValueStackFrame<int, T>> undoStack) where T : struct
        {
            List<ValueStackFrame<int, T>> frames = new List<ValueStackFrame<int, T>>();

            foreach ((int min, int max) range in ranges)
            {
                data.SetValue(range.min, range.max, newValue, maxTable, row, undoStack);

                if (undoStack != null)
                {
                    frames.Add(undoStack.Pop());
                }
            }

            if (undoStack != null)
            {
                Dictionary<int, Reference<T>> allPreviousData = new Dictionary<int, Reference<T>>();
                Dictionary<int, Reference<T>> allNewData = new Dictionary<int, Reference<T>>();
                List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

                for (int i = 0; i < frames.Count; i++)
                {
                    foreach (KeyValuePair<int, Reference<T>> kvp in frames[i].PreviousData)
                    {
                        allPreviousData[kvp.Key] = kvp.Value;
                    }

                    foreach (KeyValuePair<int, Reference<T>> kvp in frames[i].NewData)
                    {
                        allNewData[kvp.Key] = kvp.Value;
                    }

                    undoStackRanges.AddRange(frames[i].Selection);
                }

                undoStack.Push(new ValueStackFrame<int, T>(undoStackRanges.ToImmutableList(), allPreviousData, allNewData));
            }
        }

        public static (int, int) BottomRight<T>(this Dictionary<(int, int), T> data)
        {
            int maxX = 0;
            int maxY = 0;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                maxX = Math.Max(kvp.Key.Item1, maxX);
                maxY = Math.Max(kvp.Key.Item2, maxY);
            }

            return (maxX, maxY);
        }

        public static int GetRight<T>(this Dictionary<(int, int), T> data, int minY, int maxY)
        {
            int maxX = 0;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 >= minY && kvp.Key.Item2 <= maxY)
                {
                    maxX = Math.Max(kvp.Key.Item1, maxX);
                }
            }

            return maxX;
        }

        public static int GetBottom<T>(this Dictionary<(int, int), T> data, int minX, int maxX)
        {
            int maxY = 0;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 >= minX && kvp.Key.Item1 <= maxX)
                {
                    maxY = Math.Max(kvp.Key.Item2, maxY);
                }
            }

            return maxY;
        }

        public static (int, int) TopLeft<T>(this Dictionary<(int, int), T> data, Table table)
        {
            int minX = table.MaxTableWidth;
            int minY = table.MaxTableHeight;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                minX = Math.Min(kvp.Key.Item1, minX);
                minY = Math.Min(kvp.Key.Item2, minY);
            }

            return (minX, minY);
        }

        public static SelectionRange SelectAll<T>(this Dictionary<(int, int), T> data, Table table)
        {
            int minX = table.MaxTableWidth;
            int minY = table.MaxTableHeight;
            int maxX = 0;
            int maxY = 0;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                minX = Math.Min(kvp.Key.Item1, minX);
                minY = Math.Min(kvp.Key.Item2, minY);

                maxX = Math.Max(kvp.Key.Item1, maxX);
                maxY = Math.Max(kvp.Key.Item2, maxY);
            }

            return new SelectionRange(minX, minY, maxX, maxY);
        }

        public static SelectionRange Merge(this IEnumerable<SelectionRange> ranges)
        {
            int minX = int.MaxValue;
            int minY = int.MaxValue;
            int maxX = int.MinValue;
            int maxY = int.MinValue;

            foreach (SelectionRange range in ranges)
            {
                minX = Math.Min(minX, range.Left);
                minY = Math.Min(minY, range.Top);
                maxX = Math.Max(maxX, range.Right);
                maxY = Math.Max(maxY, range.Bottom);
            }

            return new SelectionRange(minX, minY, maxX, maxY);
        }

        public static StackFrame<TKey, TValue> Pop<TKey, TValue>(this Stack<StackFrame<TKey, TValue>> stack, Dictionary<TKey, TValue> state, bool undo, ref ImmutableList<SelectionRange> outputRange) where TValue : class
        {
            StackFrame<TKey, TValue> frame = stack.Pop();

            if (frame != null)
            {
                Dictionary<TKey, TValue> dataToApply = undo ? frame.PreviousData : frame.NewData;

                foreach (KeyValuePair<TKey, TValue> data in dataToApply)
                {
                    if (data.Value != null)
                    {
                        state[data.Key] = data.Value;
                    }
                    else
                    {
                        state.Remove(data.Key);
                    }
                }

                if (outputRange == null)
                {
                    outputRange = frame.Selection;
                }
            }

            return frame;
        }

        public static ValueStackFrame<TKey, TValue> Pop<TKey, TValue>(this Stack<ValueStackFrame<TKey, TValue>> stack, Dictionary<TKey, TValue> state, bool undo, ref ImmutableList<SelectionRange> outputRange) where TValue : struct
        {
            ValueStackFrame<TKey, TValue> frame = stack.Pop();

            if (frame != null)
            {
                Dictionary<TKey, Reference<TValue>> dataToApply = undo ? frame.PreviousData : frame.NewData;

                foreach (KeyValuePair<TKey, Reference<TValue>> data in dataToApply)
                {
                    if (data.Value != null)
                    {
                        state[data.Key] = data.Value.Value;
                    }
                    else
                    {
                        state.Remove(data.Key);
                    }
                }

                if (outputRange == null)
                {
                    outputRange = frame.Selection;
                }
            }

            return frame;
        }

        public static Dictionary<(int, int), T> InsertColumns<T>(this Dictionary<(int, int), T> data, SelectionRange columns, Stack<StackFrame<(int, int), T>> undoStack) where T : class
        {
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> previousData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> changedData = new Dictionary<(int, int), T>(data.Count);

            int delta = columns.Width;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 < columns.Left)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else
                {
                    newData.Add((kvp.Key.Item1 + delta, kvp.Key.Item2), kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey((kvp.Key.Item1 + delta, kvp.Key.Item2)))
                        {
                            previousData[(kvp.Key.Item1 + delta, kvp.Key.Item2)] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[(kvp.Key.Item1 + delta, kvp.Key.Item2)] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<(int, int), T>(ImmutableList.Create(columns), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<(int, int), T> InsertColumns<T>(this Dictionary<(int, int), T> data, SelectionRange columns, Stack<ValueStackFrame<(int, int), T>> undoStack) where T : struct
        {
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), Reference<T>> previousData = new Dictionary<(int, int), Reference<T>>(data.Count);
            Dictionary<(int, int), Reference<T>> changedData = new Dictionary<(int, int), Reference<T>>(data.Count);

            int delta = columns.Width;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 < columns.Left)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else
                {
                    newData.Add((kvp.Key.Item1 + delta, kvp.Key.Item2), kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey((kvp.Key.Item1 + delta, kvp.Key.Item2)))
                        {
                            previousData[(kvp.Key.Item1 + delta, kvp.Key.Item2)] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[(kvp.Key.Item1 + delta, kvp.Key.Item2)] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<(int, int), T>(ImmutableList.Create(columns), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<(int, int), T> InsertRows<T>(this Dictionary<(int, int), T> data, SelectionRange rows, Stack<StackFrame<(int, int), T>> undoStack) where T : class
        {
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> previousData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> changedData = new Dictionary<(int, int), T>(data.Count);

            int delta = rows.Height;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 < rows.Top)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else
                {
                    newData.Add((kvp.Key.Item1, kvp.Key.Item2 + delta), kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 + delta)))
                        {
                            previousData[(kvp.Key.Item1, kvp.Key.Item2 + delta)] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[(kvp.Key.Item1, kvp.Key.Item2 + delta)] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<(int, int), T>(ImmutableList.Create(rows), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<(int, int), T> InsertRows<T>(this Dictionary<(int, int), T> data, SelectionRange rows, Stack<ValueStackFrame<(int, int), T>> undoStack) where T : struct
        {
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), Reference<T>> previousData = new Dictionary<(int, int), Reference<T>>(data.Count);
            Dictionary<(int, int), Reference<T>> changedData = new Dictionary<(int, int), Reference<T>>(data.Count);

            int delta = rows.Height;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 < rows.Top)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else
                {
                    newData.Add((kvp.Key.Item1, kvp.Key.Item2 + delta), kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 + delta)))
                        {
                            previousData[(kvp.Key.Item1, kvp.Key.Item2 + delta)] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[(kvp.Key.Item1, kvp.Key.Item2 + delta)] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<(int, int), T>(ImmutableList.Create(rows), previousData, changedData));
            }

            return newData;
        }


        public static Dictionary<int, T> Insert<T>(this Dictionary<int, T> data, int where, int howMany, SelectionRange selection, Stack<StackFrame<int, T>> undoStack) where T : class
        {
            Dictionary<int, T> newData = new Dictionary<int, T>(data.Count);
            Dictionary<int, T> previousData = new Dictionary<int, T>(data.Count);
            Dictionary<int, T> changedData = new Dictionary<int, T>(data.Count);

            foreach (KeyValuePair<int, T> kvp in data)
            {
                if (kvp.Key < where)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else
                {
                    newData.Add(kvp.Key + howMany, kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey(kvp.Key + howMany))
                        {
                            previousData[kvp.Key + howMany] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[kvp.Key + howMany] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<int, T>(ImmutableList.Create(selection), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<int, T> Insert<T>(this Dictionary<int, T> data, int where, int howMany, SelectionRange selection, Stack<ValueStackFrame<int, T>> undoStack) where T : struct
        {
            Dictionary<int, T> newData = new Dictionary<int, T>(data.Count);
            Dictionary<int, Reference<T>> previousData = new Dictionary<int, Reference<T>>(data.Count);
            Dictionary<int, Reference<T>> changedData = new Dictionary<int, Reference<T>>(data.Count);

            foreach (KeyValuePair<int, T> kvp in data)
            {
                if (kvp.Key < where)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else
                {
                    newData.Add(kvp.Key + howMany, kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey(kvp.Key + howMany))
                        {
                            previousData[kvp.Key + howMany] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[kvp.Key + howMany] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<int, T>(ImmutableList.Create(selection), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<(int, int), T> DeleteColumns<T>(this Dictionary<(int, int), T> data, SelectionRange columns, Stack<StackFrame<(int, int), T>> undoStack) where T : class
        {
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> previousData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> changedData = new Dictionary<(int, int), T>(data.Count);

            int delta = columns.Width;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 < columns.Left)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else if (kvp.Key.Item1 <= columns.Right)
                {
                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData.Add(kvp.Key, null);
                        }
                    }
                }
                else
                {
                    newData.Add((kvp.Key.Item1 - delta, kvp.Key.Item2), kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey((kvp.Key.Item1 - delta, kvp.Key.Item2)))
                        {
                            previousData[(kvp.Key.Item1 - delta, kvp.Key.Item2)] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[(kvp.Key.Item1 - delta, kvp.Key.Item2)] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<(int, int), T>(ImmutableList.Create(columns), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<(int, int), T> DeleteColumns<T>(this Dictionary<(int, int), T> data, SelectionRange columns, Stack<ValueStackFrame<(int, int), T>> undoStack) where T : struct
        {
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), Reference<T>> previousData = new Dictionary<(int, int), Reference<T>>(data.Count);
            Dictionary<(int, int), Reference<T>> changedData = new Dictionary<(int, int), Reference<T>>(data.Count);

            int delta = columns.Width;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item1 < columns.Left)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else if (kvp.Key.Item1 <= columns.Right)
                {
                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData.Add(kvp.Key, null);
                        }
                    }
                }
                else
                {
                    newData.Add((kvp.Key.Item1 - delta, kvp.Key.Item2), kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey((kvp.Key.Item1 - delta, kvp.Key.Item2)))
                        {
                            previousData[(kvp.Key.Item1 - delta, kvp.Key.Item2)] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[(kvp.Key.Item1 - delta, kvp.Key.Item2)] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<(int, int), T>(ImmutableList.Create(columns), previousData, changedData));
            }

            return newData;
        }


        public static Dictionary<(int, int), T> DeleteRows<T>(this Dictionary<(int, int), T> data, SelectionRange rows, Stack<StackFrame<(int, int), T>> undoStack) where T : class
        {
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> previousData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), T> changedData = new Dictionary<(int, int), T>(data.Count);

            int delta = rows.Height;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 < rows.Top)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else if (kvp.Key.Item2 <= rows.Bottom)
                {
                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData.Add(kvp.Key, null);
                        }
                    }
                }
                else
                {
                    newData.Add((kvp.Key.Item1, kvp.Key.Item2 - delta), kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 - delta)))
                        {
                            previousData[(kvp.Key.Item1, kvp.Key.Item2 - delta)] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[(kvp.Key.Item1, kvp.Key.Item2 - delta)] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<(int, int), T>(ImmutableList.Create(rows), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<(int, int), T> DeleteRows<T>(this Dictionary<(int, int), T> data, SelectionRange rows, Stack<ValueStackFrame<(int, int), T>> undoStack) where T : struct
        {
            Dictionary<(int, int), T> newData = new Dictionary<(int, int), T>(data.Count);
            Dictionary<(int, int), Reference<T>> previousData = new Dictionary<(int, int), Reference<T>>(data.Count);
            Dictionary<(int, int), Reference<T>> changedData = new Dictionary<(int, int), Reference<T>>(data.Count);

            int delta = rows.Height;

            foreach (KeyValuePair<(int, int), T> kvp in data)
            {
                if (kvp.Key.Item2 < rows.Top)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else if (kvp.Key.Item2 <= rows.Bottom)
                {
                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData.Add(kvp.Key, null);
                        }
                    }
                }
                else
                {
                    newData.Add((kvp.Key.Item1, kvp.Key.Item2 - delta), kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey((kvp.Key.Item1, kvp.Key.Item2 - delta)))
                        {
                            previousData[(kvp.Key.Item1, kvp.Key.Item2 - delta)] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[(kvp.Key.Item1, kvp.Key.Item2 - delta)] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<(int, int), T>(ImmutableList.Create(rows), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<int, T> Delete<T>(this Dictionary<int, T> data, int where, int howMany, SelectionRange selection, Stack<StackFrame<int, T>> undoStack) where T : class
        {
            Dictionary<int, T> newData = new Dictionary<int, T>(data.Count);
            Dictionary<int, T> previousData = new Dictionary<int, T>(data.Count);
            Dictionary<int, T> changedData = new Dictionary<int, T>(data.Count);

            foreach (KeyValuePair<int, T> kvp in data)
            {
                if (kvp.Key < where)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else if (kvp.Key < where + howMany)
                {
                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData.Add(kvp.Key, null);
                        }
                    }
                }
                else
                {
                    newData.Add(kvp.Key - howMany, kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey(kvp.Key - howMany))
                        {
                            previousData[kvp.Key - howMany] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[kvp.Key - howMany] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new StackFrame<int, T>(ImmutableList.Create(selection), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<int, T> Delete<T>(this Dictionary<int, T> data, int where, int howMany, SelectionRange selection, Stack<ValueStackFrame<int, T>> undoStack) where T : struct
        {
            Dictionary<int, T> newData = new Dictionary<int, T>(data.Count);
            Dictionary<int, Reference<T>> previousData = new Dictionary<int, Reference<T>>(data.Count);
            Dictionary<int, Reference<T>> changedData = new Dictionary<int, Reference<T>>(data.Count);

            foreach (KeyValuePair<int, T> kvp in data)
            {
                if (kvp.Key < where)
                {
                    newData.Add(kvp.Key, kvp.Value);
                }
                else if (kvp.Key < where + howMany)
                {
                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData.Add(kvp.Key, null);
                        }
                    }
                }
                else
                {
                    newData.Add(kvp.Key - howMany, kvp.Value);

                    if (undoStack != null)
                    {
                        previousData[kvp.Key] = kvp.Value;

                        if (!previousData.ContainsKey(kvp.Key - howMany))
                        {
                            previousData[kvp.Key - howMany] = null;
                        }

                        if (!changedData.ContainsKey(kvp.Key))
                        {
                            changedData[kvp.Key] = null;
                        }

                        changedData[kvp.Key - howMany] = kvp.Value;
                    }
                }
            }

            if (undoStack != null)
            {
                undoStack.Push(new ValueStackFrame<int, T>(ImmutableList.Create(selection), previousData, changedData));
            }

            return newData;
        }

        public static Dictionary<(int, int), T> MoveInsertDeleteRows<T>(this Dictionary<(int, int), T> data, int minRange, int maxRange, int deltaY, Stack<StackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : class
        {
            List<StackFrame<(int, int), T>> frames = new List<StackFrame<(int, int), T>>();

            data = data.InsertRows(new SelectionRange(0, minRange + deltaY, 0, maxRange + deltaY), undoStack);
            frames.Add(undoStack.Pop());

            if (deltaY > 0)
            {
                data.MoveRows(minRange, maxRange, deltaY, undoStack, leaveOriginal);
            }
            else
            {
                data.MoveRows(minRange + (maxRange - minRange + 1), maxRange + (maxRange - minRange + 1), deltaY - (maxRange - minRange + 1), undoStack, leaveOriginal);
            }

            frames.Add(undoStack.Pop());
            if (!leaveOriginal)
            {
                if (deltaY > 0)
                {
                    data = data.DeleteRows(new SelectionRange(0, minRange, 0, maxRange), undoStack);
                }
                else
                {
                    data = data.DeleteRows(new SelectionRange(0, minRange + (maxRange - minRange + 1), 0, maxRange + (maxRange - minRange + 1)), undoStack);
                }

                frames.Add(undoStack.Pop());
            }

            Dictionary<(int, int), T> allPreviousData = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> allNewData = new Dictionary<(int, int), T>();
            List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

            for (int i = frames.Count - 1; i >= 0; i--)
            {
                foreach (KeyValuePair<(int, int), T> kvp in frames[i].PreviousData)
                {
                    allPreviousData[kvp.Key] = kvp.Value;
                }

                undoStackRanges.AddRange(frames[i].Selection);
            }

            for (int i = 0; i < frames.Count; i++)
            {
                foreach (KeyValuePair<(int, int), T> kvp in frames[i].NewData)
                {
                    allNewData[kvp.Key] = kvp.Value;
                }
            }

            undoStack.Push(new StackFrame<(int, int), T>(undoStackRanges.RemoveDuplicates().Consolidate().ToImmutableList(), allPreviousData, allNewData));

            return data;
        }

        public static Dictionary<(int, int), T> MoveInsertDeleteRows<T>(this Dictionary<(int, int), T> data, int minRange, int maxRange, int deltaY, Stack<ValueStackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : struct
        {
            List<ValueStackFrame<(int, int), T>> frames = new List<ValueStackFrame<(int, int), T>>();

            data = data.InsertRows(new SelectionRange(0, minRange + deltaY, 0, maxRange + deltaY), undoStack);
            frames.Add(undoStack.Pop());
            if (deltaY > 0)
            {
                data.MoveRows(minRange, maxRange, deltaY, undoStack, leaveOriginal);
            }
            else
            {
                data.MoveRows(minRange + (maxRange - minRange + 1), maxRange + (maxRange - minRange + 1), deltaY - (maxRange - minRange + 1), undoStack, leaveOriginal);
            }

            frames.Add(undoStack.Pop());
            if (!leaveOriginal)
            {
                if (deltaY > 0)
                {
                    data = data.DeleteRows(new SelectionRange(0, minRange, 0, maxRange), undoStack);
                }
                else
                {
                    data = data.DeleteRows(new SelectionRange(0, minRange + (maxRange - minRange + 1), 0, maxRange + (maxRange - minRange + 1)), undoStack);
                }

                frames.Add(undoStack.Pop());
            }

            Dictionary<(int, int), Reference<T>> allPreviousData = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> allNewData = new Dictionary<(int, int), Reference<T>>();
            List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

            for (int i = frames.Count - 1; i >= 0; i--)
            {
                foreach (KeyValuePair<(int, int), Reference<T>> kvp in frames[i].PreviousData)
                {
                    allPreviousData[kvp.Key] = kvp.Value;
                }

                undoStackRanges.AddRange(frames[i].Selection);
            }

            for (int i = 0; i < frames.Count; i++)
            {
                foreach (KeyValuePair<(int, int), Reference<T>> kvp in frames[i].NewData)
                {
                    allNewData[kvp.Key] = kvp.Value;
                }
            }

            undoStack.Push(new ValueStackFrame<(int, int), T>(undoStackRanges.RemoveDuplicates().Consolidate().ToImmutableList(), allPreviousData, allNewData));

            return data;
        }

        public static Dictionary<(int, int), T> MoveInsertDeleteColumns<T>(this Dictionary<(int, int), T> data, int minRange, int maxRange, int deltaX, Stack<StackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : class
        {
            List<StackFrame<(int, int), T>> frames = new List<StackFrame<(int, int), T>>();

            data = data.InsertColumns(new SelectionRange(minRange + deltaX, 0, maxRange + deltaX, 0), undoStack);
            frames.Add(undoStack.Pop());
            if (deltaX > 0)
            {
                data.MoveColumns(minRange, maxRange, deltaX, undoStack, leaveOriginal);
            }
            else
            {
                data.MoveColumns(minRange + (maxRange - minRange + 1), maxRange + (maxRange - minRange + 1), deltaX - (maxRange - minRange + 1), undoStack, leaveOriginal);
            }
            frames.Add(undoStack.Pop());
            if (!leaveOriginal)
            {
                if (deltaX > 0)
                {
                    data = data.DeleteColumns(new SelectionRange(minRange, 0, maxRange, 0), undoStack);
                }
                else
                {
                    data = data.DeleteColumns(new SelectionRange(minRange + (maxRange - minRange + 1), 0, maxRange + (maxRange - minRange + 1), 0), undoStack);
                }

                frames.Add(undoStack.Pop());
            }

            Dictionary<(int, int), T> allPreviousData = new Dictionary<(int, int), T>();
            Dictionary<(int, int), T> allNewData = new Dictionary<(int, int), T>();
            List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

            for (int i = frames.Count - 1; i >= 0; i--)
            {
                foreach (KeyValuePair<(int, int), T> kvp in frames[i].PreviousData)
                {
                    allPreviousData[kvp.Key] = kvp.Value;
                }

                undoStackRanges.AddRange(frames[i].Selection);
            }

            for (int i = 0; i < frames.Count; i++)
            {
                foreach (KeyValuePair<(int, int), T> kvp in frames[i].NewData)
                {
                    allNewData[kvp.Key] = kvp.Value;
                }
            }

            undoStack.Push(new StackFrame<(int, int), T>(undoStackRanges.RemoveDuplicates().Consolidate().ToImmutableList(), allPreviousData, allNewData));

            return data;
        }

        public static Dictionary<(int, int), T> MoveInsertDeleteColumns<T>(this Dictionary<(int, int), T> data, int minRange, int maxRange, int deltaX, Stack<ValueStackFrame<(int, int), T>> undoStack, bool leaveOriginal) where T : struct
        {
            List<ValueStackFrame<(int, int), T>> frames = new List<ValueStackFrame<(int, int), T>>();

            data = data.InsertColumns(new SelectionRange(minRange + deltaX, 0, maxRange + deltaX, 0), undoStack);
            frames.Add(undoStack.Pop());
            if (deltaX > 0)
            {
                data.MoveColumns(minRange, maxRange, deltaX, undoStack, leaveOriginal);
            }
            else
            {
                data.MoveColumns(minRange + (maxRange - minRange + 1), maxRange + (maxRange - minRange + 1), deltaX - (maxRange - minRange + 1), undoStack, leaveOriginal);
            }
            frames.Add(undoStack.Pop());
            if (!leaveOriginal)
            {
                if (deltaX > 0)
                {
                    data = data.DeleteColumns(new SelectionRange(minRange, 0, maxRange, 0), undoStack);
                }
                else
                {
                    data = data.DeleteColumns(new SelectionRange(minRange + (maxRange - minRange + 1), 0, maxRange + (maxRange - minRange + 1), 0), undoStack);
                }

                frames.Add(undoStack.Pop());
            }

            Dictionary<(int, int), Reference<T>> allPreviousData = new Dictionary<(int, int), Reference<T>>();
            Dictionary<(int, int), Reference<T>> allNewData = new Dictionary<(int, int), Reference<T>>();
            List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

            for (int i = frames.Count - 1; i >= 0; i--)
            {
                foreach (KeyValuePair<(int, int), Reference<T>> kvp in frames[i].PreviousData)
                {
                    allPreviousData[kvp.Key] = kvp.Value;
                }

                undoStackRanges.AddRange(frames[i].Selection);
            }

            for (int i = 0; i < frames.Count; i++)
            {
                foreach (KeyValuePair<(int, int), Reference<T>> kvp in frames[i].NewData)
                {
                    allNewData[kvp.Key] = kvp.Value;
                }
            }

            undoStack.Push(new ValueStackFrame<(int, int), T>(undoStackRanges.RemoveDuplicates().Consolidate().ToImmutableList(), allPreviousData, allNewData));

            return data;
        }


        public static Dictionary<int, T> MoveInsertDelete<T>(this Dictionary<int, T> data, int minRange, int maxRange, int delta, SelectionRange selection, Stack<StackFrame<int, T>> undoStack, bool leaveOriginal) where T : class
        {
            List<StackFrame<int, T>> frames = new List<StackFrame<int, T>>();

            data = data.Insert(minRange + delta, maxRange - minRange + 1, selection, undoStack);
            frames.Add(undoStack.Pop());
            if (delta > 0)
            {
                data.Move(minRange, maxRange, delta, undoStack, leaveOriginal);
            }
            else
            {
                data.Move(minRange + (maxRange - minRange + 1), maxRange + (maxRange - minRange + 1), delta - (maxRange - minRange + 1), undoStack, leaveOriginal);
            }

            frames.Add(undoStack.Pop());
            if (!leaveOriginal)
            {
                if (delta > 0)
                {
                    data = data.Delete(minRange, maxRange - minRange + 1, selection, undoStack);
                }
                else
                {
                    data = data.Delete(minRange + (maxRange - minRange + 1), maxRange - minRange + 1, selection, undoStack);
                }

                frames.Add(undoStack.Pop());
            }

            Dictionary<int, T> allPreviousData = new Dictionary<int, T>();
            Dictionary<int, T> allNewData = new Dictionary<int, T>();
            List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

            for (int i = frames.Count - 1; i >= 0; i--)
            {
                foreach (KeyValuePair<int, T> kvp in frames[i].PreviousData)
                {
                    allPreviousData[kvp.Key] = kvp.Value;
                }

                undoStackRanges.AddRange(frames[i].Selection);
            }

            for (int i = 0; i < frames.Count; i++)
            {
                foreach (KeyValuePair<int, T> kvp in frames[i].NewData)
                {
                    allNewData[kvp.Key] = kvp.Value;
                }
            }

            undoStack.Push(new StackFrame<int, T>(undoStackRanges.RemoveDuplicates().Consolidate().ToImmutableList(), allPreviousData, allNewData));

            return data;
        }

        public static Dictionary<int, T> MoveInsertDelete<T>(this Dictionary<int, T> data, int minRange, int maxRange, int delta, SelectionRange selection, Stack<ValueStackFrame<int, T>> undoStack, bool leaveOriginal) where T : struct
        {
            List<ValueStackFrame<int, T>> frames = new List<ValueStackFrame<int, T>>();

            data = data.Insert(minRange + delta, maxRange - minRange + 1, selection, undoStack);
            frames.Add(undoStack.Pop());
            if (delta > 0)
            {
                data.Move(minRange, maxRange, delta, undoStack, leaveOriginal);
            }
            else
            {
                data.Move(minRange + (maxRange - minRange + 1), maxRange + (maxRange - minRange + 1), delta - (maxRange - minRange + 1), undoStack, leaveOriginal);
            }

            frames.Add(undoStack.Pop());
            if (!leaveOriginal)
            {
                if (delta > 0)
                {
                    data = data.Delete(minRange, maxRange - minRange + 1, selection, undoStack);
                }
                else
                {
                    data = data.Delete(minRange + (maxRange - minRange + 1), maxRange - minRange + 1, selection, undoStack);
                }

                frames.Add(undoStack.Pop());
            }

            Dictionary<int, Reference<T>> allPreviousData = new Dictionary<int, Reference<T>>();
            Dictionary<int, Reference<T>> allNewData = new Dictionary<int, Reference<T>>();
            List<SelectionRange> undoStackRanges = new List<SelectionRange>(frames.Count);

            for (int i = frames.Count - 1; i >= 0; i--)
            {
                foreach (KeyValuePair<int, Reference<T>> kvp in frames[i].PreviousData)
                {
                    allPreviousData[kvp.Key] = kvp.Value;
                }

                undoStackRanges.AddRange(frames[i].Selection);
            }

            for (int i = 0; i < frames.Count; i++)
            {
                foreach (KeyValuePair<int, Reference<T>> kvp in frames[i].NewData)
                {
                    allNewData[kvp.Key] = kvp.Value;
                }
            }

            undoStack.Push(new ValueStackFrame<int, T>(undoStackRanges.RemoveDuplicates().Consolidate().ToImmutableList(), allPreviousData, allNewData));

            return data;
        }

        public static void FillRight(this Dictionary<(int, int), string> data, SelectionRange range, int howMany, Stack<StackFrame<(int, int), string>> undoStack)
        {
            Dictionary<(int, int), string> previousData = new Dictionary<(int, int), string>();
            Dictionary<(int, int), string> newData = new Dictionary<(int, int), string>();

            if (undoStack != null)
            {
                for (int i = range.Top; i <= range.Bottom; i++)
                {
                    for (int j = 0; j < howMany; j++)
                    {
                        if (data.TryGetValue((range.Right + j + 1, i), out string val))
                        {
                            previousData[(range.Right + j + 1, i)] = val;
                        }
                        else
                        {
                            previousData[(range.Right + j + 1, i)] = null;
                        }
                    }
                }
            }

            for (int i = range.Top; i <= range.Bottom; i++)
            {
                data.FillRight(range.Left, range.Right, i, howMany);
            }

            if (undoStack != null)
            {
                for (int i = range.Top; i <= range.Bottom; i++)
                {
                    for (int j = 0; j < howMany; j++)
                    {
                        if (data.TryGetValue((range.Right + j + 1, i), out string val))
                        {
                            newData[(range.Right + j + 1, i)] = val;
                        }
                        else
                        {
                            newData[(range.Right + j + 1, i)] = null;
                        }
                    }
                }

                undoStack.Push(new StackFrame<(int, int), string>(ImmutableList.Create(range), previousData, newData));
            }
        }

        public static void FillBottom(this Dictionary<(int, int), string> data, SelectionRange range, int howMany, Stack<StackFrame<(int, int), string>> undoStack)
        {
            Dictionary<(int, int), string> previousData = new Dictionary<(int, int), string>();
            Dictionary<(int, int), string> newData = new Dictionary<(int, int), string>();

            if (undoStack != null)
            {
                for (int i = range.Left; i <= range.Right; i++)
                {
                    for (int j = 0; j < howMany; j++)
                    {
                        if (data.TryGetValue((i, range.Bottom + j + 1), out string val))
                        {
                            previousData[(i, range.Bottom + j + 1)] = val;
                        }
                        else
                        {
                            previousData[(i, range.Bottom + j + 1)] = null;
                        }
                    }
                }
            }

            for (int i = range.Left; i <= range.Right; i++)
            {
                data.FillBottom(range.Top, range.Bottom, i, howMany);
            }

            if (undoStack != null)
            {
                for (int i = range.Left; i <= range.Right; i++)
                {
                    for (int j = 0; j < howMany; j++)
                    {
                        if (data.TryGetValue((i, range.Bottom + j + 1), out string val))
                        {
                            newData[(i, range.Bottom + j + 1)] = val;
                        }
                        else
                        {
                            newData[(i, range.Bottom + j + 1)] = null;
                        }
                    }
                }

                undoStack.Push(new StackFrame<(int, int), string>(ImmutableList.Create(range), previousData, newData));
            }
        }

        public static void FillLeft(this Dictionary<(int, int), string> data, SelectionRange range, int howMany, Stack<StackFrame<(int, int), string>> undoStack)
        {
            Dictionary<(int, int), string> previousData = new Dictionary<(int, int), string>();
            Dictionary<(int, int), string> newData = new Dictionary<(int, int), string>();

            if (undoStack != null)
            {
                for (int i = range.Top; i <= range.Bottom; i++)
                {
                    for (int j = 0; j < howMany; j++)
                    {
                        if (data.TryGetValue((range.Left - j - 1, i), out string val))
                        {
                            previousData[(range.Left - j - 1, i)] = val;
                        }
                        else
                        {
                            previousData[(range.Left - j - 1, i)] = null;
                        }
                    }
                }
            }

            for (int i = range.Top; i <= range.Bottom; i++)
            {
                data.FillLeft(range.Left, range.Right, i, howMany);
            }

            if (undoStack != null)
            {
                for (int i = range.Top; i <= range.Bottom; i++)
                {
                    for (int j = 0; j < howMany; j++)
                    {
                        if (data.TryGetValue((range.Left - j - 1, i), out string val))
                        {
                            newData[(range.Left - j - 1, i)] = val;
                        }
                        else
                        {
                            newData[(range.Left - j - 1, i)] = null;
                        }
                    }
                }

                undoStack.Push(new StackFrame<(int, int), string>(ImmutableList.Create(range), previousData, newData));
            }
        }

        public static void FillTop(this Dictionary<(int, int), string> data, SelectionRange range, int howMany, Stack<StackFrame<(int, int), string>> undoStack)
        {
            Dictionary<(int, int), string> previousData = new Dictionary<(int, int), string>();
            Dictionary<(int, int), string> newData = new Dictionary<(int, int), string>();

            if (undoStack != null)
            {
                for (int i = range.Left; i <= range.Right; i++)
                {
                    for (int j = 0; j < howMany; j++)
                    {
                        if (data.TryGetValue((i, range.Top - j - 1), out string val))
                        {
                            previousData[(i, range.Top - j - 1)] = val;
                        }
                        else
                        {
                            previousData[(i, range.Top - j - 1)] = null;
                        }
                    }
                }
            }

            for (int i = range.Left; i <= range.Right; i++)
            {
                data.FillTop(range.Top, range.Bottom, i, howMany);
            }

            if (undoStack != null)
            {
                for (int i = range.Left; i <= range.Right; i++)
                {
                    for (int j = 0; j < howMany; j++)
                    {
                        if (data.TryGetValue((i, range.Top - j - 1), out string val))
                        {
                            newData[(i, range.Top - j - 1)] = val;
                        }
                        else
                        {
                            newData[(i, range.Top - j - 1)] = null;
                        }
                    }
                }

                undoStack.Push(new StackFrame<(int, int), string>(ImmutableList.Create(range), previousData, newData));
            }
        }

        private static void FillRight(this Dictionary<(int, int), string> data, int left, int right, int y, int howMany)
        {
            string[] values = new string[right - left + 1];

            for (int i = left; i <= right; i++)
            {
                if (!data.TryGetValue((i, y), out values[i - left]))
                {
                    values[i - left] = null;
                }
            }

            Func<int, string> fillValue = GetPattern(values);

            for (int i = 0; i < howMany; i++)
            {
                string val = fillValue(i + 1);

                if (!string.IsNullOrEmpty(val))
                {
                    data[(right + i + 1, y)] = val;
                }
                else
                {
                    data.Remove((right + i + 1, y));
                }
            }
        }

        internal static string GetLastFillRight(this Dictionary<(int, int), string> data, int left, int right, int y, int howMany)
        {
            string[] values = new string[right - left + 1];

            for (int i = left; i <= right; i++)
            {
                if (!data.TryGetValue((i, y), out values[i - left]))
                {
                    values[i - left] = null;
                }
            }

            Func<int, string> fillValue = GetPattern(values);

            for (int i = 0; i < howMany - 1; i++)
            {
                fillValue(i + 1);
            }

            return fillValue(howMany);
        }

        private static void FillBottom(this Dictionary<(int, int), string> data, int top, int bottom, int x, int howMany)
        {
            string[] values = new string[bottom - top + 1];

            for (int i = top; i <= bottom; i++)
            {
                if (!data.TryGetValue((x, i), out values[i - top]))
                {
                    values[i - top] = null;
                }
            }

            Func<int, string> fillValue = GetPattern(values);

            for (int i = 0; i < howMany; i++)
            {
                string val = fillValue(i + 1);

                if (!string.IsNullOrEmpty(val))
                {
                    data[(x, bottom + i + 1)] = val;
                }
                else
                {
                    data.Remove((x, bottom + i + 1));
                }
            }
        }

        internal static string GetLastFillBottom(this Dictionary<(int, int), string> data, int top, int bottom, int x, int howMany)
        {
            string[] values = new string[bottom - top + 1];

            for (int i = top; i <= bottom; i++)
            {
                if (!data.TryGetValue((x, i), out values[i - top]))
                {
                    values[i - top] = null;
                }
            }

            Func<int, string> fillValue = GetPattern(values);
            
            for (int i = 0; i < howMany - 1; i++)
            {
                fillValue(i + 1);
            }

            return fillValue(howMany);
        }

        private static void FillLeft(this Dictionary<(int, int), string> data, int left, int right, int y, int howMany)
        {
            string[] values = new string[right - left + 1];

            for (int i = left; i <= right; i++)
            {
                if (!data.TryGetValue((i, y), out values[right - i]))
                {
                    values[right - i] = null;
                }
            }

            Func<int, string> fillValue = GetPattern(values);

            for (int i = 0; i < howMany; i++)
            {
                string val = fillValue(i + 1);

                if (!string.IsNullOrEmpty(val))
                {
                    data[(left - i - 1, y)] = val;
                }
                else
                {
                    data.Remove((left - i - 1, y));
                }
            }
        }

        internal static string GetLastFillLeft(this Dictionary<(int, int), string> data, int left, int right, int y, int howMany)
        {
            string[] values = new string[right - left + 1];

            for (int i = left; i <= right; i++)
            {
                if (!data.TryGetValue((i, y), out values[right - i]))
                {
                    values[right - i] = null;
                }
            }

            Func<int, string> fillValue = GetPattern(values);

            for (int i = 0; i < howMany - 1; i++)
            {
                fillValue(i + 1);
            }

            return fillValue(howMany);
        }

        private static void FillTop(this Dictionary<(int, int), string> data, int top, int bottom, int x, int howMany)
        {
            string[] values = new string[bottom - top + 1];

            for (int i = top; i <= bottom; i++)
            {
                if (!data.TryGetValue((x, i), out values[bottom - i]))
                {
                    values[bottom - i] = null;
                }
            }

            Func<int, string> fillValue = GetPattern(values);

            for (int i = 0; i < howMany; i++)
            {
                string val = fillValue(i + 1);

                if (!string.IsNullOrEmpty(val))
                {
                    data[(x, top - i - 1)] = val;
                }
                else
                {
                    data.Remove((x, top - i - 1));
                }
            }
        }

        internal static string GetLastFillTop(this Dictionary<(int, int), string> data, int top, int bottom, int x, int howMany)
        {
            string[] values = new string[bottom - top + 1];

            for (int i = top; i <= bottom; i++)
            {
                if (!data.TryGetValue((x, i), out values[bottom - i]))
                {
                    values[bottom - i] = null;
                }
            }

            Func<int, string> fillValue = GetPattern(values);

            for (int i = 0; i < howMany - 1; i++)
            {
                fillValue(i + 1);
            }

            return fillValue(howMany);
        }


        private static Func<int, string> GetPattern(string[] values)
        {
            Regex space = new Regex("\\b[^$]", RegexOptions.Compiled | RegexOptions.RightToLeft);

            (string prefix, string suffix)[] elements = new (string, string)[values.Length];

            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] != null)
                {
                    Match match = space.Match(values[i]);

                    if (match.Success)
                    {
                        elements[i] = (values[i].Substring(0, match.Index), values[i].Substring(match.Index));
                    }
                    else
                    {
                        elements[i] = ("", values[i]);
                    }
                }
                else
                {
                    elements[i] = ("@@@EMPTYPREFIX@@@", "");
                }
            }

            Dictionary<string, Func<int, string>> increments = new Dictionary<string, Func<int, string>>();
            Dictionary<string, int> lastIncrements = new Dictionary<string, int>();
            Dictionary<string, List<string>> valuesByPrefix = new Dictionary<string, List<string>>();

            for (int i = 0; i < elements.Length; i++)
            {
                if (!valuesByPrefix.TryGetValue(elements[i].prefix, out List<string> vals))
                {
                    vals = new List<string>();
                    valuesByPrefix[elements[i].prefix] = vals;
                }

                if (!string.IsNullOrEmpty(elements[i].suffix))
                {
                    vals.Add(elements[i].suffix);
                }
            }

            foreach (KeyValuePair<string, List<string>> kvp in valuesByPrefix)
            {
                increments[kvp.Key] = GetIncrement(kvp.Value);
                lastIncrements[kvp.Key] = 0;
            }

            return x =>
            {
                string prefix = elements[(x - 1) % elements.Length].prefix;

                if (prefix == "@@@EMPTYPREFIX@@@")
                {
                    return null;
                }

                int lastIncrement = lastIncrements[prefix];
                lastIncrements[prefix] = lastIncrement + 1;

                return prefix + increments[prefix](lastIncrement + 1);
            };
        }

        private static (decimal mean, decimal variance) MeanAndVariance(this IEnumerable<decimal> values)
        {
            decimal mean = 0;
            decimal meanSq = 0;
            int count = 0;

            foreach (decimal val in values)
            {
                mean += val;
                meanSq += val * val;
                count++;
            }

            if (count == 0)
            {
                return (0, 0);
            }
            else
            {
                mean /= count;
                meanSq /= count;

                return (mean, meanSq - mean * mean);
            }
        }

        // From https://stackoverflow.com/questions/29596449/is-it-possible-to-raise-a-decimal-to-an-integer-power-in-net
        private static decimal Pow(decimal x, int n)
        {
            if (n < 0)
            {
                throw new ArgumentOutOfRangeException("n");
            }

            decimal result = 1;
            decimal multiplier = x;

            while (n > 0)
            {
                if ((n & 1) > 0) result *= multiplier;
                multiplier *= multiplier;
                n >>= 1;
            }

            return result;

        }

        private static Func<int, string> GetIncrement(List<string> values)
        {
            if (values.Count == 0)
            {
                return x => "";
            }

            decimal[] vals = new decimal[values.Count];
            if (values.Select((x, i) => decimal.TryParse(x, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out vals[i])).All(x => x))
            {
                if (vals.Length == 1)
                {
                    return x => (vals[0] + x).ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                decimal[] differences = new decimal[values.Count - 1];
                decimal[] ratios = new decimal[values.Count - 1];

                bool anyZeros = false;

                for (int i = 0; i < vals.Length; i++)
                {
                    if (vals[i] == 0)
                    {
                        anyZeros = true;
                    }

                    if (i > 0)
                    {
                        differences[i - 1] = vals[i] - vals[i - 1];

                        if (!anyZeros)
                        {
                            ratios[i - 1] = vals[i] / vals[i - 1];
                        }
                    }
                }

                (decimal mean, decimal variance) differenceStats = differences.MeanAndVariance();
                (decimal mean, decimal variance) ratioStats = ratios.MeanAndVariance();

                if (anyZeros || differenceStats.variance <= ratioStats.variance)
                {
                    if (differenceStats.variance < Math.Max(Math.Abs(differenceStats.mean), 1e-4m) * 1e-4m)
                    {
                        return x => (vals[vals.Length - 1] + x * differenceStats.mean).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        return x => values[(x - 1) % values.Count];
                    }
                }
                else
                {
                    if (ratioStats.variance < Math.Max(Math.Abs(ratioStats.mean), 1e-4m) * 1e-4m)
                    {
                        return x => (vals[vals.Length - 1] * Pow(ratioStats.mean, x)).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        return x => values[(x - 1) % values.Count];
                    }
                }
            }
            else if (values.All(x => x.Length == 1))
            {
                int[] charVals = values.Select(x => (int)x[0]).ToArray();

                if (charVals.All(x => x >= 65 && x <= 90) || charVals.All(x => x >= 97 && x <= 122))
                {
                    if (charVals.Length == 1)
                    {
                        if (charVals[0] >= 65 && charVals[0] <= 90)
                        {
                            return x => new string((char)(65 + (charVals[0] + x - 65) % 26), 1);
                        }
                        else
                        {
                            return x => new string((char)(97 + (charVals[0] + x - 97) % 26), 1);
                        }
                    }

                    HashSet<int> differences = new HashSet<int>();

                    for (int i = 1; i < charVals.Length; i++)
                    {
                        int diff = charVals[i] - charVals[i - 1];

                        if (diff < 0)
                        {
                            diff = 26 + diff;
                        }

                        differences.Add(diff);
                    }

                    if (differences.Count == 1)
                    {
                        int diff = differences.ElementAt(0);

                        if (charVals[0] >= 65 && charVals[0] <= 90)
                        {
                            return x => new string((char)(65 + (charVals[charVals.Length - 1] + x * diff - 65) % 26), 1);
                        }
                        else
                        {
                            return x => new string((char)(97 + (charVals[charVals.Length - 1] + x * diff - 97) % 26), 1);
                        }
                    }
                    else
                    {
                        return x => values[(x - 1) % values.Count];
                    }
                }
                else
                {
                    return x => values[(x - 1) % values.Count];
                }
            }
            else
            {
                return x => values[(x - 1) % values.Count];
            }
        }
    }
}
