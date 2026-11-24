using System;
using System.Collections.Generic;
using System.Drawing;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace InteractiveDisplay
{
    internal sealed class Stroke
    {
        public readonly List<PointF> Points = new List<PointF>();
        public readonly Color Color;
        public readonly float Width;
        public RectangleF Bounds { get; private set; }

        public Stroke(Color color, float width)
        {
            Color = color;
            Width = width;
        }

        public void Append(PointF point)
        {
            float radius = Width / 2;
            var footprint = new RectangleF(point.X - radius, point.Y - radius, Width, Width);
            Bounds = Points.Count == 0 ? footprint : RectangleF.Union(Bounds, footprint);
            Points.Add(point);
        }

        public bool IsNear(PointF point, float radius)
        {
            float limit = radius + Width / 2;
            PointF previous = Points[0];
            foreach (PointF next in Points)
            {
                float dx = next.X - previous.X;
                float dy = next.Y - previous.Y;
                float length = dx * dx + dy * dy;
                float t = length == 0 ? 0 : ((point.X - previous.X) * dx + (point.Y - previous.Y) * dy) / length;
                t = Math.Max(0, Math.Min(1, t));
                float x = point.X - (previous.X + t * dx);
                float y = point.Y - (previous.Y + t * dy);
                if (x * x + y * y <= limit * limit)
                    return true;
                previous = next;
            }
            return false;
        }
    }

    internal sealed class InkPage
    {
        private readonly Stack<List<Stroke>> undo = new Stack<List<Stroke>>();
        public List<Stroke> Marks { get; private set; }

        public InkPage()
        {
            Marks = new List<Stroke>();
        }

        public void Remember()
        {
            undo.Push(new List<Stroke>(Marks));
        }

        public void Undo()
        {
            if (undo.Count > 0)
                Marks = undo.Pop();
        }

        public void Clear()
        {
            if (Marks.Count == 0)
                return;
            Remember();
            Marks.Clear();
        }
    }

    internal sealed class SlideInk
    {
        public readonly Dictionary<int, InkPage> Pages = new Dictionary<int, InkPage>();

        public bool HasMarks
        {
            get
            {
                foreach (InkPage page in Pages.Values)
                {
                    if (page.Marks.Count > 0)
                        return true;
                }
                return false;
            }
        }

        public InkPage Page(int slideId)
        {
            InkPage page;
            if (!Pages.TryGetValue(slideId, out page))
            {
                page = new InkPage();
                Pages.Add(slideId, page);
            }
            return page;
        }
    }

    internal static class PowerPointInkWriter
    {
        public static void AddToPresentation(
            PowerPoint.Application application, PowerPoint.Presentation presentation, SlideInk ink)
        {
            var targets = new Dictionary<int, PowerPoint.Slide>();
            foreach (var entry in ink.Pages)
            {
                if (entry.Value.Marks.Count > 0)
                    targets.Add(entry.Key, presentation.Slides.FindBySlideID(entry.Key));
            }

            var added = new List<PowerPoint.Shape>();
            application.StartNewUndoEntry();
            try
            {
                foreach (var target in targets)
                {
                    foreach (Stroke mark in ink.Pages[target.Key].Marks)
                        AddStroke(target.Value.Shapes, mark, added);
                }
            }
            catch
            {
                for (int index = added.Count - 1; index >= 0; index--)
                    added[index].Delete();
                throw;
            }
        }

        private static void AddStroke(PowerPoint.Shapes shapes, Stroke stroke, List<PowerPoint.Shape> added)
        {
            PointF first = stroke.Points[0];
            PowerPoint.Shape shape;
            if (stroke.Points.Count == 1)
            {
                float radius = stroke.Width / 2;
                shape = shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval,
                    first.X - radius, first.Y - radius, stroke.Width, stroke.Width);
                added.Add(shape);
                shape.Line.Visible = Office.MsoTriState.msoFalse;
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(stroke.Color);
                shape.Fill.Transparency = 1 - stroke.Color.A / 255f;
            }
            else
            {
                PowerPoint.FreeformBuilder path = shapes.BuildFreeform(Office.MsoEditingType.msoEditingCorner, first.X, first.Y);
                for (int index = 1; index < stroke.Points.Count; index++)
                {
                    PointF point = stroke.Points[index];
                    path.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, point.X, point.Y, 0, 0, 0, 0);
                }

                shape = path.ConvertToShape();
                added.Add(shape);
                shape.Fill.Visible = Office.MsoTriState.msoFalse;
                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.Weight = stroke.Width;
                shape.Line.ForeColor.RGB = ColorTranslator.ToOle(stroke.Color);
                shape.Line.Transparency = 1 - stroke.Color.A / 255f;
            }
            shape.Tags.Add("INTERACTIVE_DISPLAY", "INK");
        }
    }
}
