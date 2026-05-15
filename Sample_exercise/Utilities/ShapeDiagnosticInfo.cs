using System;

namespace Sample_exercise.Utilities
{
    public class ShapeDiagnosticInfo
    {
        public int SlideNumber { get; set; }

        public string ShapeName { get; set; }

        public string ShapeType { get; set; }

        public float Left { get; set; }

        public float Top { get; set; }

        public float Width { get; set; }

        public float Height { get; set; }

        public bool HasTextFrame { get; set; }

        public bool IsGroupShape { get; set; }

        public string PlaceholderType { get; set; }
    }
}