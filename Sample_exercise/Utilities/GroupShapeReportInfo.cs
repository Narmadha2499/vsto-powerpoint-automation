using System;
using System.Collections.Generic;
using System.Text;

namespace Sample_exercise.Utilities
{
    public class GroupShapeReportInfo
    {
        public int ShapeId { get; set; }

        public string ShapeName { get; set; }

        public string ShapeType { get; set; }

        public float LeftPosition { get; set; }

        public float TopPosition { get; set; }

        public float Width { get; set; }

        public float Height { get; set; }

        public float Rotation { get; set; }

        public int DepthLevel { get; set; }

        public bool IsGroupShape { get; set; }
    }
}
