using System;
using System.Drawing;

namespace PowerPointAddIn1
{
    [Serializable]
    public class NavBarSettings
    {
        public enum ShapeType
        {
            Circle,
            Square,
            NumberOnly
        }

        public Color BackgroundColor { get; set; } = Color.Black;
        public Color SectionNameColor { get; set; } = Color.White;
        public Color CurrentSectionNameColor { get; set; } = Color.White;
        public Color OtherSectionNameColor { get; set; } = Color.Gainsboro;
        public Color CurrentSlideColor { get; set; } = Color.White;
        public Color SameSubsectionBorderColor { get; set; } = Color.Red;
        public Color SameSubsectionFillColor { get; set; } = Color.Black;
        public Color OtherSlidesBorderColor { get; set; } = Color.White;
        public float SubsectionBoxTransparency { get; set; } = 0.3f;
        
        // Shape customization
        public ShapeType SlideShapeType { get; set; } = ShapeType.Circle;
        public bool ShowSlideNumbers { get; set; } = true;
        public Color SlideNumberColor { get; set; } = Color.White;

        public Color[] SubsectionBoxColors { get; set; } = new Color[]
        {
            Color.SteelBlue,
            Color.MediumSeaGreen,
            Color.Goldenrod,
            Color.IndianRed,
            Color.BlueViolet,
            Color.DarkOrange,
            Color.DarkSlateBlue,
            Color.RosyBrown
        };

        public void ResetToDefaults()
        {
            BackgroundColor = Color.Black;
            SectionNameColor = Color.White;
            CurrentSectionNameColor = Color.White;
            OtherSectionNameColor = Color.Gainsboro;
            CurrentSlideColor = Color.White;
            SameSubsectionBorderColor = Color.Red;
            SameSubsectionFillColor = Color.Black;
            OtherSlidesBorderColor = Color.White;
            SubsectionBoxTransparency = 0.3f;
            SlideShapeType = ShapeType.Circle;
            ShowSlideNumbers = true;
            SlideNumberColor = Color.White;
            SubsectionBoxColors = new Color[]
            {
                Color.SteelBlue,
                Color.MediumSeaGreen,
                Color.Goldenrod,
                Color.IndianRed,
                Color.BlueViolet,
                Color.DarkOrange,
                Color.DarkSlateBlue,
                Color.RosyBrown
            };
        }
    }
}