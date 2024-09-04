using SixLabors.ImageSharp;

namespace NPOI.SS.Converter
{
    public class Picture
    {
        public byte[] Data;
        public Size Size;
        public Point Offset;
        public CellCoordinates CellCoordinates;
        public string Extension;
    }

    public struct CellCoordinates
    {
        public int Column;
        public int Row;

        public CellCoordinates(int column, int row)
        {
            Column = column;
            Row = row;
        }
    }
}