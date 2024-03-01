namespace Excel.Extension
{
    internal class AnchorModel
    {
        public int CellX { get; set; } = 0;
        public int CellY { get; set; } = 0;
    }

    internal class DataRangeModel
    {
        public int RangeX { get; set; } = 0;
        public int RangeY { get; set; } = 0;
    }

    internal class SheetRangeModel
    {
        public int StartIndex { get; set; } = 0;
        public int EndIndex { get; set; } = 0;
    }
}
