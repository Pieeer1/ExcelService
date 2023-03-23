namespace ExcelService.Models.Numerics
{
    public struct UIntVector2
    {
        public uint X { get; set; }
        public uint Y { get; set; }
        public UIntVector2(uint x, uint y)
        {
            X = x;
            Y = y;
        }
        public static UIntVector2 Zero => new UIntVector2(0, 0);
    }
}
