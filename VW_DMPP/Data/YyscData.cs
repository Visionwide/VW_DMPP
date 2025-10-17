namespace VW_DMPP.Data
{
    public class YyscData
    {
        public DateTime 出機日期 { get; set; }
        public string 機號 { get; set; } = string.Empty;
        public string 機型 { get; set; } = string.Empty;
        public decimal? 光學尺_X軸 { get; set; }
        public decimal? 光學尺_Y軸 { get; set; }
        public decimal? 光學尺_Z軸 { get; set; }
        public decimal? 雷射_X軸_P { get; set; }
        public decimal? 雷射_X軸_PS { get; set; }
        public decimal? 雷射_Y軸_P { get; set; }
        public decimal? 雷射_Y軸_PS { get; set; }
        public decimal? 雷射_Z軸_P { get; set; }
        public decimal? 雷射_Z軸_PS { get; set; }
        public decimal? 循圓_XY軸 { get; set; }
        public decimal? 幾何_工作台面平行度_X軸 { get; set; }
        public decimal? 幾何_工作台面平行度_Y軸 { get; set; }
        public decimal? 幾何_T型槽平行度_X軸方向 { get; set; }
        public decimal? 幾何_直角度_XY軸 { get; set; }
        public decimal? 幾何_直角度_ZX軸 { get; set; }
        public decimal? 幾何_直角度_ZY軸 { get; set; }
        public decimal? 幾何_主軸直角度_X軸方向 { get; set; }
        public decimal? 幾何_主軸直角度_Y軸方向 { get; set; }
        public decimal? 主軸_外圓偏擺 { get; set; }
        public decimal? 主軸_端面偏擺_上端 { get; set; }
        public decimal? 主軸_端面偏擺_下端 { get; set; }
        public decimal? 主軸_錐孔偏擺_300mm { get; set; }
    }
}