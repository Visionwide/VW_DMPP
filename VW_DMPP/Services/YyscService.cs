using OfficeOpenXml;
using VW_DMPP.Data;

namespace VW_DMPP.Services
{
    public class YyscService
    {
        private readonly IWebHostEnvironment _env;
        private string ExcelPath => Path.Combine(_env.WebRootPath, "data", "yysc.xlsx");

        // ⭐ 靜態建構子：在類別載入時執行一次
        static YyscService()
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        public YyscService(IWebHostEnvironment env)
        {
            _env = env;
        }

        public List<YyscData> ReadData()
        {
            var data = new List<YyscData>();

            if (!File.Exists(ExcelPath))
            {
                throw new FileNotFoundException($"找不到檔案: {ExcelPath}");
            }

            using (var package = new ExcelPackage(new FileInfo(ExcelPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension?.Rows ?? 0;

                // 從第 2 列開始讀取（第 1 列是標題）
                for (int row = 2; row <= rowCount; row++)
                {
                    var item = new YyscData
                    {
                        出機日期 = GetCellValue<DateTime>(worksheet, row, 1),
                        機號 = GetCellValue<string>(worksheet, row, 2) ?? string.Empty,
                        機型 = GetCellValue<string>(worksheet, row, 3) ?? string.Empty,
                        光學尺_X軸 = GetCellValue<decimal?>(worksheet, row, 4),
                        光學尺_Y軸 = GetCellValue<decimal?>(worksheet, row, 5),
                        光學尺_Z軸 = GetCellValue<decimal?>(worksheet, row, 6),
                        雷射_X軸_P = GetCellValue<decimal?>(worksheet, row, 7),
                        雷射_X軸_PS = GetCellValue<decimal?>(worksheet, row, 8),
                        雷射_Y軸_P = GetCellValue<decimal?>(worksheet, row, 9),
                        雷射_Y軸_PS = GetCellValue<decimal?>(worksheet, row, 10),
                        雷射_Z軸_P = GetCellValue<decimal?>(worksheet, row, 11),
                        雷射_Z軸_PS = GetCellValue<decimal?>(worksheet, row, 12),
                        循圓_XY軸 = GetCellValue<decimal?>(worksheet, row, 13),
                        幾何_工作台面平行度_X軸 = GetCellValue<decimal?>(worksheet, row, 14),
                        幾何_工作台面平行度_Y軸 = GetCellValue<decimal?>(worksheet, row, 15),
                        幾何_T型槽平行度_X軸方向 = GetCellValue<decimal?>(worksheet, row, 16),
                        幾何_直角度_XY軸 = GetCellValue<decimal?>(worksheet, row, 17),
                        幾何_直角度_ZX軸 = GetCellValue<decimal?>(worksheet, row, 18),
                        幾何_直角度_ZY軸 = GetCellValue<decimal?>(worksheet, row, 19),
                        幾何_主軸直角度_X軸方向 = GetCellValue<decimal?>(worksheet, row, 20),
                        幾何_主軸直角度_Y軸方向 = GetCellValue<decimal?>(worksheet, row, 21),
                        主軸_外圓偏擺 = GetCellValue<decimal?>(worksheet, row, 22),
                        主軸_端面偏擺_上端 = GetCellValue<decimal?>(worksheet, row, 23),
                        主軸_端面偏擺_下端 = GetCellValue<decimal?>(worksheet, row, 24),
                        主軸_錐孔偏擺_300mm = GetCellValue<decimal?>(worksheet, row, 25)
                    };

                    data.Add(item);
                }
            }

            return data;
        }

        public void SaveData(List<YyscData> data)
        {
            if (!File.Exists(ExcelPath))
            {
                // 如果檔案不存在，建立新檔案
                CreateNewExcelFile();
            }

            using (var package = new ExcelPackage(new FileInfo(ExcelPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // 清除舊資料（保留標題列）
                if (worksheet.Dimension != null)
                {
                    worksheet.DeleteRow(2, worksheet.Dimension.Rows - 1);
                }

                // 寫入新資料
                for (int i = 0; i < data.Count; i++)
                {
                    int row = i + 2;
                    var item = data[i];

                    worksheet.Cells[row, 1].Value = item.出機日期;
                    worksheet.Cells[row, 2].Value = item.機號;
                    worksheet.Cells[row, 3].Value = item.機型;
                    worksheet.Cells[row, 4].Value = item.光學尺_X軸;
                    worksheet.Cells[row, 5].Value = item.光學尺_Y軸;
                    worksheet.Cells[row, 6].Value = item.光學尺_Z軸;
                    worksheet.Cells[row, 7].Value = item.雷射_X軸_P;
                    worksheet.Cells[row, 8].Value = item.雷射_X軸_PS;
                    worksheet.Cells[row, 9].Value = item.雷射_Y軸_P;
                    worksheet.Cells[row, 10].Value = item.雷射_Y軸_PS;
                    worksheet.Cells[row, 11].Value = item.雷射_Z軸_P;
                    worksheet.Cells[row, 12].Value = item.雷射_Z軸_PS;
                    worksheet.Cells[row, 13].Value = item.循圓_XY軸;
                    worksheet.Cells[row, 14].Value = item.幾何_工作台面平行度_X軸;
                    worksheet.Cells[row, 15].Value = item.幾何_工作台面平行度_Y軸;
                    worksheet.Cells[row, 16].Value = item.幾何_T型槽平行度_X軸方向;
                    worksheet.Cells[row, 17].Value = item.幾何_直角度_XY軸;
                    worksheet.Cells[row, 18].Value = item.幾何_直角度_ZX軸;
                    worksheet.Cells[row, 19].Value = item.幾何_直角度_ZY軸;
                    worksheet.Cells[row, 20].Value = item.幾何_主軸直角度_X軸方向;
                    worksheet.Cells[row, 21].Value = item.幾何_主軸直角度_Y軸方向;
                    worksheet.Cells[row, 22].Value = item.主軸_外圓偏擺;
                    worksheet.Cells[row, 23].Value = item.主軸_端面偏擺_上端;
                    worksheet.Cells[row, 24].Value = item.主軸_端面偏擺_下端;
                    worksheet.Cells[row, 25].Value = item.主軸_錐孔偏擺_300mm;
                }

                package.Save();
            }
        }

        private void CreateNewExcelFile()
        {
            // 確保資料夾存在
            var directory = Path.GetDirectoryName(ExcelPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("驗收資料");

                // 建立標題列
                string[] headers = new[]
                {
                    "出機日期", "機號", "機型", "光學尺_X軸", "光學尺_Y軸", "光學尺_Z軸",
                    "雷射_X軸_P", "雷射_X軸_PS", "雷射_Y軸_P", "雷射_Y軸_PS",
                    "雷射_Z軸_P", "雷射_Z軸_PS", "循圓_XY軸",
                    "幾何_工作台面平行度_X軸", "幾何_工作台面平行度_Y軸", "幾何_T型槽平行度_X軸方向",
                    "幾何_直角度_XY軸", "幾何_直角度_ZX軸", "幾何_直角度_ZY軸",
                    "幾何_主軸直角度_X軸方向", "幾何_主軸直角度_Y軸方向",
                    "主軸_外圓偏擺", "主軸_端面偏擺_上端", "主軸_端面偏擺_下端", "主軸_錐孔偏擺_300mm"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                }

                package.SaveAs(new FileInfo(ExcelPath));
            }
        }

        private T? GetCellValue<T>(ExcelWorksheet worksheet, int row, int col)
        {
            try
            {
                var cellValue = worksheet.Cells[row, col].Value;
                if (cellValue == null)
                    return default;

                if (typeof(T) == typeof(decimal?) || typeof(T) == typeof(decimal))
                {
                    if (decimal.TryParse(cellValue.ToString(), out decimal result))
                        return (T)(object)result;
                    return default;
                }

                return worksheet.Cells[row, col].GetValue<T>();
            }
            catch
            {
                return default;
            }
        }
    }
}