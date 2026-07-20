using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.UserModel;   // reads legacy .xls (BIFF8)
using NPOI.XSSF.UserModel;   // writes new .xlsx (OOXML)
using NPOI.SS.UserModel;

namespace XlsConverter
{
    public class XlsToXlsxFunction
    {
        private readonly ILogger _logger;

        public XlsToXlsxFunction(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<XlsToXlsxFunction>();
        }

        /// <summary>
        /// HTTP trigger: POST an .xls file as the raw request body (or multipart form data),
        /// receive the converted .xlsx file back in the response.
        /// Example: POST /api/ConvertXlsToXlsx  with body = raw bytes of the .xls file
        /// </summary>
        [Function("ConvertXlsToXlsx")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = "ConvertXlsToXlsx")] HttpRequestData req)
        {
            _logger.LogInformation("Received request to convert .xls to .xlsx");

            try
            {
                // 1. Read the incoming .xls stream from the request body
                using var inputStream = new MemoryStream();
                await req.Body.CopyToAsync(inputStream);
                inputStream.Position = 0;

                if (inputStream.Length == 0)
                {
                    var badRequest = req.CreateResponse(HttpStatusCode.BadRequest);
                    await badRequest.WriteStringAsync("No file content received. Send the .xls file as the raw request body.");
                    return badRequest;
                }

                // 2. Convert
                byte[] xlsxBytes = ConvertXlsToXlsx(inputStream);

                // 3. Return the .xlsx file
                var response = req.CreateResponse(HttpStatusCode.OK);
                response.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                response.Headers.Add("Content-Disposition", "attachment; filename=\"converted.xlsx\"");
                await response.Body.WriteAsync(xlsxBytes, 0, xlsxBytes.Length);
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error converting .xls to .xlsx");
                var errorResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                await errorResponse.WriteStringAsync($"Conversion failed: {ex.Message}");
                return errorResponse;
            }
        }

        /// <summary>
        /// Core conversion logic: reads a legacy .xls workbook (HSSFWorkbook)
        /// and rebuilds it as a modern .xlsx workbook (XSSFWorkbook).
        /// Copies cell values, formulas, basic number/date formats, and simple styling.
        /// </summary>
        public static byte[] ConvertXlsToXlsx(Stream xlsInputStream)
        {
            using var oldWorkbook = new HSSFWorkbook(xlsInputStream);
            using var newWorkbook = new XSSFWorkbook();

            for (int sheetIndex = 0; sheetIndex < oldWorkbook.NumberOfSheets; sheetIndex++)
            {
                ISheet oldSheet = oldWorkbook.GetSheetAt(sheetIndex);
                ISheet newSheet = newWorkbook.CreateSheet(oldSheet.SheetName);

                CopySheet(oldSheet, newSheet, newWorkbook);
            }

            using var outputStream = new MemoryStream();
            newWorkbook.Write(outputStream, leaveOpen: true);
            return outputStream.ToArray();
        }

        private static void CopySheet(ISheet oldSheet, ISheet newSheet, IWorkbook newWorkbook)
        {
            // Track cell styles we've already recreated, so we don't create duplicates for every cell
            var styleCache = new System.Collections.Generic.Dictionary<int, ICellStyle>();

            for (int rowIndex = oldSheet.FirstRowNum; rowIndex <= oldSheet.LastRowNum; rowIndex++)
            {
                IRow oldRow = oldSheet.GetRow(rowIndex);
                if (oldRow == null) continue;

                IRow newRow = newSheet.CreateRow(rowIndex);
                newRow.HeightInPoints = oldRow.HeightInPoints;

                for (int cellIndex = oldRow.FirstCellNum; cellIndex < oldRow.LastCellNum; cellIndex++)
                {
                    if (cellIndex < 0) continue;
                    ICell oldCell = oldRow.GetCell(cellIndex);
                    if (oldCell == null) continue;

                    ICell newCell = newRow.CreateCell(cellIndex);
                    CopyCellValue(oldCell, newCell);
                    CopyCellStyle(oldCell, newCell, newWorkbook, styleCache);
                }
            }

            // Preserve column widths
            int maxCellNum = 0;
            for (int rowIndex = oldSheet.FirstRowNum; rowIndex <= oldSheet.LastRowNum; rowIndex++)
            {
                IRow row = oldSheet.GetRow(rowIndex);
                if (row != null && row.LastCellNum > maxCellNum)
                    maxCellNum = row.LastCellNum;
            }

            for (int col = 0; col < maxCellNum; col++)
            {
                try { newSheet.SetColumnWidth(col, oldSheet.GetColumnWidth(col)); }
                catch { /* ignore columns without width info */ }
            }
        }

        private static void CopyCellValue(ICell oldCell, ICell newCell)
        {
            switch (oldCell.CellType)
            {
                case CellType.String:
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;

                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(oldCell))
                        newCell.SetCellValue(oldCell.DateCellValue);
                    else
                        newCell.SetCellValue(oldCell.NumericCellValue);
                    break;

                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;

                case CellType.Formula:
                    // Copy the formula text; NPOI recalculates on open in Excel
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;

                case CellType.Error:
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                    break;

                case CellType.Blank:
                default:
                    newCell.SetBlank();
                    break;
            }
        }

        private static void CopyCellStyle(
            ICell oldCell,
            ICell newCell,
            IWorkbook newWorkbook,
            System.Collections.Generic.Dictionary<int, ICellStyle> styleCache)
        {
            int oldStyleIndex = oldCell.CellStyle.Index;

            if (!styleCache.TryGetValue(oldStyleIndex, out ICellStyle newStyle))
            {
                newStyle = newWorkbook.CreateCellStyle();
                newStyle.CloneStyleFrom(oldCell.CellStyle);
                styleCache[oldStyleIndex] = newStyle;
            }

            newCell.CellStyle = newStyle;
        }
    }
}
