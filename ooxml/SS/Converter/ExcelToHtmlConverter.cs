/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System.Globalization;
using System.IO;

namespace NPOI.SS.Converter
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Xml;
    using NPOI.SS.Util;
    using NPOI.SS.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.Util;
    using NPOI.HSSF.Util;
    using NPOI.SS;
    using NPOI.HSSF.UserModel;
    using NPOI.HPSF;
    using NPOI.XSSF.UserModel;
    using NPOI.XSSF.Model;
    using System.Reflection;
    using System.Linq;
    using SixLabors.ImageSharp;
    using NPOI.OpenXmlFormats.Dml.Spreadsheet;
    using NPOI.POIFS.Storage;

    public class ExcelToHtmlConverter
    {
        POILogger logger = POILogFactory.GetLogger(typeof(ExcelToHtmlConverter));

        public ExcelToHtmlConverter()
        {
            XmlDocument doc = new XmlDocument();
            htmlDocumentFacade = new HtmlDocumentFacade(doc);
            cssClassTable = htmlDocumentFacade.GetOrCreateCssClass("table", "t",
                "border-collapse:collapse;border-spacing:0;table-layout:fixed;");
        }

        protected static double GetColumnWidth(ISheet sheet, int columnIndex)
        {
            return ExcelToHtmlUtils.GetColumnWidthInPx(sheet.GetColumnWidth(columnIndex));
        }

        private DataFormatter _formatter = new DataFormatter();

        private string cssClassContainerCell = null;

        private string cssClassContainerDiv = null;

        private string cssClassTable;

        private Dictionary<int, string> rotationClassNames = new Dictionary<int, string>();

        private HtmlDocumentFacade htmlDocumentFacade;

        private bool outputColumnHeaders = true;

        /// <summary>
        /// 是否输出列头
        /// </summary>
        public bool OutputColumnHeaders
        {
            get { return outputColumnHeaders; }
            set { outputColumnHeaders = value; }
        }

        private bool outputHiddenColumns = false;

        /// <summary>
        /// 是否输出隐藏的列
        /// </summary>
        public bool OutputHiddenColumns
        {
            get { return outputHiddenColumns; }
            set { outputHiddenColumns = value; }
        }

        private bool outputHiddenRows = false;

        /// <summary>
        /// 是否输出隐藏的行
        /// </summary>
        public bool OutputHiddenRows
        {
            get { return outputHiddenRows; }
            set { outputHiddenRows = value; }
        }

        private bool outputLeadingSpacesAsNonBreaking = true;

        /// <summary>
        /// 是否输出文本前的空格
        /// </summary>
        public bool OutputLeadingSpacesAsNonBreaking
        {
            get { return outputLeadingSpacesAsNonBreaking; }
            set { outputLeadingSpacesAsNonBreaking = value; }
        }

        private bool outputRowNumbers = true;

        /// <summary>
        /// 是否输出行号
        /// </summary>
        public bool OutputRowNumbers
        {
            get { return outputRowNumbers; }
            set { outputRowNumbers = value; }
        }

        private bool useDivsToSpan = false;

        /// <summary>
        /// 在跨列的单元格使用DIV标记
        /// </summary>
        public bool UseDivsToSpan
        {
            get { return useDivsToSpan; }
            set { useDivsToSpan = value; }
        }

        public bool ApplyTextRotation { get; set; }
        public bool DisableFormulas { get; set; }

        public static XmlDocument Process(string excelFile)
        {
            IWorkbook workbook = WorkbookFactory.Create(excelFile, null);
            var excelToHtmlConverter = new ExcelToHtmlConverter();
            excelToHtmlConverter.ProcessWorkbook(workbook);
            return excelToHtmlConverter.Document;
        }

        public XmlDocument Document
        {
            get
            {
                return htmlDocumentFacade.Document;
            }
        }

        public void ProcessWorkbook(IWorkbook workbook)
        {
            ProcessDocumentInformation(workbook);

            if(UseDivsToSpan)
            {
                // prepare CSS classes for later usage
                this.cssClassContainerCell = htmlDocumentFacade
                    .GetOrCreateCssClass("td", "c",
                        "padding:0;margin:0;align:left;vertical-align:top;");
                this.cssClassContainerDiv = htmlDocumentFacade.GetOrCreateCssClass(
                    "div", "d", "position:relative;");
            }

            for(int s = 0; s < workbook.NumberOfSheets; s++)
            {
                var sheet = workbook.GetSheetAt(s);
                ProcessSheet(sheet);
            }

            htmlDocumentFacade.UpdateStylesheet();
        }

        protected void ProcessSheet(ISheet sheet)
        {
            ProcessSheetHeader(htmlDocumentFacade.Body, sheet);

            int lastNotEmptyRowNum = GetLastNotEmptyRowNum(sheet);
            if(lastNotEmptyRowNum <= 0)
                return;

            var table = htmlDocumentFacade.CreateTable();
            table.SetAttribute("class", cssClassTable);

            var tableBody = htmlDocumentFacade.CreateTableBody();

            var images = GetPicturesAndShapes(sheet).Where(x => x is not null).ToArray();

            var imagesDict = images.GroupBy(i => i.CellCoordinates.Row)
                .ToDictionary(g => g.Key, g => g.ToArray());

            var mergedRanges = ExcelToHtmlUtils.BuildMergedRangesMap(sheet);
            int maxSheetColumns = GetLastColumn(sheet, lastNotEmptyRowNum);

            for(int r = 0; r <= lastNotEmptyRowNum; r++)
            {
                IRow row = sheet.GetRow(r);
                bool isNullRow = row is null;
                if(isNullRow)
                    row = sheet.CreateRow(r);

                if(!OutputHiddenRows && row.ZeroHeight)
                    continue;

                var tableRowElement = htmlDocumentFacade.CreateTableRow();
                var heightInPx = Units.PointsToPixel(row.HeightInPoints);
                tableRowElement.SetAttribute("height", heightInPx.ToString(CultureInfo.InvariantCulture));
                htmlDocumentFacade.AddStyleClass(tableRowElement, "r", "height:"
                                                                       + heightInPx + "px;");

                var rowPictures = imagesDict.ContainsKey(r) ? imagesDict[r] : System.Array.Empty<Picture>();
                ProcessRow(mergedRanges, row,
                    tableRowElement, isNullRow, rowPictures, maxSheetColumns);

                tableBody.AppendChild(tableRowElement);
            }

            var tableWidth = ProcessColumnWidths(sheet, maxSheetColumns, table);
            table.SetAttribute("width", tableWidth.ToString(CultureInfo.InvariantCulture));
            table.SetAttribute("style", $"min-width:{tableWidth}px;");


            if(OutputColumnHeaders)
            {
                ProcessColumnHeaders(sheet, maxSheetColumns, table);
            }

            table.AppendChild(tableBody);

            htmlDocumentFacade.Body.AppendChild(table);
        }

        protected int GetLastNotEmptyRowNum(ISheet sheet)
        {
            var lastRow = 0;

            for(int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);

                if(row is not null && row.PhysicalNumberOfCells > 0)
                    lastRow = i;
            }

            return lastRow;
        }

        protected int GetLastColumn(ISheet sheet, int lastRowNum)
        {
            var lastColumn = 0;

            for(var i = 0; i <= lastRowNum; i++)
            {
                var row = sheet.GetRow(i);

                if(row is null)
                    continue;

                if(row.LastCellNum > lastColumn)
                    lastColumn = row.LastCellNum;
            }

            return lastColumn;
        }

        protected Picture[] GetPicturesAndShapes(ISheet sheet)
        {
            if(sheet is XSSFSheet xssfSheet)
                return GetXSSFPicturesAndShapes(xssfSheet);
            if(sheet is HSSFSheet hssfSheet)
                return GetHSSFPicturesAndShapes(hssfSheet);
            return System.Array.Empty<Picture>();
        }

        protected Picture[] GetXSSFPicturesAndShapes(XSSFSheet sheet)
        {
            var drawing = sheet.GetDrawingPatriarch();
            
            if (drawing is null)
	            return System.Array.Empty<Picture>();
            
            return drawing.GetShapes().Select(s => ConvertShapeToPicture(s, sheet)).ToArray();
        }

        protected Picture[] GetHSSFPicturesAndShapes(HSSFSheet sheet)
        {
            var drawing = (HSSFPatriarch) sheet.DrawingPatriarch;

            if (drawing is null)
	            return System.Array.Empty<Picture>();
            
            return drawing.GetShapes().Select(s => ConvertShapeToPicture(s, sheet)).ToArray();
        }

        protected Picture ConvertShapeToPicture(XSSFShape shape, ISheet sheet)
        {
            if(shape is XSSFPicture picture)
                return CreatePicture(picture.PictureData.Data, picture.PictureData.SuggestFileExtension(),
                    picture.ClientAnchor, picture.Sheet);

            if(shape is XSSFSimpleShape simpleShape)
                return ConvertSimpleShapeToPicture(simpleShape, sheet);

            return null;
        }

        protected Picture ConvertShapeToPicture(HSSFShape shape, ISheet sheet)
        {
            if(shape is HSSFPicture picture)
                return CreatePicture(picture.PictureData.Data, picture.PictureData.SuggestFileExtension(),
                    picture.ClientAnchor, picture.Sheet);

            if(shape is HSSFSimpleShape simpleShape)
                return ConvertSimpleShapeToPicture(simpleShape, sheet);

            return null;
        }

        protected Picture ConvertSimpleShapeToPicture(HSSFSimpleShape shape, ISheet sheet)
        {
            if(shape.Anchor is not HSSFClientAnchor clientAnchor)
                return null;

            if (TryGetSimpleShapeData(@"SS\Converter\xls", shape.ShapeType, out var data))
                return CreatePicture(data.data, data.extension, clientAnchor, sheet);

            return null;
        }
        
        protected Picture ConvertSimpleShapeToPicture(XSSFSimpleShape shape, ISheet sheet)
        {
            if(shape.anchor is not XSSFClientAnchor clientAnchor)
                return null;

            if (TryGetSimpleShapeData(@"SS\Converter\xlsx", shape.ShapeType, out var data))
                return CreatePicture(data.data, data.extension, clientAnchor, sheet);

            return null;
        }

        private Dictionary<string, Dictionary<int, (byte[], string)>> cache = new ();

        protected bool TryGetSimpleShapeData(string path, int id, out (byte[] data, string extension) data)
        {
            if(!cache.ContainsKey(path))
                cache[path] = new Dictionary<int, (byte[], string)>();

            if(cache[path].TryGetValue(id, out data))
                return true;

            if(!Directory.Exists(path))
                return false;

            var fileName = Directory
                .GetFiles(path).SingleOrDefault(x => Path.GetFileNameWithoutExtension(x) == id.ToString());

            if(fileName == null)
                return false;

            data = (File.ReadAllBytes(fileName), Path.GetExtension(fileName));
            cache[path][id] = data;
            return true;
        }

        protected Picture CreatePicture(byte[] data, string extension, IClientAnchor anchor, ISheet sheet)
        {
            return new Picture() {
                Data = data,
                Extension = extension,
                CellCoordinates = new CellCoordinates(anchor.Col1, anchor.Row1),
                Offset = ImageUtils.GetAnchorOffsetInPixels(anchor, sheet),
                Size = ImageUtils.GetAnchorDimensionsInPixels(anchor, sheet)
            };
        }

        protected void ProcessSheetHeader(XmlElement htmlBody, ISheet sheet)
        {
            XmlElement h2 = htmlDocumentFacade.CreateHeader2();
            h2.AppendChild(htmlDocumentFacade.CreateText(sheet.SheetName));
            htmlBody.AppendChild(h2);
        }

        protected void ProcessDocumentInformation(IWorkbook workbook)
        {
            if(workbook is NPOI.HSSF.UserModel.HSSFWorkbook)
            {
                SummaryInformation summaryInformation = ((HSSFWorkbook) workbook).SummaryInformation;
                if(summaryInformation != null)
                {
                    if(!string.IsNullOrEmpty(summaryInformation.Title))
                        htmlDocumentFacade.Title = summaryInformation.Title;

                    if(!string.IsNullOrEmpty(summaryInformation.Author))
                        htmlDocumentFacade.AddAuthor(summaryInformation.Author);

                    if(!string.IsNullOrEmpty(summaryInformation.Keywords))
                        htmlDocumentFacade.AddKeywords(summaryInformation.Keywords);

                    if(!string.IsNullOrEmpty(summaryInformation.Comments))
                        htmlDocumentFacade.AddDescription(summaryInformation.Comments);
                }
            }
            else if(workbook is NPOI.XSSF.UserModel.XSSFWorkbook)
            {
                POIXMLProperties props = ((NPOI.XSSF.UserModel.XSSFWorkbook) workbook).GetProperties();
                if(!string.IsNullOrEmpty(props.CoreProperties.Title))
                {
                    htmlDocumentFacade.Title = props.CoreProperties.Title;
                }

                if(!string.IsNullOrEmpty(props.CoreProperties.Creator))
                    htmlDocumentFacade.AddAuthor(props.CoreProperties.Creator);

                if(!string.IsNullOrEmpty(props.CoreProperties.Keywords))
                    htmlDocumentFacade.AddKeywords(props.CoreProperties.Keywords);

                if(!string.IsNullOrEmpty(props.CoreProperties.Description))
                    htmlDocumentFacade.AddDescription(props.CoreProperties.Description);
            }
        }

        protected void ProcessRow(CellRangeAddress[][] mergedRanges, IRow row,
            XmlElement tableRowElement, bool isNullRow, Picture[] rowPictures, int maxCollNumTotal)
        {
            if(isNullRow)
                row.CreateCell(0).SetBlank();

            ISheet sheet = (ISheet) row.Sheet;
            var lastImgCol = rowPictures.Length != 0 ? rowPictures.Select(i => i.CellCoordinates.Column).Max() : -1;
            int maxRowColIx = Math.Max(row.LastCellNum, lastImgCol + 1);
            if(maxRowColIx <= 0)
                return;

            List<XmlElement> emptyCells = new List<XmlElement>(maxRowColIx);

            if(OutputRowNumbers)
            {
                XmlElement tableRowNumberCellElement = htmlDocumentFacade.CreateTableHeaderCell();
                ProcessRowNumber(row, tableRowNumberCellElement);
                emptyCells.Add(tableRowNumberCellElement);
            }

            var pictureDict = rowPictures.ToDictionary(i => i.CellCoordinates.Column);

            int maxRenderedColumn = 0;
            for(int colIx = 0; colIx < maxRowColIx; colIx++)
            {
                if(!OutputHiddenColumns && sheet.IsColumnHidden(colIx))
                    continue;

                CellRangeAddress range = ExcelToHtmlUtils.GetMergedRange(
                    mergedRanges, row.RowNum, colIx);

                if(range != null && (range.FirstColumn != colIx || range.FirstRow != row.RowNum))
                {
                    colIx = range.LastColumn;
                    continue;
                }

                ICell cell = row.GetCell(colIx);
				
                var prevBottom = cell is null ? default : cell.CellStyle.BorderBottom;
                var prevRight = cell is null ? default : cell.CellStyle.BorderRight;
                
                if(range != null)
                {
                    var bottomCell = row.Sheet.GetRow(range.LastRow)?.GetCell(range.LastColumn);

                    if (cell != null && bottomCell != null)
                    {
	                    if (cell.CellStyle.BorderBottom != bottomCell.CellStyle.BorderBottom ||
	                        cell.CellStyle.BorderRight != bottomCell.CellStyle.BorderRight)
	                    {
		                    cell.CellStyle.BorderBottom = bottomCell.CellStyle.BorderBottom;
		                    cell.CellStyle.BorderRight = bottomCell.CellStyle.BorderRight;
	                    }
                    }
                }

                double divWidthPx = 0;

                XmlElement tableCellElement = htmlDocumentFacade.CreateTableCell();
                tableCellElement.SetAttribute("style", "padding: 0px;");

                int width;
                if(range != null)
                    width = GetCellWidth(sheet, range.FirstColumn, range.LastColumn);
                else
                    width = GetCellWidth(sheet, colIx, colIx);

                tableCellElement.SetAttribute("width", $"{width.ToString(CultureInfo.InvariantCulture)}");

                if(pictureDict.ContainsKey(colIx))
                {
                    var picture = pictureDict[colIx];
                    var div = htmlDocumentFacade.CreateBlock();
                    var meta = $"data:image/{picture.Extension};base64, ";
                    var imageElem =
                        htmlDocumentFacade.CreateImage(meta + Convert.ToBase64String(picture.Data));

                    imageElem.SetAttribute("width", picture.Size.Width.ToString(CultureInfo.InvariantCulture));
                    imageElem.SetAttribute("height", picture.Size.Height.ToString(CultureInfo.InvariantCulture));
                    div.SetAttribute("style", BuildStyleForImageDiv(picture.Size, picture.Offset));
                    div.AppendChild(imageElem);

                    tableCellElement.AppendChild(div);
                    tableCellElement.Attributes["style"].Value += "position:relative;";
                }

                if(range != null)
                {
                    if(range.FirstColumn != range.LastColumn)
                        tableCellElement.SetAttribute("colspan", (range.LastColumn - range.FirstColumn + 1).ToString());
                    if(range.FirstRow != range.LastRow)
                        tableCellElement.SetAttribute("rowspan", (range.LastRow - range.FirstRow + 1).ToString());
                }

                if(cell != null)
                    ProcessCell(cell, tableCellElement, 
	                    GetColumnWidth(sheet, colIx), divWidthPx, 
	                    row.Height / 20f);

                tableRowElement.AppendChild(tableCellElement);

                if (cell != null)
                {
	                cell.CellStyle.BorderRight = prevRight;
	                cell.CellStyle.BorderBottom = prevBottom;
                }
            }

            // creates a cell to fill the row to make all rows same width

            if(maxRowColIx < maxCollNumTotal)
            {
                var cell = htmlDocumentFacade.CreateTableCell();
                cell.SetAttribute("colspan", (maxCollNumTotal - maxRowColIx).ToString());
                if(row.HeightInPoints > 10)
                {
                    var text = htmlDocumentFacade.CreateText("\u00a0");
                    cell.AppendChild(text);
                }

                tableRowElement.AppendChild(cell);
            }
        }

        private string BuildStyleForImageDiv(Size imgSize, Point offset)
        {
            var sb = new StringBuilder();

            sb.Append("position: absolute;");
            sb.Append($"width:{imgSize.Width}px;");
            sb.Append($"height:{imgSize.Height}px;");
            sb.Append($"margin-top:{offset.Y}px;");
            sb.Append($"margin-left:{offset.X}px;");
            sb.Append("top:0px;");
            sb.Append("left:0px;");

            return sb.ToString();
        }

        private int GetCellWidth(ISheet sheet, int start, int end)
        {
            var res = 0;
            for(int i = start; i <= end; i++)
            {
                res += (int) Math.Round(ExcelToHtmlUtils.GetColumnWidthInPx(sheet.GetColumnWidth(i)));
            }

            return res;
        }

        private string GetRowName(IRow row)
        {
            return (row.RowNum + 1).ToString();
        }

        protected void ProcessRowNumber(IRow row, XmlElement tableRowNumberCellElement)
        {
            tableRowNumberCellElement.SetAttribute("class", "rownumber");
            XmlText text = htmlDocumentFacade.CreateText(GetRowName(row));
            tableRowNumberCellElement.AppendChild(text);
        }

        /**
     * Creates COLGROUP element with width specified for all columns. (Except
     * first if <tt>{@link #isOutputRowNumbers()}==true</tt>)
     */
        protected int ProcessColumnWidths(ISheet sheet, int maxSheetColumns,
            XmlElement table)
        {
            // draw COLS after we know max column number
            XmlElement columnGroup = htmlDocumentFacade.CreateTableColumnGroup();
            if(OutputRowNumbers)
            {
                columnGroup.AppendChild(htmlDocumentFacade.CreateTableColumn());
            }

            var tableWidth = 0;
            for(int c = 0; c < maxSheetColumns; c++)
            {
                if(!OutputHiddenColumns && sheet.IsColumnHidden(c))
                    continue;

                XmlElement col = htmlDocumentFacade.CreateTableColumn();
                var colWidth = (int) Math.Round(GetColumnWidth(sheet, c));
                col.SetAttribute("width", $"{colWidth.ToString(CultureInfo.InvariantCulture)}");
                columnGroup.AppendChild(col);

                tableWidth += colWidth;
            }

            table.AppendChild(columnGroup);

            return tableWidth;
        }

        protected void ProcessColumnHeaders(ISheet sheet, int maxSheetColumns,
            XmlElement table)
        {
            XmlElement tableHeader = htmlDocumentFacade.CreateTableHeader();
            table.AppendChild(tableHeader);

            XmlElement tr = htmlDocumentFacade.CreateTableRow();

            if(OutputRowNumbers)
            {
                // empty row at left-top corner
                tr.AppendChild(htmlDocumentFacade.CreateTableHeaderCell());
            }

            for(int c = 0; c < maxSheetColumns; c++)
            {
                if(!OutputHiddenColumns && sheet.IsColumnHidden(c))
                    continue;

                XmlElement th = htmlDocumentFacade.CreateTableHeaderCell();
                string text = GetColumnName(c);
                th.AppendChild(htmlDocumentFacade.CreateText(text));
                tr.AppendChild(th);
            }

            tableHeader.AppendChild(tr);
        }

        protected string GetColumnName(int columnIndex)
        {
            return (columnIndex + 1).ToString();
        }

        protected bool ProcessCell(ICell cell, XmlElement tableCellElement,
            double normalWidthPx, double maxSpannedWidthPx, float normalHeightPt)
        {
            ICellStyle cellStyle = cell.CellStyle;

            string value;
            var isRichString = false;
            switch(cell.CellType)
            {
                case CellType.String:
                    value = cell.RichStringCellValue.String;
                    if(!string.IsNullOrWhiteSpace(value))
                    {
                        isRichString = true;
                        AddRichString(cell.RichStringCellValue, tableCellElement, cell.Sheet.Workbook);
                        break;
                    }

                    value = "";
                    break;
                case CellType.Formula:
                    if(DisableFormulas)
                    {
                        value = "";
                        break;
                    }

                    switch(cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            IRichTextString str = cell.RichStringCellValue;
                            if(str != null && str.Length > 0)
                            {
                                value = (str.String);
                                value = string.IsNullOrWhiteSpace(value) ? "" : value;
                            }
                            else
                            {
                                value = string.Empty;
                            }

                            break;
                        case CellType.Numeric:
                            ICellStyle style = cellStyle;
                            if(style == null)
                            {
                                value = cell.NumericCellValue.ToString();
                            }
                            else
                            {
                                value = (_formatter.FormatRawCellContents(cell.NumericCellValue, style.DataFormat,
                                    style.GetDataFormatString()));
                            }

                            break;
                        case CellType.Boolean:
                            value = cell.BooleanCellValue.ToString();
                            break;
                        case CellType.Error:
                            value = ErrorEval.GetText(cell.ErrorCellValue);
                            break;
                        default:
                            logger.Log(POILogger.WARN,
                                "Unexpected cell cachedFormulaResultType (" + cell.CachedFormulaResultType.ToString() +
                                ")");
                            value = string.Empty;
                            break;
                    }

                    break;
                case CellType.Blank:
                    value = string.Empty;
                    break;
                case CellType.Numeric:
                    //value = _formatter.FormatCellValue(cell); - not working right
                    value = cell.NumericCellValue.ToString(CultureInfo.CurrentCulture);
                    break;
                case CellType.Boolean:
                    value = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    value = ErrorEval.GetText(cell.ErrorCellValue);
                    break;
                default:
                    logger.Log(POILogger.WARN, "Unexpected cell type (" + cell.CellType.ToString() + ")");
                    return true;
            }

            bool noText = string.IsNullOrEmpty(value);
            bool wrapInDivs = !noText && UseDivsToSpan && !cellStyle.WrapText;

            short cellStyleIndex = cellStyle.Index;
            if(cellStyleIndex != 0)
            {
                IWorkbook workbook = cell.Row.Sheet.Workbook as IWorkbook;
                string mainCssClass = GetStyleClassName(workbook, cellStyle);
                if(wrapInDivs)
                {
                    tableCellElement.SetAttribute("class", mainCssClass + " "
                                                                        + cssClassContainerCell);
                }
                else
                {
                    tableCellElement.SetAttribute("class", mainCssClass);
                }
            }
            
            // add this style cause microsoft excel does
            if(cellStyle.Rotation != 0 && cellStyle.Rotation >= -180 && cellStyle.Rotation <= 180)
	            tableCellElement.Attributes["style"].Value += $"mso-rotate:{cellStyle.Rotation};";

            if(isRichString)
                return false;

            if(OutputLeadingSpacesAsNonBreaking && value.StartsWith(" "))
            {
                StringBuilder builder = new StringBuilder();
                for(int c = 0; c < value.Length; c++)
                {
                    if(value[c] != ' ')
                        break;
                    builder.Append('\u00a0');
                }

                if(value.Length != builder.Length)
                    builder.Append(value.Substring(builder.Length));

                value = builder.ToString();
            }

            if(value == "" && cell.Row.HeightInPoints > 10)
                value = "\u00a0";
            XmlText text = htmlDocumentFacade.CreateText(value);

            if(cellStyle.Rotation != 0 && ApplyTextRotation)
            {
                var div = htmlDocumentFacade.CreateBlock();
                var className = GetRotationClassName(cell, cellStyle.Rotation);

                div.SetAttribute("class", className);
                div.AppendChild(htmlDocumentFacade.CreateText(value));

                tableCellElement.AppendChild(div);
            }
            else if(wrapInDivs)
            {
                XmlElement outerDiv = htmlDocumentFacade.CreateBlock();
                outerDiv.SetAttribute("class", this.cssClassContainerDiv);

                XmlElement innerDiv = htmlDocumentFacade.CreateBlock();
                StringBuilder innerDivStyle = new StringBuilder();
                innerDivStyle.Append("position:absolute;min-width:");
                innerDivStyle.Append(normalWidthPx);
                innerDivStyle.Append("px;");
                if(maxSpannedWidthPx != int.MaxValue)
                {
                    innerDivStyle.Append("max-width:");
                    innerDivStyle.Append(maxSpannedWidthPx);
                    innerDivStyle.Append("px;");
                }

                innerDivStyle.Append("overflow:hidden;max-height:");
                innerDivStyle.Append(normalHeightPt);
                innerDivStyle.Append("pt;white-space:nowrap;");
                ExcelToHtmlUtils.AppendAlign(innerDivStyle, cellStyle.Alignment, cellStyle.VerticalAlignment);
                htmlDocumentFacade.AddStyleClass(outerDiv, "d", innerDivStyle.ToString());

                innerDiv.AppendChild(text);
                outerDiv.AppendChild(innerDiv);
                tableCellElement.AppendChild(outerDiv);
            }
            else
            {
                tableCellElement.AppendChild(text);
            }

            return string.IsNullOrEmpty(value) && cellStyleIndex == 0;
        }

        protected void AddRichString(IRichTextString richTextString, XmlElement tableCell, IWorkbook workbook)
        {
            if (richTextString is XSSFRichTextString xssfRichTextString)
                AddRichString(xssfRichTextString, tableCell);
            else if (richTextString is HSSFRichTextString hssfRichTextString)
                AddRichString(hssfRichTextString, tableCell, workbook);
        }

        protected void AddRichString(HSSFRichTextString richTextString, XmlElement tableCell, IWorkbook workbook)
        {
            var prev = 0;
            for(var i = 0; i < richTextString.NumFormattingRuns; i++)
            {
                var cur = richTextString.GetIndexOfFormattingRun(i);
                
                var font = workbook.GetFontAt(richTextString.GetFontAtIndex(prev));
                var text = richTextString.String.Substring(prev, cur - prev);
                prev = cur;
                
                AppendText(font, text, tableCell);
            }
            AppendText(
                workbook.GetFontAt(richTextString.GetFontAtIndex(prev)), richTextString.String.Substring(prev), tableCell);
        }

        protected void AppendText(IFont font, string text, XmlElement tableCell)
        {
            if(font.TypeOffset == FontSuperScript.None)
                tableCell.AppendChild(htmlDocumentFacade.CreateText(text));
            else if(font.TypeOffset == FontSuperScript.Super)
                tableCell.AppendChild(htmlDocumentFacade.CreateSup(text));
            else
                tableCell.AppendChild(htmlDocumentFacade.CreateSub(text));
        }
        
        protected void AddRichString(XSSFRichTextString richTextString, XmlElement tableCell)
        {
            var prev = 0;
            for(var i = 1; i <= richTextString.NumFormattingRuns; i++)
            {
                var cur = i != richTextString.NumFormattingRuns
                    ? richTextString.GetIndexOfFormattingRun(i)
                    : richTextString.String.Length;
                var font = richTextString.GetFontAtIndex(prev);
                var text = richTextString.String.Substring(prev, cur - prev);
                prev = cur;
                if(font.TypeOffset == FontSuperScript.None)
                    tableCell.AppendChild(htmlDocumentFacade.CreateText(text));
                else if(font.TypeOffset == FontSuperScript.Super)
                    tableCell.AppendChild(htmlDocumentFacade.CreateSup(text));
                else
                    tableCell.AppendChild(htmlDocumentFacade.CreateSub(text));
            }
        }

        protected string GetRotationClassName(ICell cell, int rotationInDegrees)
        {
            if(rotationClassNames.ContainsKey(rotationInDegrees))
                return rotationClassNames[rotationInDegrees];

            var style = BuildRotationClass(cell, rotationInDegrees);
            var cssClass = htmlDocumentFacade.GetOrCreateCssClass("div", "rot", style);
            rotationClassNames.Add(rotationInDegrees, cssClass);

            return cssClass;
        }

        protected string BuildRotationClass(ICell cell, int rotationInDegrees)
        {
            var strRotation = (rotationInDegrees + 90).ToString(CultureInfo.InvariantCulture);
            var width = 96f / 72 * cell.Row.HeightInPoints;

            return "writing-mode: vertical-rl;" +
                   $"transform: rotate({strRotation}deg);" +
                   "white-space: wrap;" +
                   "word-break: break-all;" +
                   $"height:{width};";
        }

        protected string GetStyleClassName(IWorkbook workbook, ICellStyle cellStyle)
        {
            String cssStyle = BuildStyle(workbook, cellStyle);
            String cssClass = htmlDocumentFacade.GetOrCreateCssClass("td", "c",
	            cssStyle);

            return cssClass;
        }

        protected String BuildStyle(IWorkbook workbook, ICellStyle cellStyle)
        {
            StringBuilder style = new StringBuilder();

            if(workbook is HSSFWorkbook)
            {
                HSSFPalette palette = ((HSSFWorkbook) workbook).GetCustomPalette();
                style.Append("white-space: pre-wrap; ");
                ExcelToHtmlUtils.AppendAlign(style, cellStyle.Alignment, cellStyle.VerticalAlignment);

                if(cellStyle.FillPattern == FillPattern.NoFill)
                {
                    // no fill
                }
                else if(cellStyle.FillPattern == FillPattern.SolidForeground)
                {
                    //cellStyle.
                    //HSSFColor.
                    HSSFColor foregroundColor = palette.GetColor(cellStyle.FillForegroundColor);
                    if(foregroundColor != null)
                        style.AppendFormat("background-color:{0}; ", ExcelToHtmlUtils.GetColor(foregroundColor));
                }
                else
                {
                    HSSFColor backgroundColor = palette.GetColor(cellStyle.FillBackgroundColor);
                    if(backgroundColor != null)
                        style.AppendFormat("background-color:{0}; ", ExcelToHtmlUtils.GetColor(backgroundColor));
                }
            }
            else
            {
                style.Append("white-space: pre-wrap; ");
                ExcelToHtmlUtils.AppendAlign(style, cellStyle.Alignment, cellStyle.VerticalAlignment);

                if(cellStyle.FillPattern == FillPattern.NoFill)
                {
                    // no fill
                }
                else if(cellStyle.FillPattern == FillPattern.SolidForeground)
                {
                    //cellStyle
                    IndexedColors clr = IndexedColors.ValueOf(cellStyle.FillForegroundColor);
                    string hexstring = null;
                    if(clr != null)
                    {
                        hexstring = clr.HexString;
                    }
                    else
                    {
                        XSSFColor foregroundColor = (XSSFColor) cellStyle.FillForegroundColorColor;
                        if(foregroundColor != null)
                            hexstring = ExcelToHtmlUtils.GetColor(foregroundColor);
                    }

                    if(hexstring != null)
                        style.AppendFormat("background-color:{0}; ", hexstring);
                }
                else
                {
                    IndexedColors clr = IndexedColors.ValueOf(cellStyle.FillBackgroundColor);
                    string hexstring = null;
                    if(clr != null)
                    {
                        hexstring = clr.HexString;
                    }
                    else
                    {
                        XSSFColor backgroundColor = (XSSFColor) cellStyle.FillBackgroundColorColor;
                        if(backgroundColor != null)
                            hexstring = ExcelToHtmlUtils.GetColor(backgroundColor);
                        hexstring = ExcelToHtmlUtils.GetColor(backgroundColor);
                    }

                    if(hexstring != null)
                        style.AppendFormat("background-color:{0}; ", hexstring);
                }
            }

            BuildStyle_Border(workbook, style, "top", cellStyle.BorderTop, cellStyle.TopBorderColor);
            BuildStyle_Border(workbook, style, "right", cellStyle.BorderRight, cellStyle.RightBorderColor);
            BuildStyle_Border(workbook, style, "bottom", cellStyle.BorderBottom, cellStyle.BottomBorderColor);
            BuildStyle_Border(workbook, style, "left", cellStyle.BorderLeft, cellStyle.LeftBorderColor);

            IFont font = cellStyle.GetFont(workbook);
            BuildStyle_Font(workbook, style, font);

            return style.ToString();
        }

        private void BuildStyle_Border(IWorkbook workbook, StringBuilder style,
            String type, BorderStyle xlsBorder, short borderColor)
        {
            if(xlsBorder == BorderStyle.None)
                return;

            StringBuilder borderStyle = new StringBuilder();
            borderStyle.Append(ExcelToHtmlUtils.GetBorderWidth(xlsBorder));
            borderStyle.Append(' ');
            borderStyle.Append(ExcelToHtmlUtils.GetBorderStyle(xlsBorder));

            if(workbook is HSSFWorkbook)
            {
                var customPalette = ((HSSFWorkbook) workbook).GetCustomPalette();
                HSSFColor color = null;
                if(customPalette != null)
                    color = customPalette.GetColor(borderColor);
                if(color != null)
                {
                    borderStyle.Append(' ');
                    borderStyle.Append(ExcelToHtmlUtils.GetColor(color));
                }
            }
            else
            {
                IndexedColors clr = IndexedColors.ValueOf(borderColor);
                if(clr != null)
                {
                    borderStyle.Append(' ');
                    borderStyle.Append(clr.HexString);
                }
                else
                {
                    XSSFColor color = null;
                    var stylesSource = ((XSSFWorkbook) workbook).GetStylesSource();
                    if(stylesSource != null)
                    {
                        var theme = stylesSource.GetTheme();
                        if(theme != null)
                            color = theme.GetThemeColor(borderColor);
                    }

                    if(color != null)
                    {
                        borderStyle.Append(' ');
                        borderStyle.Append(ExcelToHtmlUtils.GetColor(color));
                    }
                }
            }

            style.AppendFormat("border-{0}: {1}; ", type, borderStyle);
        }

        void BuildStyle_Font(IWorkbook workbook, StringBuilder style,
            IFont font)
        {
            switch(font.Boldweight)
            {
                case (short) FontBoldWeight.Bold:
                    style.Append("font-weight: bold; ");
                    break;
                case (short) FontBoldWeight.Normal:
                    // by default, not not increase HTML size
                    // style.Append( "font-weight: normal; " );
                    break;
            }

            if(workbook is HSSFWorkbook)
            {
                var customPalette = ((HSSFWorkbook) workbook).GetCustomPalette();
                HSSFColor fontColor = null;
                if(customPalette != null)
                    fontColor = customPalette.GetColor(font.Color);
                if(fontColor != null)
                    style.AppendFormat("color:{0}; ", ExcelToHtmlUtils.GetColor(fontColor));
            }
            else
            {
                IndexedColors clr = IndexedColors.ValueOf(font.Color);
                string hexstring = null;
                if(clr != null)
                {
                    hexstring = clr.HexString;
                }
                else
                {
                    StylesTable st = ((XSSFWorkbook) workbook).GetStylesSource();
                    XSSFColor fontColor = null;
                    if(st != null && st.GetTheme() != null)
                    {
                        fontColor = st.GetTheme().GetThemeColor(font.Color);
                    }
                    else
                    {
                        fontColor = ((XSSFFont) font).GetXSSFColor();
                    }

                    if(fontColor != null)
                        hexstring = ExcelToHtmlUtils.GetColor(fontColor);
                }

                if(hexstring != null)
                    style.AppendFormat("color:{0}; ", hexstring);
            }

            if(font.FontHeightInPoints != 0)
                style.Append("font-size: " + font.FontHeightInPoints.ToString(CultureInfo.InvariantCulture) + "pt; ");
            if(font.IsItalic)
            {
                style.Append("font-style: italic; ");
            }

            var defaultFont = "sans-serif";
            style.Append($"font-family: '{font.FontName}', {defaultFont}");
        }
    }
}