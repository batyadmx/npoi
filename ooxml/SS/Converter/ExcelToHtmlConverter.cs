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

        private Dictionary<short, string> excelStyleToClass = new Dictionary<short, string>();

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

            var images = GetPictures(sheet);

            var imagesDict = images.GroupBy(i => i.GetPreferredSize().Row1)
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

                var rowPictures = imagesDict.ContainsKey(r) ? imagesDict[r] : System.Array.Empty<XSSFPicture>();
                ProcessRow(mergedRanges, row,
                    tableRowElement, isNullRow, rowPictures, maxSheetColumns);

                tableBody.AppendChild(tableRowElement);
            }

            ProcessColumnWidths(sheet, maxSheetColumns, table);

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

        protected XSSFPicture[] GetPictures(ISheet sheet)
        {
            if(sheet is not XSSFSheet)
                throw new ArgumentException("Not a XSSFSheet");

            var drawing = ((XSSFSheet) sheet).GetDrawingPatriarch();

            if(drawing is null)
                return System.Array.Empty<XSSFPicture>();

            return drawing.GetShapes().Select(s => (XSSFPicture) s).ToArray();
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
            XmlElement tableRowElement, bool isNullRow, XSSFPicture[] rowPictures, int maxCollNumTotal)
        {
            if(isNullRow)
                row.CreateCell(0).SetBlank();

            ISheet sheet = (ISheet) row.Sheet;
            var lastImgCol = rowPictures.Length != 0 ? rowPictures.Select(i => i.GetPreferredSize().Col1).Max() : -1;
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

            var pictureDict = rowPictures.ToDictionary(i => i.GetPreferredSize().Col1);

            int maxRenderedColumn = 0;
            for(int colIx = 0; colIx < maxRowColIx; colIx++)
            {
                if(!OutputHiddenColumns && sheet.IsColumnHidden(colIx))
                    continue;

                CellRangeAddress range = ExcelToHtmlUtils.GetMergedRange(
                    mergedRanges, row.RowNum, colIx);

                if(range != null && (range.FirstColumn != colIx || range.FirstRow != row.RowNum))
                    continue;

                ICell cell = row.GetCell(colIx);

                if(range != null)
                {
                    var bottomCell = row.Sheet.GetRow(range.LastRow)?.GetCell(range.LastColumn);

                    if(cell == null)
                    {
                        throw new Exception("Cell in merged region is null");
                    }

                    var newStyle = row.Sheet.Workbook.CreateCellStyle();
                    newStyle.CloneStyleFrom(cell.CellStyle);

                    newStyle.BorderBottom = bottomCell.CellStyle.BorderBottom;
                    newStyle.BorderRight = bottomCell.CellStyle.BorderRight;

                    cell.CellStyle = newStyle;
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
                    var image = pictureDict[colIx];
                    var div = htmlDocumentFacade.CreateBlock();
                    var meta = $"data:image/{image.PictureData.SuggestFileExtension()};base64, ";
                    var imageElem =
                        htmlDocumentFacade.CreateImage(meta + Convert.ToBase64String(image.PictureData.Data));

                    var imageSize = ImageUtils.GetPictureDimensionInPixels(image);
                    var imageOffset = ImageUtils.GetPictureOffsetInPixels(image);
                    imageElem.SetAttribute("width", imageSize.Width.ToString(CultureInfo.InvariantCulture));
                    imageElem.SetAttribute("height", imageSize.Height.ToString(CultureInfo.InvariantCulture));
                    div.SetAttribute("style", BuildStyleForImageDiv(imageSize, imageOffset));
                    div.AppendChild(imageElem);

                    tableCellElement.AppendChild(div);
                }
                else if(range != null)
                {
                    if(range.FirstColumn != range.LastColumn)
                        tableCellElement.SetAttribute("colspan", (range.LastColumn - range.FirstColumn + 1).ToString());
                    if(range.FirstRow != range.LastRow)
                        tableCellElement.SetAttribute("rowspan", (range.LastRow - range.FirstRow + 1).ToString());
                }

                if(cell != null)
                    ProcessCell(cell, tableCellElement, GetColumnWidth(sheet, colIx), divWidthPx, row.Height / 20f);

                tableRowElement.AppendChild(tableCellElement);
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

        private string BuildStyleForImageDiv(Size imgSize, Size offset)
        {
            var sb = new StringBuilder();

            sb.Append("position: absolute;");
            sb.Append($"width:{imgSize.Width}px;");
            sb.Append($"height:{imgSize.Height}px;");
            sb.Append($"margin-top:{offset.Height}px;");
            sb.Append($"margin-left:{offset.Width}px;");

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
        protected void ProcessColumnWidths(ISheet sheet, int maxSheetColumns,
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

            table.SetAttribute("width", tableWidth.ToString(CultureInfo.InvariantCulture));
            table.AppendChild(columnGroup);
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

        protected bool IsTextEmpty(ICell cell)
        {
            string value;
            switch(cell.CellType)
            {
                case CellType.String:
                    // XXX: enrich
                    value = cell.RichStringCellValue.String;
                    break;
                case CellType.Formula:
                    switch(cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            IRichTextString str = cell.RichStringCellValue as IRichTextString;
                            if(str == null || str.Length <= 0)
                                return false;

                            value = str.ToString();
                            break;
                        case CellType.Numeric:
                            ICellStyle style = cell.CellStyle as ICellStyle;
                            if(style == null)
                            {
                                return false;
                            }

                            value = (_formatter.FormatRawCellContents(cell.NumericCellValue, style.DataFormat,
                                style.GetDataFormatString()));
                            break;
                        case CellType.Boolean:
                            value = cell.BooleanCellValue.ToString();
                            break;
                        case CellType.Error:
                            value = ErrorEval.GetText(cell.ErrorCellValue);
                            break;
                        default:
                            value = string.Empty;
                            break;
                    }

                    break;
                case CellType.Blank:
                    value = string.Empty;
                    break;
                case CellType.Numeric:
                    value = _formatter.FormatCellValue(cell);
                    break;
                case CellType.Boolean:
                    value = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    value = ErrorEval.GetText(cell.ErrorCellValue);
                    break;
                default:
                    return true;
            }

            return string.IsNullOrEmpty(value);
        }

        protected bool ProcessCell(ICell cell, XmlElement tableCellElement,
            double normalWidthPx, double maxSpannedWidthPx, float normalHeightPt)
        {
            ICellStyle cellStyle = cell.CellStyle;

            string value;
            switch(cell.CellType)
            {
                case CellType.String:
                    value = cell.RichStringCellValue.String;
                    value = string.IsNullOrWhiteSpace(value) ? "" : value;
                    break;
                case CellType.Formula:
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
            
            // add this style cause microsoft excel does
            if(cellStyle.Rotation != 0)
                tableCellElement.Attributes["style"].Value += $"mso-rotate:{cellStyle.Rotation};";

            return string.IsNullOrEmpty(value) && cellStyleIndex == 0;
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
            short cellStyleKey = cellStyle.Index;

            if(excelStyleToClass.ContainsKey(cellStyleKey))
                return excelStyleToClass[cellStyleKey];

            String cssStyle = BuildStyle(workbook, cellStyle);
            String cssClass = htmlDocumentFacade.GetOrCreateCssClass("td", "c",
                cssStyle);
            excelStyleToClass.Add(cellStyleKey, cssClass);
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
                style.Append("font-size: " + font.FontHeightInPoints + "pt; ");
            if(font.IsItalic)
            {
                style.Append("font-style: italic; ");
            }

            var defaultFont = "sans-serif";
            style.Append($"font-family: '{font.FontName}', {defaultFont}");
        }
    }
}