using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BoardMemberReportGenerator
{
    class Program
    {
        // Logger configuration
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        // Configuration
        private static readonly string overviewFilePath = @"S:\ITD\Kei\overview_file_template\董事會成員定額紀錄 2024.xlsx";
        private static string OverviewFilePath;
        private static readonly string outputDirectory = @"S:\ITD\Kei\board_member_reports";
        private static readonly string year = "2024/25";

        // Constants
        private const string ProgramSponsorshipSheetName = "節目贊助";
        private const string ProgramQuotaSheetName = "節目定額";
        private const string TicketQuotaSheetName = "購券定額";
        private const string BoardMemberIdentifierPrefix = "董事會成員";

        static void Main(string[] args)
        {
            try
            {
                // Configure NLog
                ConfigureLogging();

                Console.WriteLine("Board Member Report Generator");
                Console.WriteLine("============================");

                // Ensure output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                    Logger.Info($"Created output directory: {outputDirectory}");
                }

                // Generate reports for all board members
                GenerateAllBoardMemberReports(overviewFilePath, outputDirectory, year);

                Console.WriteLine("\nReport generation completed successfully!");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "An error occurred during report generation");
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void ConfigureLogging()
        {
            var config = new NLog.Config.LoggingConfiguration();
            
            // Targets
            var logfile = new NLog.Targets.FileTarget("logfile") 
            { 
                FileName = "board_member_reports.log",
                Layout = "${longdate} ${level:uppercase=true} ${logger} - ${message} ${exception:format=tostring}"
            };
            
            var logconsole = new NLog.Targets.ConsoleTarget("logconsole")
            {
                Layout = "${date:format=HH\\:mm\\:ss} ${level:uppercase=true} - ${message} ${exception:format=message}"
            };
            
            // Rules
            config.AddRule(LogLevel.Info, LogLevel.Fatal, logconsole);
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);
            
            // Apply config
            LogManager.Configuration = config;
        }

        private static string GetUserInput(string prompt, string defaultValue)
        {
            Console.WriteLine(prompt);
            Console.Write($"[{defaultValue}]: ");
            string input = Console.ReadLine();
            return string.IsNullOrWhiteSpace(input) ? defaultValue : input;
        }

        private static void GenerateAllBoardMemberReports(string overviewFilePath, string outputDirectory, string year)
        {
            // Store the overview file path as a static variable for use in other methods
            OverviewFilePath = overviewFilePath;
            Logger.Info($"Starting report generation from overview file: {overviewFilePath}");

            // Load the overview file
            using (FileStream fs = new FileStream(overviewFilePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook overviewWorkbook = WorkbookFactory.Create(fs);
                
                // Get all sheets
                ISheet sponsorshipSheet = overviewWorkbook.GetSheet(ProgramSponsorshipSheetName);
                ISheet programQuotaSheet = overviewWorkbook.GetSheet(ProgramQuotaSheetName);
                ISheet ticketQuotaSheet = overviewWorkbook.GetSheet(TicketQuotaSheetName);

                if (sponsorshipSheet == null)
                {
                    throw new Exception($"Sheet '{ProgramSponsorshipSheetName}' not found in overview file");
                }
                if (programQuotaSheet == null)
                {
                    throw new Exception($"Sheet '{ProgramQuotaSheetName}' not found in overview file");
                }
                if (ticketQuotaSheet == null)
                {
                    throw new Exception($"Sheet '{TicketQuotaSheetName}' not found in overview file");
                }

                // Find board member identifier cells in each sheet
                CellReference sponsorshipBoardMemberCell = FindCellWithText(sponsorshipSheet, BoardMemberIdentifierPrefix);
                CellReference programQuotaBoardMemberCell = FindCellWithText(programQuotaSheet, BoardMemberIdentifierPrefix);
                CellReference ticketQuotaBoardMemberCell = FindCellWithText(ticketQuotaSheet, BoardMemberIdentifierPrefix);

                if (sponsorshipBoardMemberCell == null || programQuotaBoardMemberCell == null || ticketQuotaBoardMemberCell == null)
                {
                    throw new Exception($"Board member identifier cell with text '{BoardMemberIdentifierPrefix}' not found in one or more sheets");
                }

                // Get all board members from the sponsorship sheet
                Dictionary<string, int> boardMembers = GetBoardMembers(sponsorshipSheet, sponsorshipBoardMemberCell);
                Logger.Info($"Found {boardMembers.Count} board members in overview file");

                // Get all events from each sheet
                Dictionary<string, int> sponsorshipEvents = GetExistingEvents(sponsorshipSheet, sponsorshipBoardMemberCell, 5);
                Dictionary<string, int> programQuotaEvents = GetExistingEvents(programQuotaSheet, programQuotaBoardMemberCell, 5);
                Dictionary<string, int> ticketQuotaEvents = GetExistingEvents(ticketQuotaSheet, ticketQuotaBoardMemberCell, 6);

                // Combine all unique events
                HashSet<string> allEventNames = new HashSet<string>();
                foreach (var eventName in sponsorshipEvents.Keys) allEventNames.Add(eventName);
                foreach (var eventName in programQuotaEvents.Keys) allEventNames.Add(eventName);
                foreach (var eventName in ticketQuotaEvents.Keys) allEventNames.Add(eventName);

                Logger.Info($"Found {allEventNames.Count} unique events across all sheets");

                // Generate report for each board member
                int successCount = 0;
                foreach (var boardMember in boardMembers)
                {
                    try
                    {
                        ExportBoardMemberReport(
                            sponsorshipSheet, programQuotaSheet, ticketQuotaSheet,
                            sponsorshipBoardMemberCell, programQuotaBoardMemberCell, ticketQuotaBoardMemberCell,
                            sponsorshipEvents, programQuotaEvents, ticketQuotaEvents,
                            outputDirectory, boardMember.Key, boardMember.Value, allEventNames, year);
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error generating report for board member '{boardMember.Key}'");
                    }
                }

                Logger.Info($"Successfully generated {successCount} of {boardMembers.Count} board member reports");
                Console.WriteLine($"Generated {successCount} of {boardMembers.Count} board member reports");
            }
        }
        
        private static Dictionary<string, int> GetBoardMembers(ISheet sheet, CellReference boardMemberCell)
        {
            Dictionary<string, int> boardMembers = new Dictionary<string, int>();
            int expectedMemberCount = 20; // Expected number of board members
            int foundMemberCount = 0;
            
            int row = boardMemberCell.Row + 1;
            int maxRowsToCheck = row + 20; // Check up to 20 rows to find board members
            
            while (row < maxRowsToCheck)
            {
                IRow currentRow = sheet.GetRow(row);
                if (currentRow == null)
                    break;
                
                ICell cell = currentRow.GetCell(boardMemberCell.Col);
                if (cell == null)
                {
                    row++;
                    continue;
                }
                
                string cellValue = cell.ToString().Trim();
                if (string.IsNullOrEmpty(cellValue))
                {
                    row++;
                    continue;
                }
                
                // Check if the cell value matches the expected format: number followed by a dot
                if (Regex.IsMatch(cellValue, @"^\d+\."))
                {
                    boardMembers.Add(cellValue, row);
                    foundMemberCount++;
                    Logger.Debug($"Found board member: {cellValue} at row {row + 1}");
                }
                else
                {
                    // If we've already found some members but this one doesn't match the pattern,
                    // it might indicate we've reached the end of the member list
                    if (foundMemberCount > 0)
                    {
                        Logger.Debug($"Possible end of board member list at row {row + 1}: '{cellValue}'");
                        // Don't break immediately, as there might be valid members after this one
                    }
                }
                
                row++;
            }
            
            // Log warning if we found significantly fewer or more members than expected
            if (foundMemberCount < expectedMemberCount - 5)
            {
                Logger.Warn($"Found only {foundMemberCount} board members, expected around {expectedMemberCount}");
            }
            else if (foundMemberCount > expectedMemberCount + 5)
            {
                Logger.Warn($"Found {foundMemberCount} board members, which is more than the expected {expectedMemberCount}");
            }
            else
            {
                Logger.Info($"Found {foundMemberCount} board members");
            }
            
            return boardMembers;
        }

        private static Dictionary<string, int> GetExistingEvents(ISheet sheet, CellReference boardMemberCell, int startColumn)
        {
            Dictionary<string, int> events = new Dictionary<string, int>();
            
            IRow headerRow = sheet.GetRow(boardMemberCell.Row);
            if (headerRow == null)
                return events;
            
            int column = startColumn;
            while (true)
            {
                ICell cell = headerRow.GetCell(column);
                if (cell == null || string.IsNullOrEmpty(cell.ToString()))
                    break;
                
                string eventName = cell.ToString();
                events.Add(eventName, column);
                column++;
            }
            
            return events;
        }

        private static CellReference FindCellWithText(ISheet sheet, string text)
        {
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                
                for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    if (cell != null && cell.ToString().Replace("\r\n", string.Empty).Replace("\n", string.Empty) == text)
                    {
                        return new CellReference(rowIndex, colIndex);
                    }
                }
            }
            
            return null;
        }

        private static void ExportBoardMemberReport(
            ISheet sponsorshipSheet, ISheet programQuotaSheet, ISheet ticketQuotaSheet,
            CellReference sponsorshipBoardMemberCell, CellReference programQuotaBoardMemberCell, CellReference ticketQuotaBoardMemberCell,
            Dictionary<string, int> sponsorshipEvents, Dictionary<string, int> programQuotaEvents, Dictionary<string, int> ticketQuotaEvents,
            string outputDirectory, string boardMemberName, int boardMemberRow, IEnumerable<string> allEventNames, string year)
        {
            Logger.Info($"Generating report for board member: {boardMemberName}");
            Console.WriteLine($"Generating report for: {boardMemberName}");

            // Create a new workbook for the board member report
            IWorkbook reportWorkbook = new XSSFWorkbook();
            ISheet reportSheet = reportWorkbook.CreateSheet("Report");

            /// Extract event data for this board member, passing the overview file path
            List<EventData> boardMemberEvents = ExtractBoardMemberEventData(
                sponsorshipSheet, programQuotaSheet, ticketQuotaSheet,
                boardMemberRow, 
                sponsorshipBoardMemberCell, programQuotaBoardMemberCell, ticketQuotaBoardMemberCell,
                sponsorshipEvents, programQuotaEvents, ticketQuotaEvents,
                allEventNames,
                OverviewFilePath);  // Add overview file path here

            // Create the report
            CreateReportHeader(reportWorkbook, reportSheet, boardMemberName, year);
            CreateReportColumnHeaders(reportWorkbook, reportSheet);
            PopulateEventData(reportWorkbook, reportSheet, boardMemberEvents);
            CreateSummaryRow(reportWorkbook, reportSheet, boardMemberEvents.Count);
            ApplyReportFormatting(reportWorkbook, reportSheet, boardMemberEvents.Count);

            // Save the report
            string sanitizedName = SanitizeFileName(boardMemberName);
            string outputPath = Path.Combine(outputDirectory, 
                $"{year.Replace("/", "")}籌款活動應收款_{sanitizedName}.xlsx");
            
            using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                reportWorkbook.Write(fs);
            }
            
            Logger.Info($"Created board member report at {outputPath}");
        }

        private static string SanitizeFileName(string fileName)
        {
            // Remove invalid file name characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));
        }

        private static List<EventData> ExtractBoardMemberEventData(
            ISheet sponsorshipSheet, ISheet programQuotaSheet, ISheet ticketQuotaSheet,
            int boardMemberRow,
            CellReference sponsorshipBoardMemberCell, CellReference programQuotaBoardMemberCell, CellReference ticketQuotaBoardMemberCell,
            Dictionary<string, int> sponsorshipEvents, Dictionary<string, int> programQuotaEvents, Dictionary<string, int> ticketQuotaEvents,
            IEnumerable<string> allEventNames,
            string overviewFilePath)
        {
            List<EventData> eventDataList = new List<EventData>();
            
            IRow sponsorshipRow = sponsorshipSheet.GetRow(boardMemberRow);
            IRow programQuotaRow = programQuotaSheet.GetRow(boardMemberRow);
            IRow ticketQuotaRow = ticketQuotaSheet.GetRow(boardMemberRow);
            
            if (sponsorshipRow == null || programQuotaRow == null || ticketQuotaRow == null)
            {
                Logger.Warn($"One or more data rows not found for board member at row {boardMemberRow}");
                return eventDataList;
            }
            
            // Get the full path of the overview file for external references
            string fullOverviewPath = Path.GetFullPath(overviewFilePath);
            string overviewDir = Path.GetDirectoryName(fullOverviewPath).Replace("\\", "/");
            string overviewFileName = Path.GetFileName(fullOverviewPath);
            
            int eventIndex = 1;
            foreach (string eventName in allEventNames)
            {
                double sponsorshipAmount = 0;
                double programQuota = 0;
                double ticketQuota = 0;
                
                string sponsorshipFormula = null;
                string programQuotaFormula = null;
                string ticketQuotaFormula = null;
                
                // Get sponsorship amount and formula
                if (sponsorshipEvents.ContainsKey(eventName))
                {
                    int sponsorshipCol = sponsorshipEvents[eventName];
                    ICell sponsorshipCell = sponsorshipRow.GetCell(sponsorshipCol);
                    
                    if (sponsorshipCell != null)
                    {
                        sponsorshipAmount = GetCellNumericValue(sponsorshipCell);
                        
                        // Create external reference formula to sponsorship sheet
                        sponsorshipFormula = $"'{overviewDir}/[{overviewFileName}]{ProgramSponsorshipSheetName}'!" + 
                            $"{CellReference.ConvertNumToColString(sponsorshipCol)}{boardMemberRow + 1}";
                    }
                }
                
                // Get program quota and formula
                if (programQuotaEvents.ContainsKey(eventName))
                {
                    int programQuotaCol = programQuotaEvents[eventName];
                    ICell programQuotaCell = programQuotaRow.GetCell(programQuotaCol);
                    
                    if (programQuotaCell != null)
                    {
                        programQuota = GetCellNumericValue(programQuotaCell);
                        
                        // Create external reference formula to program quota sheet
                        programQuotaFormula = $"'{overviewDir}/[{overviewFileName}]{ProgramQuotaSheetName}'!" + 
                            $"{CellReference.ConvertNumToColString(programQuotaCol)}{boardMemberRow + 1}";
                    }
                }
                
                // Get ticket quota and formula
                if (ticketQuotaEvents.ContainsKey(eventName))
                {
                    int ticketQuotaCol = ticketQuotaEvents[eventName];
                    ICell ticketQuotaCell = ticketQuotaRow.GetCell(ticketQuotaCol);
                    
                    if (ticketQuotaCell != null)
                    {
                        ticketQuota = GetCellNumericValue(ticketQuotaCell);
                        
                        // Create external reference formula to ticket quota sheet
                        ticketQuotaFormula = $"'{overviewDir}/[{overviewFileName}]{TicketQuotaSheetName}'!" + 
                            $"{CellReference.ConvertNumToColString(ticketQuotaCol)}{boardMemberRow + 1}";
                    }
                }
                
                // Calculate total and receivable based on the current values
                // (These will be calculated by Excel formulas when the report is opened)
                double total = sponsorshipAmount;
                double receivable = total - ticketQuota;
                
                // Only include events with non-zero values
                if (sponsorshipAmount > 0 || programQuota > 0 || ticketQuota > 0)
                {
                    EventData eventData = new EventData
                    {
                        Index = eventIndex++,
                        Name = eventName,
                        ProgramSponsorship = sponsorshipAmount,
                        Total = total,
                        ProgramQuota = programQuota,
                        TicketQuota = ticketQuota,
                        Receivable = receivable,
                        // Store formulas
                        ProgramSponsorshipFormula = sponsorshipFormula,
                        ProgramQuotaFormula = programQuotaFormula,
                        TicketQuotaFormula = ticketQuotaFormula
                    };
                    
                    eventDataList.Add(eventData);
                }
            }
            
            return eventDataList;
        }

        private static double GetCellNumericValue(ICell cell)
        {
            if (cell == null) return 0;
            
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        return cell.NumericCellValue;
                    case CellType.Formula:
                        try
                        {
                            return cell.NumericCellValue;
                        }
                        catch
                        {
                            // If formula evaluation fails, try to parse the cached formula result
                            string formulaResult = cell.ToString();
                            if (double.TryParse(formulaResult, out double result))
                                return result;
                            return 0;
                        }
                    case CellType.String:
                        if (double.TryParse(cell.StringCellValue, out double value))
                            return value;
                        return 0;
                    default:
                        return 0;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"Error getting numeric value from cell: {ex.Message}");
                return 0;
            }
        }

        private static void CreateReportHeader(
            IWorkbook workbook,
            ISheet sheet,
            string boardMemberName,
            string year)
        {
            // Remove the number and dot prefix from board member name (e.g., "1.龐董晶怡主席" -> "龐董晶怡主席")
            string cleanedName = boardMemberName;
            if (Regex.IsMatch(boardMemberName, @"^\d+\."))
            {
                // Find first dot position and take everything after it
                int dotPosition = boardMemberName.IndexOf('.');
                if (dotPosition >= 0)
                {
                    cleanedName = boardMemberName.Substring(dotPosition + 1).Trim();
                }
            }
            
            // Create title row
            IRow titleRow = sheet.CreateRow(0);
            ICell titleCell = titleRow.CreateCell(0);
            titleCell.SetCellValue(cleanedName);
            
            // Create year row
            IRow yearRow = sheet.CreateRow(1);
            ICell yearCell = yearRow.CreateCell(0);
            yearCell.SetCellValue($"{year} 籌款活動應收款頂");
            
            // Create date row
            IRow dateRow = sheet.CreateRow(3);
            ICell dateCell = dateRow.CreateCell(0);
            dateCell.SetCellValue($"製作日期：{DateTime.Now:dd/M/yyyy}");
            
            // Apply styles
            ICellStyle titleStyle = workbook.CreateCellStyle();
            IFont titleFont = workbook.CreateFont();
            titleFont.FontHeightInPoints = 14;
            titleFont.IsBold = true;
            titleStyle.SetFont(titleFont);
            titleStyle.Alignment = HorizontalAlignment.Center; // Center align
            titleCell.CellStyle = titleStyle;
            
            ICellStyle headerStyle = workbook.CreateCellStyle();
            IFont headerFont = workbook.CreateFont();
            headerFont.FontHeightInPoints = 12;
            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);
            headerStyle.Alignment = HorizontalAlignment.Center; // Center align
            yearCell.CellStyle = headerStyle;
            
            // Merge cells for title row (columns A-G, which are 0-6 in 0-based index)
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 6));
            
            // Merge cells for year row (columns A-G, which are 0-6 in 0-based index)
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 6));
        }

        private static void CreateReportColumnHeaders(IWorkbook workbook, ISheet sheet)
        {
            // Create header row
            IRow headerRow = sheet.CreateRow(4);
            
            // Create headers
            string[] headers = new string[] 
            { 
                "籌款項目", "", "節目贊助", "總額", "節目定額", "購劵定額", "應收款" 
            };
            
            for (int i = 0; i < headers.Length; i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(headers[i]);
            }
            
            // Apply styles
            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.Alignment = HorizontalAlignment.Center;
            IFont headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);
            
            for (int i = 0; i < headers.Length; i++)
            {
                if (!string.IsNullOrEmpty(headers[i]))
                {
                    headerRow.GetCell(i).CellStyle = headerStyle;
                }
            }
        }

        private static void PopulateEventData(
            IWorkbook workbook,
            ISheet sheet,
            List<EventData> events)
        {
            // Create currency style
            ICellStyle currencyStyle = workbook.CreateCellStyle();
            currencyStyle.Alignment = HorizontalAlignment.Center;
            IDataFormat format = workbook.CreateDataFormat();
            currencyStyle.DataFormat = format.GetFormat("#,##0.00");
            
            // Populate event data
            for (int i = 0; i < events.Count; i++)
            {
                EventData eventData = events[i];
                IRow row = sheet.CreateRow(5 + i);
                
                // Index
                ICell indexCell = row.CreateCell(0);
                indexCell.SetCellValue(eventData.Index);
                indexCell.CellStyle = currencyStyle;
                
                // Name
                ICell nameCell = row.CreateCell(1);
                nameCell.SetCellValue(eventData.Name);
                
                // Program Sponsorship
                ICell sponsorshipCell = row.CreateCell(2);
                if (eventData.ProgramSponsorshipFormula != null)
                {
                    // Use formula reference to overview file
                    sponsorshipCell.SetCellFormula(eventData.ProgramSponsorshipFormula);
                }
                else
                {
                    sponsorshipCell.SetCellValue(eventData.ProgramSponsorship);
                }
                sponsorshipCell.CellStyle = currencyStyle;
                
                // Total - always calculated from sponsorship
                ICell totalCell = row.CreateCell(3);
                totalCell.SetCellFormula($"C{6+i}"); // Reference to sponsorship cell
                totalCell.CellStyle = currencyStyle;
                
                // Program Quota
                ICell programQuotaCell = row.CreateCell(4);
                if (eventData.ProgramQuotaFormula != null)
                {
                    // Use formula reference to overview file
                    programQuotaCell.SetCellFormula(eventData.ProgramQuotaFormula);
                }
                else if (eventData.ProgramQuota > 0)
                {
                    programQuotaCell.SetCellValue(eventData.ProgramQuota);
                }
                programQuotaCell.CellStyle = currencyStyle;
                
                // Ticket Quota
                ICell ticketQuotaCell = row.CreateCell(5);
                if (eventData.TicketQuotaFormula != null)
                {
                    // Use formula reference to overview file
                    ticketQuotaCell.SetCellFormula(eventData.TicketQuotaFormula);
                }
                else if (eventData.TicketQuota > 0)
                {
                    ticketQuotaCell.SetCellValue(eventData.TicketQuota);
                }
                ticketQuotaCell.CellStyle = currencyStyle;
                
                // Receivable - always calculated as total minus ticket quota
                ICell receivableCell = row.CreateCell(6);
                receivableCell.SetCellFormula($"D{6+i}-F{6+i}"); // Total - Ticket Quota
                receivableCell.CellStyle = currencyStyle;
            }
        }

        private static void CreateSummaryRow(
            IWorkbook workbook,
            ISheet sheet,
            int eventCount)
        {
            // Create summary row
            IRow summaryRow = sheet.CreateRow(5 + eventCount);
            
            // Create summary cells
            for (int i = 2; i <= 6; i++)
            {
                ICell cell = summaryRow.CreateCell(i);
                
                // Create SUM formula
                string colLetter = CellReference.ConvertNumToColString(i);
                string formula = $"SUM({colLetter}6:{colLetter}{5 + eventCount})";
                cell.SetCellFormula(formula);
            }
            
            // Apply styles
            ICellStyle summaryStyle = workbook.CreateCellStyle();
            summaryStyle.Alignment = HorizontalAlignment.Center;
            IDataFormat format = workbook.CreateDataFormat();
            summaryStyle.DataFormat = format.GetFormat("#,##0.00");
            IFont boldFont = workbook.CreateFont();
            boldFont.IsBold = true;
            summaryStyle.SetFont(boldFont);
            
            for (int i = 2; i <= 6; i++)
            {
                summaryRow.GetCell(i).CellStyle = summaryStyle;
            }
        }

        private static void ApplyReportFormatting(
            IWorkbook workbook,
            ISheet sheet,
            int eventCount)
        {
            // Set column widths
            sheet.SetColumnWidth(0, 10 * 256); // Index
            sheet.SetColumnWidth(1, 20 * 256); // Event name
            sheet.SetColumnWidth(2, 15 * 256); // Program sponsorship
            sheet.SetColumnWidth(3, 15 * 256); // Total
            sheet.SetColumnWidth(4, 15 * 256); // Program quota
            sheet.SetColumnWidth(5, 15 * 256); // Ticket quota
            sheet.SetColumnWidth(6, 15 * 256); // Receivable
            
            // Create border style
            ICellStyle borderStyle = workbook.CreateCellStyle();
            borderStyle.BorderBottom = BorderStyle.Thin;
            borderStyle.BorderTop = BorderStyle.Thin;
            borderStyle.BorderLeft = BorderStyle.Thin;
            borderStyle.BorderRight = BorderStyle.Thin;
            
            // Apply borders to data cells
            for (int rowIndex = 4; rowIndex <= 5 + eventCount; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                
                for (int colIndex = 0; colIndex <= 6; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    if (cell == null) continue;
                    
                    // Clone existing style and add borders
                    ICellStyle newStyle = workbook.CreateCellStyle();
                    if (cell.CellStyle != null)
                    {
                        newStyle.CloneStyleFrom(cell.CellStyle);
                    }
                    
                    newStyle.BorderBottom = BorderStyle.Thin;
                    newStyle.BorderTop = BorderStyle.Thin;
                    newStyle.BorderLeft = BorderStyle.Thin;
                    newStyle.BorderRight = BorderStyle.Thin;
                    
                    cell.CellStyle = newStyle;
                }
            }
        }
    }

    public class EventData
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public double ProgramSponsorship { get; set; }
        public double Total { get; set; }
        public double ProgramQuota { get; set; }
        public double TicketQuota { get; set; }
        public double Receivable { get; set; }
        
        // Formula references to overview file
        public string ProgramSponsorshipFormula { get; set; }
        public string ProgramQuotaFormula { get; set; }
        public string TicketQuotaFormula { get; set; }
    }
}
