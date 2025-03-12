using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BoardMemberReportGenerator
{
    class Program
    {
        // Logger configuration
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        // Constants
        private const string ProgramSponsorshipSheetName = "節目贊助";
        private const string ProgramQuotaSheetName = "節目定額";
        private const string TicketQuotaSheetName = "購劵定額";
        private const string BoardMemberIdentifierPrefix = "編號";

        static void Main(string[] args)
        {
            try
            {
                // Configure NLog
                ConfigureLogging();

                Console.WriteLine("Board Member Report Generator");
                Console.WriteLine("============================");

                // Get input parameters
                string overviewFilePath = GetUserInput("Enter the path to the overview file:", @"S:\ITD\Kei\overview_file_template\董事會成員定額紀錄 2024.xlsx");
                string outputDirectory = GetUserInput("Enter the output directory for reports:", @"S:\ITD\Kei\board_member_reports");
                string year = GetUserInput("Enter the year for reports (e.g., 2024/25):", "2024/25");

                // Ensure output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                    Logger.Info($"Created output directory: {outputDirectory}");
                }

                // Generate reports for all board members
                GenerateAllBoardMemberReports(overviewFilePath, outputDirectory, year);

                Console.WriteLine("\nReport generation completed successfully!");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "An error occurred during report generation");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
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
                Dictionary<string, int> sponsorshipEvents = GetEvents(sponsorshipSheet, sponsorshipBoardMemberCell);
                Dictionary<string, int> programQuotaEvents = GetEvents(programQuotaSheet, programQuotaBoardMemberCell);
                Dictionary<string, int> ticketQuotaEvents = GetEvents(ticketQuotaSheet, ticketQuotaBoardMemberCell);

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
            
            // Start from the row after the identifier
            int startRow = boardMemberCell.Row + 1;
            
            // Iterate through rows to find board members
            for (int rowIndex = startRow; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                
                ICell cell = row.GetCell(boardMemberCell.Col);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                
                string boardMemberName = cell.StringCellValue.Trim();
                boardMembers.Add(boardMemberName, rowIndex);
            }
            
            return boardMembers;
        }

        private static Dictionary<string, int> GetEvents(ISheet sheet, CellReference boardMemberCell)
        {
            Dictionary<string, int> events = new Dictionary<string, int>();
            
            // Get the header row
            IRow headerRow = sheet.GetRow(boardMemberCell.Row);
            if (headerRow == null) return events;
            
            // Start from column after the board member column
            int startCol = boardMemberCell.Col + 1;
            
            // Iterate through columns to find events
            for (int colIndex = startCol; colIndex < headerRow.LastCellNum; colIndex++)
            {
                ICell cell = headerRow.GetCell(colIndex);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                
                string eventName = cell.StringCellValue.Trim();
                events.Add(eventName, colIndex);
            }
            
            return events;
        }

        private static CellReference FindCellWithText(ISheet sheet, string text)
        {
            for (int rowIndex = 0; rowIndex <= Math.Min(20, sheet.LastRowNum); rowIndex++) // Limit search to first 20 rows
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                
                for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    if (cell == null || cell.CellType != CellType.String) continue;
                    
                    if (cell.StringCellValue.Contains(text))
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

            // Extract event data for this board member
            List<EventData> boardMemberEvents = ExtractBoardMemberEventData(
                sponsorshipSheet, programQuotaSheet, ticketQuotaSheet,
                boardMemberRow, 
                sponsorshipBoardMemberCell, programQuotaBoardMemberCell, ticketQuotaBoardMemberCell,
                sponsorshipEvents, programQuotaEvents, ticketQuotaEvents,
                allEventNames);

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
            IEnumerable<string> allEventNames)
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
            
            int eventIndex = 1;
            foreach (string eventName in allEventNames)
            {
                // Get sponsorship amount
                double sponsorshipAmount = 0;
                if (sponsorshipEvents.ContainsKey(eventName))
                {
                    int sponsorshipCol = sponsorshipEvents[eventName];
                    sponsorshipAmount = GetCellNumericValue(sponsorshipRow.GetCell(sponsorshipCol));
                }
                
                // Get program quota
                double programQuota = 0;
                if (programQuotaEvents.ContainsKey(eventName))
                {
                    int programQuotaCol = programQuotaEvents[eventName];
                    programQuota = GetCellNumericValue(programQuotaRow.GetCell(programQuotaCol));
                }
                
                // Get ticket quota
                double ticketQuota = 0;
                if (ticketQuotaEvents.ContainsKey(eventName))
                {
                    int ticketQuotaCol = ticketQuotaEvents[eventName];
                    ticketQuota = GetCellNumericValue(ticketQuotaRow.GetCell(ticketQuotaCol));
                }
                
                // Calculate total and receivable
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
                        Receivable = receivable
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
            // Create title row
            IRow titleRow = sheet.CreateRow(0);
            ICell titleCell = titleRow.CreateCell(0);
            titleCell.SetCellValue($"{boardMemberName}");
            
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
            titleCell.CellStyle = titleStyle;
            
            ICellStyle headerStyle = workbook.CreateCellStyle();
            IFont headerFont = workbook.CreateFont();
            headerFont.FontHeightInPoints = 12;
            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);
            yearCell.CellStyle = headerStyle;
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
            currencyStyle.DataFormat = format.GetFormat("#,##0");
            
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
                sponsorshipCell.SetCellValue(eventData.ProgramSponsorship);
                sponsorshipCell.CellStyle = currencyStyle;
                
                // Total
                ICell totalCell = row.CreateCell(3);
                totalCell.SetCellValue(eventData.Total);
                totalCell.CellStyle = currencyStyle;
                
                // Program Quota
                ICell programQuotaCell = row.CreateCell(4);
                if (eventData.ProgramQuota > 0)
                {
                    programQuotaCell.SetCellValue(eventData.ProgramQuota);
                }
                programQuotaCell.CellStyle = currencyStyle;
                
                // Ticket Quota
                ICell ticketQuotaCell = row.CreateCell(5);
                if (eventData.TicketQuota > 0)
                {
                    ticketQuotaCell.SetCellValue(eventData.TicketQuota);
                }
                ticketQuotaCell.CellStyle = currencyStyle;
                
                // Receivable
                ICell receivableCell = row.CreateCell(6);
                receivableCell.SetCellValue(eventData.Receivable);
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
            summaryStyle.DataFormat = format.GetFormat("#,##0");
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
    }
}
