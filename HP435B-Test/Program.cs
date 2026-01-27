using Ivi.Visa;
using NationalInstruments.Visa;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Grid;
using System.Drawing;
using System.Data;
using Spectre.Console;
using Color = System.Drawing.Color;
using System.Diagnostics;
using System.Data.Common;
using System.IO;

namespace HP435B_Test
{
    /// <summary>
    /// Main program class for HP435B power meter testing application.
    /// Provides automated testing capabilities including zero carryover, accuracy, and calibration factor tests.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// GPIB session for communicating with the test instrument.
        /// </summary>
        private static GpibSession gpibSession;

        /// <summary>
        /// VISA resource manager for managing instrument connections.
        /// </summary>
        private static NationalInstruments.Visa.ResourceManager resManager;

        /// <summary>
        /// GPIB address of the digital multimeter (34401A).
        /// </summary>
        private static int gpibIntAddress = 14;

        /// <summary>
        /// Full GPIB address string for instrument connection.
        /// </summary>
        private static string gpibAddress = string.Format("GPIB0::{0}::INSTR", gpibIntAddress);

        /// <summary>
        /// Semaphore used to wait for service request (SRQ) events from the instrument.
        /// </summary>
        private static readonly SemaphoreSlim srqWait = new SemaphoreSlim(0, 1);

        /// <summary>
        /// Array of test stage labels for range switch positions.
        /// </summary>
        public static readonly string[] testRangeStages = { "Fully CCW", "1 Step CW", "2 Steps CW", "3 Steps CW", "4 Steps CW", "5 Steps CW", "6 Steps CW", "7 Steps CW", "8 Steps CW", "Fully CW" };

        /// <summary>
        /// Array of test stage labels for calibration switch positions (filled at runtime).
        /// </summary>
        public static readonly string[] testCalibrationStages = new string[16];

        /// <summary>
        /// Expected value ranges for zero carryover test at each range switch position.
        /// First dimension is stage index, second dimension is [min, max] in volts.
        /// </summary>
        public static readonly double[,] zeroTestStageValues =
        {
            {-15E-3, 15E-3},
            {-17E-3, 17E-3},
            {-14E-3, 14E-3},
            {-11E-3, 11E-3},
            {-8E-3, 8E-3},
            {-5E-3, 5E-3},
            {-5E-3, 5E-3},
            {-5E-3, 5E-3},
            {-5E-3, 5E-3},
            {-5E-3, 5E-3}
        };

        /// <summary>
        /// Expected value ranges for accuracy test at each range switch position.
        /// First dimension is stage index, second dimension is [min, max] in volts.
        /// </summary>
        public static readonly double[,] accuracyTestStageValues =
{
            {975E-3, 1025E-3},
            {978E-3, 1022E-3},
            {981E-3, 1019E-3},
            {984E-3, 1016E-3},
            {987E-3, 1013E-3},
            {998E-3, 1002E-3},
            {990E-3, 1010E-3},
            {990E-3, 1010E-3},
            {990E-3, 1015E-3},
            {990E-3, 1015E-3}
        };

        /// <summary>
        /// Expected value ranges for calibration factor test at each calibration switch position.
        /// First dimension is stage index, second dimension is [min, max] in volts.
        /// </summary>
        public static readonly double[,] calibrationFactorTestStageValues =
        {
            {994E-3, 1006E-3},
            {1004E-3, 1016E-3},
            {1014E-3, 1026E-3},
            {1025E-3, 1037E-3},
            {1036E-3, 1048E-3},
            {1047E-3, 1059E-3},
            {1058E-3, 1070E-3},
            {1069E-3, 1081E-3},
            {1081E-3, 1093E-3},
            {1093E-3, 1105E-3},
            {1105E-3, 1117E-3},
            {1118E-3, 1130E-3},
            {1130E-3, 1142E-3},
            {1143E-3, 1155E-3},
            {1157E-3, 1169E-3},
            {1170E-3, 1182E-3}
        };

        /// <summary>
        /// Represents statistical data for a set of measurements.
        /// </summary>
        public struct StatisticalValues
        {
            /// <summary>
            /// Gets or sets the minimum value in the data set.
            /// </summary>
            public double Min { get; set; }

            /// <summary>
            /// Gets or sets the maximum value in the data set.
            /// </summary>
            public double Max { get; set; }

            /// <summary>
            /// Gets or sets the average (mean) value of the data set.
            /// </summary>
            public double Average { get; set; }

            /// <summary>
            /// Gets or sets the standard deviation of the data set.
            /// </summary>
            public double StdDev { get; set; }

            /// <summary>
            /// Initializes a new instance of the <see cref="StatisticalValues"/> struct.
            /// </summary>
            /// <param name="min">The minimum value.</param>
            /// <param name="max">The maximum value.</param>
            /// <param name="average">The average value.</param>
            /// <param name="stdDev">The standard deviation.</param>
            public StatisticalValues(double min, double max, double average, double stdDev)
            {
                Min = min;
                Max = max;
                Average = average;
                StdDev = stdDev;
            }

            /// <summary>
            /// Returns a string representation of the statistical values.
            /// </summary>
            /// <returns>A string containing all statistical values.</returns>
            public override string ToString()
            {
                return $"Min: {Min}, Max: {Max}, Average: {Average}, StdDev: {StdDev}";
            }

            /// <summary>
            /// Returns a string representation of the statistical values in engineering format with units.
            /// </summary>
            /// <returns>A formatted string with engineering notation and voltage units.</returns>
            public string ToEngineeringString()
            {
                return $"Min: {ToEngineeringFormat.Convert(Min, 4, "Vdc").PadRight(9)}, Max: {ToEngineeringFormat.Convert(Max, 4, "Vdc").PadRight(9)}, Average: {ToEngineeringFormat.Convert(Average, 4, "Vdc").PadRight(9)}, StdDev: {ToEngineeringFormat.Convert(StdDev, 4, "Vdc").PadRight(9)}";
            }
        }

        /// <summary>
        /// Represents a row in the test results table with two sets of range/measurement data.
        /// </summary>
        public class ResultListRow
        {
            /// <summary>
            /// Gets or sets the first range identifier.
            /// </summary>
            public string Range1 { get; set; }

            /// <summary>
            /// Gets or sets the minimum value for the first range.
            /// </summary>
            public string Min1 { get; set; }

            /// <summary>
            /// Gets or sets the actual measured value for the first range.
            /// </summary>
            public string Actual1 { get; set; }

            /// <summary>
            /// Gets or sets the maximum value for the first range.
            /// </summary>
            public string Max1 { get; set; }

            /// <summary>
            /// Gets or sets the second range identifier.
            /// </summary>
            public string Range2 { get; set; }

            /// <summary>
            /// Gets or sets the minimum value for the second range.
            /// </summary>
            public string Min2 { get; set; }

            /// <summary>
            /// Gets or sets the actual measured value for the second range.
            /// </summary>
            public string Actual2 { get; set; }

            /// <summary>
            /// Gets or sets the maximum value for the second range.
            /// </summary>
            public string Max2 { get; set; }

            /// <summary>
            /// Initializes a new instance of the <see cref="ResultListRow"/> class.
            /// </summary>
            /// <param name="range1">The first range identifier.</param>
            /// <param name="min1">The minimum value for the first range.</param>
            /// <param name="actual1">The actual measured value for the first range.</param>
            /// <param name="max1">The maximum value for the first range.</param>
            /// <param name="range2">The second range identifier.</param>
            /// <param name="min2">The minimum value for the second range.</param>
            /// <param name="actual2">The actual measured value for the second range.</param>
            /// <param name="max2">The maximum value for the second range.</param>
            public ResultListRow(string range1, string min1, string actual1, string max1, string range2, string min2, string actual2, string max2)
            {
                Range1 = range1;
                Min1 = min1;
                Actual1 = actual1;
                Max1 = max1;
                Range2 = range2;
                Min2 = min2;
                Actual2 = actual2;
                Max2 = max2;
            }

            /// <summary>
            /// Returns a string representation of the result row.
            /// </summary>
            /// <returns>A string containing all result values.</returns>
            public override string ToString()
            {
                return $"Range1 {Range1}, Min1 {Min1}, Actual1 {Actual1}, Max1 {Max1}, Range2 {Range2}, Min2 {Min2}, Actual2 {Actual2}, Max2 {Max2}";
            }
        }

        /// <summary>
        /// Main entry point for the HP435B test application.
        /// </summary>
        /// <param name="args">Command line arguments (not used).</param>
        static void Main(string[] args)
        {
            int testPoints = 100;

            StatisticalValues[] results = new StatisticalValues[16];

            try
            {
                // Setup the GPIB connection via the ResourceManager
                resManager = new NationalInstruments.Visa.ResourceManager();

                for (int i = 0, num = 100; num >= 85; i++, num--)
                {
                    testCalibrationStages[i] = num.ToString();
                }

                // Display application title and description
                AnsiConsole.Write(
                    new FigletText("HP435B Test")
                        .Centered()
                        .Color(Spectre.Console.Color.Green));

                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[bold cyan]HP435B Power Meter Test Automation Tool[/]");
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("This program performs automated performance tests on the HP435B power meter");
                AnsiConsole.MarkupLine("using an HP 34401A DMM and HP 11683A Range Calibrator.");
                AnsiConsole.WriteLine();
                
                var panel = new Panel(
                    "[bold yellow]Available Tests:[/]\n\n" +
                    "[green]  Zero Carryover[/] - Validates zero carryover across all ranges\n" +
                    "  Specification: ±0.5% of full scale when zeroed in the most sensitive range.\n\n" +
                    "[green]  Instrument Accuracy with Calibrator[/] - Tests instrumentation accuracy\n" +
                    "  Specification: ±1% of full scale on all ranges.\n\n" +
                    "[green]  Calibration Factor[/] - Tests calibration factor across 16 positions\n" +
                    "  Specification: 16-position switch normalizes meter reading to account for\n" +
                    "  calibration factor or effective efficiency (85% to 100% in 1% steps).")
                {
                    Header = new PanelHeader(" [bold white]Test Options[/] ", Justify.Center),
                    Border = BoxBorder.Rounded,
                    BorderStyle = new Style(Spectre.Console.Color.Cyan1)
                };
                
                AnsiConsole.Write(panel);
                AnsiConsole.WriteLine();

                var testChoice = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title("Select an option:")
                        .PageSize(10)
                        .AddChoices(new[] { "Settings", "Zero Carryover", "Instrument Accuracy with Calibrator", "Calibration Factor", "Exit" })
                        );

                while (testChoice != "Exit")
                {
                    if (testChoice == "Settings")
                    {
                        AnsiConsole.Clear();
                        
                        // Display settings header
                        AnsiConsole.Write(
                            new FigletText("Settings")
                                .Centered()
                                .Color(Spectre.Console.Color.Yellow));
                        
                        AnsiConsole.WriteLine();
                        
                        var settingsChoice = AnsiConsole.Prompt(
                            new SelectionPrompt<string>()
                                .Title("Select a setting to configure:")
                                .PageSize(10)
                                .AddChoices(new[] { "Set GPIB Address", "Connect to DMM", "Back to Main Menu" })
                                );
                        
                        switch (settingsChoice)
                        {
                            case "Set GPIB Address":
                                SetGPIBAddress();
                                break;
                            case "Connect to DMM":
                                if (!ConnectToDevice())
                                {
                                    AnsiConsole.MarkupLine("[yellow]Press any key to continue...[/]");
                                    Console.ReadKey(true);
                                }
                                break;
                            case "Back to Main Menu":
                                break;
                        }
                        
                        AnsiConsole.Clear();
                        
                        // Redisplay application title
                        AnsiConsole.Write(
                            new FigletText("HP435B Test")
                                .Centered()
                                .Color(Spectre.Console.Color.Green));
                        
                        AnsiConsole.WriteLine();
                        AnsiConsole.MarkupLine("[bold cyan]HP435B Power Meter Test Automation Tool[/]");
                        AnsiConsole.WriteLine();
                        
                        AnsiConsole.Write(panel);
                        AnsiConsole.WriteLine();
                    }
                    else
                    {
                        // Ensure device is connected before running tests
                        if (gpibSession == null)
                        {
                            AnsiConsole.MarkupLine("[yellow]Attempting to connect to DMM at address {0}...[/]", gpibIntAddress);
                            if (!ConnectToDevice())
                            {
                                AnsiConsole.MarkupLine("[red]Cannot run tests without a connection to the DMM.[/]");
                                AnsiConsole.MarkupLine("[yellow]Please configure GPIB address in Settings and connect to the device.[/]");
                                AnsiConsole.MarkupLine("[yellow]Press any key to continue...[/]");
                                Console.ReadKey(true);
                                
                                AnsiConsole.Clear();
                                
                                // Redisplay application title
                                AnsiConsole.Write(
                                    new FigletText("HP435B Test")
                                        .Centered()
                                        .Color(Spectre.Console.Color.Green));
                                
                                AnsiConsole.WriteLine();
                                AnsiConsole.MarkupLine("[bold cyan]HP435B Power Meter Test Automation Tool[/]");
                                AnsiConsole.WriteLine();
                                
                                AnsiConsole.Write(panel);
                                AnsiConsole.WriteLine();
                                
                                testChoice = AnsiConsole.Prompt(
                                    new SelectionPrompt<string>()
                                        .Title("Select an option:")
                                        .PageSize(10)
                                        .AddChoices(new[] { "Settings", "Zero Carryover", "Instrument Accuracy with Calibrator", "Calibration Factor", "Exit" })
                                        );
                                continue;
                            }
                        }

                        AnsiConsole.WriteLine($"DMM Details are: {QueryString("*IDN?")}");

                        SetupDMM(testPoints);

                        string reportFilename = string.Empty;

                        switch (testChoice)
                        {
                            case "Zero Carryover":
                                TestRun(results, testChoice, testRangeStages);
                                reportFilename = CreateTestReport(
                                    "Zero Carryover Test",
                                    "SPECIFICATION: ±0.5% of full scale when zeroed in the most sensitive range.",
                                    Properties.Resources.TestSetup,
                                    testRangeStages,
                                    zeroTestStageValues,
                                    results,
                                    "ZeroCarryoverTestReport",
                                    "Range Switch Position",
                                    4);
                                break;
                            case "Instrument Accuracy with Calibrator":
                                TestRun(results, testChoice, testRangeStages);
                                reportFilename = CreateTestReport(
                                    "Instrument Accuracy Test",
                                    "SPECIFICATION: ±1% of full scale on all ranges.",
                                    Properties.Resources.AccuracyTestSetup,
                                    testRangeStages,
                                    accuracyTestStageValues,
                                    results,
                                    "AccuracyTestReport",
                                    "Range Switch Position",
                                    4);
                                break;
                            case "Calibration Factor":
                                TestRun(results, testChoice, testCalibrationStages);
                                reportFilename = CreateTestReport(
                                    "Calibration Factor Test",
                                    "SPECIFICATION: 16-position switch normalizes meter reading to account for calibration factor or effective efficiency. Range 85% to 100% in 1% steps.",
                                    Properties.Resources.CalibrationTestSetup,
                                    testCalibrationStages,
                                    calibrationFactorTestStageValues,
                                    results,
                                    "CalibrationFactorTestReport",
                                    "Calibration Switch Position",
                                    4);
                                break;
                            default:
                                break;
                        }

                        SendCommand("*CLS;*RST");

                        var openReportChoice = AnsiConsole.Prompt(
                            new SelectionPrompt<string>()
                                .Title("Open the report PDF?")
                                .PageSize(10)
                                .AddChoices(new[] { "Yes", "No", })
                                );

                        if (openReportChoice == "Yes")
                            Process.Start("explorer.exe", reportFilename);

                        AnsiConsole.Clear();
                    }

                    testChoice = AnsiConsole.Prompt(
                        new SelectionPrompt<string>()
                            .Title("Select an option:")
                            .PageSize(10)
                            .AddChoices(new[] { "Settings", "Zero Carryover", "Instrument Accuracy with Calibrator", "Calibration Factor", "Exit" })
                            );
                }

                if (gpibSession != null)
                {
                    gpibSession.SendRemoteLocalCommand(GpibInstrumentRemoteLocalMode.GoToLocalDeassertRen);
                }
            }
            finally
            {
                // Unsubscribe from event before disposing to prevent resource leaks
                if (gpibSession != null)
                {
                    gpibSession.ServiceRequest -= SRQHandler;
                }
                
                gpibSession?.Dispose();
                resManager?.Dispose();
            }
        }

        /// <summary>
        /// Executes a test run across multiple stages and displays results in a live table.
        /// </summary>
        /// <param name="results">Array to store statistical results for each stage.</param>
        /// <param name="testChoice">The type of test being performed.</param>
        /// <param name="testStages">Array of stage labels for the test.</param>
        private static void TestRun(StatisticalValues[] results, string testChoice, string[] testStages)
        {
            string columnName = testChoice == "Calibration Factor" 
                ? "Calibration Switch Position" 
                : "Range Switch Position";

            // Define a Spectre table to display the test reults
            var table = new Table()
                    .AddColumn(columnName)
                    .AddColumn("Min")
                    .AddColumn("Actual")
                    .AddColumn("Max")
                    .Centered();

            table.Title = new TableTitle($"{testChoice}");

            AnsiConsole.Live(table).Start(ctx =>
            {
                for (int i = 0; i < testStages.Length; i++)
                {
                    Console.Beep(1000, 500);
                    var caption = new TableTitle($"Testing {testStages[i]} - Set DUT and hit <Enter>");
                    var captionStyle = new Style(Spectre.Console.Color.White, Spectre.Console.Color.Black, Decoration.Bold | Decoration.SlowBlink);
                    caption.Style = captionStyle;
                    table.Caption = caption;

                    ctx.Refresh();

                    DisplayTextPause(testStages[i]);

                    caption = new TableTitle($"Testing {testStages[i]} - Testing");
                    caption.Style = captionStyle;
                    table.Caption = caption;
                    ctx.Refresh();

                    results[i] = GetData(testStages[i]);

                    table.AddRow(testStages[i], ToEngineeringFormat.Convert(results[i].Min, 5, "Vdc"), ToEngineeringFormat.Convert(results[i].Average, 5, "Vdc"), ToEngineeringFormat.Convert(results[i].Max, 5, "Vdc"));

                    ctx.Refresh();
                }
            });
        }

        /// <summary>
        /// Configures the digital multimeter for testing.
        /// </summary>
        /// <param name="testPoints">Number of measurement points to acquire.</param>
        private static void SetupDMM(int testPoints)
        {
            SendCommand("*RST;*CLS");

            SendCommand("*ESE 1;*SRE 32");

            var srqSyncString = QueryString("*OPC?");

            SendCommand(":SENSe:FUNCtion \'VOLTage:DC\'");
            SendCommand(":SENSe:VOLTage:DC:RANGe 1");

            SendCommand("INPut:IMPedance:AUTO ON");

            SendCommand("TRIG:COUN " + testPoints);
        }

        /// <summary>
        /// Acquires measurement data from the instrument for a specific test stage.
        /// </summary>
        /// <param name="stage">The current test stage identifier.</param>
        /// <returns>Statistical analysis of the measurement data.</returns>
        private static StatisticalValues GetData(string stage)
        {
            SendCommand(":INIT");
            SendCommand("*OPC");

            if (!srqWait.Wait(TimeSpan.FromSeconds(30)))
            {
                throw new TimeoutException("Timeout waiting for instrument SRQ. Check instrument connection and configuration.");
            }

            var result = QueryString(":FETCh?");

            List<double> doubleList = ConvertStringToDoubleList(result);
            
            // Verify we have valid data before calculating statistics
            if (doubleList.Count == 0)
            {
                throw new InvalidOperationException($"No valid measurement data received from instrument at stage '{stage}'. Check instrument configuration and data format.");
            }
            
            return new StatisticalValues(doubleList.Min(), doubleList.Max(), doubleList.Average(), StdDev(doubleList));
        }

        /// <summary>
        /// Creates a PDF test report with results, specifications, and test setup information.
        /// </summary>
        /// <param name="reportTitle">Title of the report.</param>
        /// <param name="specification">Test specification description.</param>
        /// <param name="setupImage">Image showing the test setup.</param>
        /// <param name="stageNames">Array of stage names for the test.</param>
        /// <param name="stageLimits">Expected min/max limits for each stage.</param>
        /// <param name="results">Measured statistical results for each stage.</param>
        /// <param name="filePrefix">Prefix for the output filename.</param>
        /// <param name="switchPositionHeader">Header text for the switch position column.</param>
        /// <param name="valuePrecision">Number of significant digits for values.</param>
        /// <returns>The filename of the created PDF report.</returns>
        private static string CreateTestReport(string reportTitle, string specification, Image setupImage, string[] stageNames, double[,] stageLimits, StatisticalValues[] results, string filePrefix, string switchPositionHeader, short valuePrecision = 4)
        {
            using (PdfDocument document = new PdfDocument())
            {
                PdfPage page = document.Pages.Add();
                PdfGraphics graphics = page.Graphics;

                PdfFont titleFont = new PdfStandardFont(PdfFontFamily.Helvetica, 20, PdfFontStyle.Bold);
                PdfFont textFont = new PdfStandardFont(PdfFontFamily.Helvetica, 10, PdfFontStyle.Bold);
                PdfFont resultsFont = new PdfStandardFont(PdfFontFamily.Courier, 8, PdfFontStyle.Bold);

                PdfLayoutFormat layoutFormat = new PdfLayoutFormat
                {
                    Layout = PdfLayoutType.Paginate,
                    Break = PdfLayoutBreakType.FitPage
                };

                PdfTextElement textElement = new PdfTextElement(reportTitle, titleFont, new PdfSolidBrush(Color.Blue));
                PdfLayoutResult layoutResult = textElement.Draw(page, new RectangleF(0, 0, page.GetClientSize().Width, page.GetClientSize().Height));

                textElement = new PdfTextElement(specification, textFont, new PdfSolidBrush(Color.Black));
                layoutResult = textElement.Draw(page, new RectangleF(0, layoutResult.Bounds.Bottom + 10, page.GetClientSize().Width, page.GetClientSize().Height));

                PdfImage image = PdfImage.FromImage(setupImage);
                float scaleFactor = (float)page.GetClientSize().Width / image.Width;
                int targetHeight = (int)(image.Height * scaleFactor);
                graphics.DrawImage(image, 0, layoutResult.Bounds.Bottom + 20, page.GetClientSize().Width, targetHeight);

                textElement.Text = "Results";
                textElement.Font = new PdfStandardFont(PdfFontFamily.Helvetica, 14, PdfFontStyle.Bold);
                layoutResult = textElement.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + targetHeight + 20));

                PdfLine line = new PdfLine(new PointF(0, 0), new PointF(page.GetClientSize().Width, 0)) { Pen = PdfPens.DarkGray };
                layoutResult = line.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 5));

                PdfGrid grid = new PdfGrid();
                List<Program.ResultListRow> data = new List<Program.ResultListRow>();

                int numRows = stageNames.Length / 2;
                for (int i = 0; i < numRows; i++)
                {
                    int idx1 = i;
                    int idx2 = i + numRows;
                    data.Add(new Program.ResultListRow(
                        stageNames[idx1],
                        ToEngineeringFormat.Convert(stageLimits[idx1, 0], valuePrecision, "Vdc"),
                        ToEngineeringFormat.Convert(results[idx1].Average, valuePrecision, "Vdc"),
                        ToEngineeringFormat.Convert(stageLimits[idx1, 1], valuePrecision, "Vdc"),
                        stageNames[idx2],
                        ToEngineeringFormat.Convert(stageLimits[idx2, 0], valuePrecision, "Vdc"),
                        ToEngineeringFormat.Convert(results[idx2].Average, valuePrecision, "Vdc"),
                        ToEngineeringFormat.Convert(stageLimits[idx2, 1], valuePrecision, "Vdc")
                    ));
                }
                grid.DataSource = data;

                grid.Headers[0].Cells[0].Value = switchPositionHeader;
                grid.Headers[0].Cells[1].Value = "Results";
                grid.Headers[0].Cells[4].Value = switchPositionHeader;
                grid.Headers[0].Cells[5].Value = "Results";

                grid.Headers[0].Cells[0].RowSpan = 2;
                grid.Headers[0].Cells[0].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
                grid.Headers[0].Cells[1].ColumnSpan = 3;
                grid.Headers[0].Cells[1].StringFormat = new PdfStringFormat(PdfTextAlignment.Center);
                grid.Headers[0].Cells[4].RowSpan = 2;
                grid.Headers[0].Cells[4].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
                grid.Headers[0].Cells[5].ColumnSpan = 3;
                grid.Headers[0].Cells[5].StringFormat = new PdfStringFormat(PdfTextAlignment.Center);

                PdfGridRow[] header = grid.Headers.Add(1);
                header[1].Cells[0].Value = "";
                header[1].Cells[1].Value = "Min";
                header[1].Cells[1].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
                header[1].Cells[2].Value = "Actual";
                header[1].Cells[2].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
                header[1].Cells[3].Value = "Max";
                header[1].Cells[3].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
                header[1].Cells[4].Value = "";
                header[1].Cells[5].Value = "Min";
                header[1].Cells[5].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
                header[1].Cells[6].Value = "Actual";
                header[1].Cells[6].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
                header[1].Cells[7].Value = "Max";
                header[1].Cells[7].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);

                PdfStringFormat resultCellFormat = new PdfStringFormat { Alignment = PdfTextAlignment.Center };
                foreach (PdfGridRow gridRow in grid.Rows)
                {
                    gridRow.Cells[1].Style.StringFormat = resultCellFormat;
                    gridRow.Cells[2].Style.StringFormat = resultCellFormat;
                    gridRow.Cells[3].Style.StringFormat = resultCellFormat;
                    gridRow.Cells[5].Style.StringFormat = resultCellFormat;
                    gridRow.Cells[6].Style.StringFormat = resultCellFormat;
                    gridRow.Cells[7].Style.StringFormat = resultCellFormat;
                }

                for (int i = 0; i < grid.Rows.Count; i++)
                {
                    // First column
                    if (results[i].Average >= stageLimits[i, 0] && results[i].Average <= stageLimits[i, 1])
                    {
                        grid.Rows[i].Cells[2].Style.BackgroundBrush = new PdfSolidBrush(Color.LightGreen);
                    }
                    else
                    {
                        grid.Rows[i].Cells[2].Style.BackgroundBrush = new PdfSolidBrush(Color.Red);
                        grid.Rows[i].Cells[2].Style.TextBrush = new PdfSolidBrush(Color.White);
                    }

                    // Second column
                    int offsetValue = i + numRows;
                    if (results[offsetValue].Average >= stageLimits[offsetValue, 0] && results[offsetValue].Average <= stageLimits[offsetValue, 1])
                    {
                        grid.Rows[i].Cells[6].Style.BackgroundBrush = new PdfSolidBrush(Color.LightGreen);
                    }
                    else
                    {
                        grid.Rows[i].Cells[6].Style.BackgroundBrush = new PdfSolidBrush(Color.Red);
                        grid.Rows[i].Cells[6].Style.TextBrush = new PdfSolidBrush(Color.White);
                    }
                }
                grid.Style.CellPadding.All = 5;
                layoutResult = grid.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 20));

                textElement.Text = "Detailed Position Results";
                textElement.Font = new PdfStandardFont(PdfFontFamily.Helvetica, 14, PdfFontStyle.Bold);
                layoutResult = textElement.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 20));

                line = new PdfLine(new PointF(0, 0), new PointF(page.GetClientSize().Width, 0)) { Pen = PdfPens.DarkGray };
                layoutResult = line.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 5));

                string detailedResults = string.Empty;

                for (int i = 0; i < stageNames.Length; i++)
                {
                    detailedResults += stageNames[i].PadRight(10) + " - " + results[i].ToEngineeringString() + "\n";
                }

                PdfTextElement resultElement = new PdfTextElement(detailedResults, resultsFont, new PdfSolidBrush(Color.Black));

                layoutResult = resultElement.Draw(page, new RectangleF(0, layoutResult.Bounds.Bottom + 5, page.GetClientSize().Width, page.GetClientSize().Height), layoutFormat);

                // Include milliseconds in filename to prevent overwrites when tests run in quick succession
                var fileName = filePrefix + DateTime.Now.ToString("HH-mm-ss-fff") + ".pdf";
                
                try
                {
                    document.Save(fileName);
                    document.Close(true);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is PathTooLongException || ex is NotSupportedException)
                {
                    throw new IOException($"Failed to save PDF report '{fileName}'. Check disk space, file permissions, and path validity.", ex);
                }

                return fileName;
            }
        }

        /// <summary>
        /// Displays measurement results to the console.
        /// This is a utility method for debugging and manual result verification.
        /// </summary>
        /// <param name="title">Title for the results display.</param>
        /// <param name="doubleList">List of measured values.</param>
        private static void PrintMeasurementResults(string title, List<double> doubleList)
        {
            Console.WriteLine(title);
            Console.WriteLine($"Min Value: {ToEngineeringFormat.Convert(doubleList.Min(), 3, "Vdc")}");
            Console.WriteLine($"Max Value: {ToEngineeringFormat.Convert(doubleList.Max(), 3, "Vdc")}");
            Console.WriteLine($"Avg Value: {ToEngineeringFormat.Convert(doubleList.Average(), 3, "Vdc")}");
            Console.WriteLine($"SDev Value: {ToEngineeringFormat.Convert(StdDev(doubleList), 3, "Vdc")}");
        }

        /// <summary>
        /// Displays text on the DMM display and waits for user confirmation.
        /// </summary>
        /// <param name="text">Text to display (max 12 characters).</param>
        public static void DisplayTextPause(string text)
        {
            if (text.Length > 12)
            {
                SendCommand(":Display:Text:Data \'" + text.Substring(0, 12) + "\'");
            }
            else
            {
                SendCommand(":Display:Text:Data \'" + text + "\'");
            }

            while (Console.ReadKey(true).Key != ConsoleKey.Enter) ;

            SendCommand(":Display:Text:CLEar");
        }

        /// <summary>
        /// Calculates the standard deviation of a collection of values.
        /// </summary>
        /// <param name="values">Collection of numeric values.</param>
        /// <returns>The standard deviation of the values.</returns>
        public static double StdDev(IEnumerable<double> values)
        {
            double mean = 0.0;
            double sum = 0.0;
            double stdDev = 0.0;
            int n = 0;
            foreach (double val in values)
            {
                n++;
                double delta = val - mean;
                mean += delta / n;
                sum += delta * (val - mean);
            }
            if (1 < n)
                stdDev = Math.Sqrt(sum / (n - 1));

            return stdDev;
        }

        /// <summary>
        /// Converts a comma-separated string of numeric values to a list of doubles.
        /// </summary>
        /// <param name="input">Comma-separated string of numeric values.</param>
        /// <returns>List of double values parsed from the input string.</returns>
        public static List<double> ConvertStringToDoubleList(string input)
        {
            List<double> result = new List<double>();
            string[] values = input.Split(',');

            foreach (string value in values)
            {
                if (double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double doubleValue))
                {
                    result.Add(doubleValue);
                }
                else
                {
                    Console.WriteLine($"Conversion failed for value: {value}");
                }
            }

            return result;
        }

        /// <summary>
        /// Event handler for GPIB service request events.
        /// </summary>
        /// <param name="sender">The event sender (GPIB session).</param>
        /// <param name="e">Event arguments containing VISA event data.</param>
        public static void SRQHandler(object sender, Ivi.Visa.VisaEventArgs e)
        {
            var gbs = (GpibSession)sender;
            StatusByteFlags sb = gbs.ReadStatusByte();

            Debug.WriteLine($"SRQHandler - Status Byte: {sb}");

            gpibSession.DiscardEvents(EventType.ServiceRequest);

            SendCommand("*CLS");

            // Handle the case where multiple SRQ events fire rapidly
            // to prevent SemaphoreFullException
            try
            {
                srqWait.Release();
            }
            catch (SemaphoreFullException)
            {
                // Semaphore already released, ignore this event
                Debug.WriteLine("SRQHandler - Semaphore already at maximum count, ignoring duplicate SRQ");
            }
        }

        /// <summary>
        /// Sends a SCPI command to the instrument.
        /// </summary>
        /// <param name="command">SCPI command string to send.</param>
        static private void SendCommand(string command)
        {
            gpibSession.FormattedIO.WriteLine(command);
        }

        /// <summary>
        /// Reads a response from the instrument.
        /// </summary>
        /// <returns>The response string from the instrument.</returns>
        static private string ReadResponse()
        {
            return gpibSession.FormattedIO.ReadLine();
        }

        /// <summary>
        /// Sends a query command to the instrument and returns the response.
        /// </summary>
        /// <param name="command">SCPI query command to send.</param>
        /// <returns>The response string from the instrument.</returns>
        static private string QueryString(string command)
        {
            SendCommand(command);
            return (ReadResponse());
        }

        /// <summary>
        /// Prompts the user to set the GPIB address for the DMM.
        /// </summary>
        private static void SetGPIBAddress()
        {
            gpibIntAddress = AnsiConsole.Prompt(
                new TextPrompt<int>("Enter HP 34401A DMM GPIB address (Default is 14):")
                .DefaultValue(14)
                .Validate(n => n >= 0 && n <= 30 ? ValidationResult.Success() : ValidationResult.Error("Address must be between 0 and 30"))
                );

            // Update the GPIB address string
            gpibAddress = string.Format("GPIB0::{0}::INSTR", gpibIntAddress);

            // If we are currently connected, disconnect so we don't keep using the old instrument.
            if (gpibSession != null)
            {
                AnsiConsole.MarkupLine("[yellow]GPIB address changed while connected. Disconnecting current session.[/]");
                gpibSession.ServiceRequest -= SRQHandler;
                gpibSession.Dispose();
                gpibSession = null;
            }

            AnsiConsole.MarkupLine("[green]GPIB Address updated to: {0}[/]", gpibIntAddress);
            Thread.Sleep(1000); // Pause for a moment to let the user see the message
        }

        /// <summary>
        /// Connects to the GPIB device and initializes the session.
        /// </summary>
        /// <returns>True if connection is successful, false otherwise.</returns>
        private static bool ConnectToDevice()
        {
            if (gpibSession != null)
            {
                AnsiConsole.MarkupLine("[yellow]Warning: Already connected to a device. Disconnecting and reconnecting.[/]");
                gpibSession.ServiceRequest -= SRQHandler;
                gpibSession.Dispose();
                gpibSession = null;
                Thread.Sleep(1000); // Pause for a moment to let the user see the message
            }

            try
            {
                // Create a GPIB session for the specified address
                gpibSession = (GpibSession)resManager.Open(gpibAddress);
                gpibSession.TimeoutMilliseconds = 8000;
                gpibSession.TerminationCharacterEnabled = true;
                gpibSession.Clear();

                gpibSession.ServiceRequest += SRQHandler;

                // Test the connection by sending a simple query
                SendCommand("*IDN?");
                string response = ReadResponse();
                
                if (string.IsNullOrWhiteSpace(response))
                {
                    AnsiConsole.MarkupLine("[red]Error: Device failed to respond. Check GPIB address and device state.[/]");
                    gpibSession.ServiceRequest -= SRQHandler;
                    gpibSession.Dispose();
                    gpibSession = null;
                    Thread.Sleep(2000);
                    return false;
                }

                AnsiConsole.MarkupLine("[green]Successfully connected to device at address {0}[/]", gpibIntAddress);
                AnsiConsole.MarkupLine("[cyan]Device: {0}[/]", response.Trim());
                Thread.Sleep(2000);
                return true;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine("[red]Error: Failed to connect to GPIB device.[/]");
                AnsiConsole.MarkupLine("[red]Details: {0}[/]", ex.Message);
                
                if (gpibSession != null)
                {
                    gpibSession.ServiceRequest -= SRQHandler;
                    gpibSession.Dispose();
                    gpibSession = null;
                }
                
                Thread.Sleep(2000);
                return false;
            }
        }
    }
}
