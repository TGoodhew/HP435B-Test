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

namespace HP435B_Test
{
    internal class Program
    {
        public static GpibSession gpibSession;
        public static NationalInstruments.Visa.ResourceManager resManager;
        public static int gpibIntAddress = 16;
        public static string gpibAddress = string.Format("GPIB0::{0}::INSTR", gpibIntAddress);
        public static SemaphoreSlim srqWait = new SemaphoreSlim(0, 1); // use a semaphore to wait for the SRQ events

        public static readonly string[] testStages = { "Fully CCW", "1 Step CW", "2 Steps CW", "3 Steps CW", "4 Steps CW", "5 Steps CW", "6 Steps CW", "7 Steps CW", "8 Steps CW", "Fully CW" };
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


        public struct StatisticalValues
        {
            public double Min { get; set; }
            public double Max { get; set; }
            public double Average { get; set; }
            public double StdDev { get; set; }

            public StatisticalValues(double min, double max, double average, double stdDev)
            {
                Min = min;
                Max = max;
                Average = average;
                StdDev = stdDev;
            }

            public override string ToString()
            {
                return $"Min: {Min}, Max: {Max}, Average: {Average}, StdDev: {StdDev}";
            }

            public string ToEngineeringString()
            {
                return $"Min: {ToEngineeringFormat.Convert(Min,3,"Vdc").PadRight(9)}, Max: {ToEngineeringFormat.Convert(Max, 3, "Vdc").PadRight(9)}, Average: {ToEngineeringFormat.Convert(Average, 3, "Vdc").PadRight(9)}, StdDev: {ToEngineeringFormat.Convert(StdDev, 3, "Vdc").PadRight(9)}";
            }

        }

        public class ResultListRow
        {
            public string Range1 { get; set; }
            public string Min1 { get; set; }
            public string Actual1 { get; set; }
            public string Max1 { get; set; }
            public string Range2 { get; set; }
            public string Min2 { get; set; }
            public string Actual2 { get; set; }
            public string Max2 { get; set; }

            // Example constructor for initialization
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

            public override string ToString()
            {
                return $"Range1 {Range1}, Min1 {Min1}, Actual1 {Actual1}, Max1 {Max1}, Range2 {Range2}, Min2 {Min2}, Actual2 {Actual2}, Max2 {Max2}";
            }
        }

        static void Main(string[] args)
        {
            int testPoints = 100; // Number of test points to take

            StatisticalValues[] results = new StatisticalValues[10];

            // Setup the GPIB connection via the ResourceManager
            resManager = new NationalInstruments.Visa.ResourceManager();

            // Create a GPIB session for the specified address
            gpibSession = (GpibSession)resManager.Open(gpibAddress);
            gpibSession.TimeoutMilliseconds = 8000; // Set the timeout to be 4s
            gpibSession.TerminationCharacterEnabled = true;
            gpibSession.Clear(); // Clear the session

            gpibSession.ServiceRequest += SRQHandler;

            // Ask for the user's favorite fruit
            var TestChoice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Select the test to run?")
                    .PageSize(10)
                    .AddChoices(new[] {
            "Zero Carryover", "Instrument Accuracy with Calibrator",
                    }));

            // Echo the fruit back to the terminal
            AnsiConsole.WriteLine($"DMM Details are: {QueryString("*IDN?")}");


            // Reset the DMM
            SendCommand("*RST");

            // Set the DMM to generate an SRQ when done
            SendCommand(":status: preset");
            SendCommand("*CLS");
            SendCommand(":status:measurement:enable 512");
            SendCommand("*sre 1");

            // Set the DMM to DC Voltage mode and the range to 200mV
            SendCommand(":SENSe:FUNCtion \'VOLTage:DC\'");
            SendCommand(":SENSe:VOLTage:DC:RANGe 0.2");

            // Set the DMM to trigger for specified measurements
            SendCommand(":sample:count " + testPoints);
            SendCommand(":trigger:source bus");
            SendCommand(":TRACe:POINts " + testPoints);
            SendCommand(":TRACe:FEED SENSe1");

            // Iterate through the test stages
            for (int i = 0; i < testStages.Length; i++)
            {
                // Prompt the user to set the range switch position
                DisplayTextPause(testStages[i]);

                // Get the data for the current stage
                results[i] = GetData(testStages[i]);
            }

            // Create a PDF report with the results
            CreeateAccuracyTestReport(results);

            // Reset the intrument and return to local control
            SendCommand("*RST");
            gpibSession.SendRemoteLocalCommand(GpibInstrumentRemoteLocalMode.GoToLocalDeassertRen);
        }

        private static StatisticalValues GetData(string stage)
        {
            // Clear the trace and status bytes for next measurement
            var result = QueryString(":status:measurement?");
            SendCommand(":TRACe:CLEar");
            SendCommand(":TRACe:FEED:CONTrol NEXT");
            SendCommand(":INIT");

            // Trigger the reading
            gpibSession.AssertTrigger();

            // Wait for the data to be available
            Console.WriteLine("Waiting for data");
            srqWait.Wait();

            Console.WriteLine("Retrieving data");
            result = QueryString(":TRACe:DATA?");

            // Convert the string to a list of doubles
            List<double> doubleList = ConvertStringToDoubleList(result);

            // Print the results
            PrintMeasurementResults("Zero Carryover - " + stage, doubleList);

            return new StatisticalValues(doubleList.Min(), doubleList.Max(), doubleList.Average(), StdDev(doubleList));

        }
        private static void CreeateZeroTestReport(StatisticalValues[] results)
        {
            // Based off the SyncFusion example code for PDF generation

            // Create a new PDF document.
            PdfDocument document = new PdfDocument();

            // Add a page to the document.
            PdfPage page = document.Pages.Add();

            // Create PDF graphics for the page.
            PdfGraphics graphics = page.Graphics;

            // Initialize our fonts font.
            PdfFont titleFont = new PdfStandardFont(PdfFontFamily.Helvetica, 20, PdfFontStyle.Bold);
            PdfFont textFont = new PdfStandardFont(PdfFontFamily.Helvetica, 10, PdfFontStyle.Bold);
            PdfFont resultsFont = new PdfStandardFont(PdfFontFamily.Courier, 10, PdfFontStyle.Bold);

            // Load the paragraph text into PdfTextElement with initialized standard font.
            PdfTextElement textElement = new PdfTextElement("Zero Carryover Test", titleFont, new PdfSolidBrush(Color.Blue));

            // Draw the paragraph text on page and maintain the position in PdfLayoutResult.
            PdfLayoutResult layoutResult = textElement.Draw(page, new RectangleF(0, 0, page.GetClientSize().Width, page.GetClientSize().Height));

            // Load the paragraph text into PdfTextElement with initialized standard font.
            textElement = new PdfTextElement("SPECIFICATION: ±0.5% of full scale when zeroed in the most sensitive range.", textFont, new PdfSolidBrush(Color.Black));

            // Draw the paragraph text on page and maintain the position in PdfLayoutResult.
            layoutResult = textElement.Draw(page, new RectangleF(0, layoutResult.Bounds.Bottom + 10, page.GetClientSize().Width, page.GetClientSize().Height));

            PdfImage image = PdfImage.FromImage(Properties.Resources.TestSetup);

            // Calculate the scale factor based on the target width
            float scaleFactor = (float)page.GetClientSize().Width / image.Width;

            // Compute the new height while maintaining the aspect ratio
            int targetHeight = (int)(image.Height * scaleFactor);

            graphics.DrawImage(image, 0, layoutResult.Bounds.Bottom + 20, page.GetClientSize().Width, targetHeight);

            // Assign header text to PdfTextElement.
            textElement.Text = "Results";

            // Assign standard font to PdfTextElement.
            textElement.Font = new PdfStandardFont(PdfFontFamily.Helvetica, 14, PdfFontStyle.Bold);

            // Draw the header text on page, below the paragraph text with a height gap of 20 and maintain the position in PdfLayoutResult.
            layoutResult = textElement.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + targetHeight + 20));

            // Initialize PdfLine with start point and end point for drawing the line.
            PdfLine line = new PdfLine(new PointF(0, 0), new PointF(page.GetClientSize().Width, 0))
            {
                Pen = PdfPens.DarkGray
            };

            // Draw the line on page, below the header text with a height gap of 5 and maintain the position in PdfLayoutResult.
            layoutResult = line.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 5));

            // Initialize PdfGrid for drawing the table.
            PdfGrid grid = new PdfGrid();

            List<ResultListRow> data = new List<ResultListRow>();
            data.Add(new ResultListRow("Fully CCW", "-15", ToEngineeringFormat.Convert(results[0].Average, 3, "Vdc"), "+15", "5 Steps CW", "-5", ToEngineeringFormat.Convert(results[5].Average, 3, "Vdc"), "+5"));
            data.Add(new ResultListRow("1 Step CW", "-17", ToEngineeringFormat.Convert(results[1].Average, 3, "Vdc"), "+17", "6 Steps CW", "-5", ToEngineeringFormat.Convert(results[6].Average, 3, "Vdc"), "+5"));
            data.Add(new ResultListRow("2 Steps CW", "-14", ToEngineeringFormat.Convert(results[2].Average, 3, "Vdc"), "+14", "7 Steps CW", "-5", ToEngineeringFormat.Convert(results[7].Average, 3, "Vdc"), "+5"));
            data.Add(new ResultListRow("3 Steps CW", "-11", ToEngineeringFormat.Convert(results[3].Average, 3, "Vdc"), "+11", "8 Steps CW", "-5", ToEngineeringFormat.Convert(results[8].Average, 3, "Vdc"), "+5"));
            data.Add(new ResultListRow("4 Steps CW", "-8", ToEngineeringFormat.Convert(results[4].Average, 3, "Vdc"), "+8", "Fully CW", "-5", ToEngineeringFormat.Convert(results[9].Average, 3, "Vdc"), "+5"));

            // Add list to IEnumerable.
            IEnumerable<object> dataTable = data;

            // Assign the DataTable as data source to grid.
            grid.DataSource = dataTable;

            //Using the Header collection
            grid.Headers[0].Cells[0].Value = "Range Switch Position";
            grid.Headers[0].Cells[1].Value = "Results";
            grid.Headers[0].Cells[4].Value = "Range Switch Position";
            grid.Headers[0].Cells[5].Value = "Results";

            // OSM Grid Layout
            // 8 columns, 8 rows, Header - Cell 0, rowspan 2 - Cell 1, columnspan 3 - Cell 4, rowspan 2 - Cell 5, columnspan 3
            grid.Headers[0].Cells[0].RowSpan = 2;
            grid.Headers[0].Cells[0].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
            grid.Headers[0].Cells[1].ColumnSpan = 3;
            grid.Headers[0].Cells[1].StringFormat = new PdfStringFormat(PdfTextAlignment.Center);
            grid.Headers[0].Cells[4].RowSpan = 2;
            grid.Headers[0].Cells[4].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
            grid.Headers[0].Cells[5].ColumnSpan = 3;
            grid.Headers[0].Cells[5].StringFormat = new PdfStringFormat(PdfTextAlignment.Center);

            //Add new row to the header
            PdfGridRow[] header = grid.Headers.Add(1);
            header[1].Cells[0].Value = "";
            header[1].Cells[1].Value = "Min (mVdc)";
            header[1].Cells[1].StringFormat = new PdfStringFormat(PdfTextAlignment.Center,PdfVerticalAlignment.Middle);
            header[1].Cells[2].Value = "Actual";
            header[1].Cells[2].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
            header[1].Cells[3].Value = "Max (mVdc)";
            header[1].Cells[3].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
            header[1].Cells[4].Value = "";
            header[1].Cells[5].Value = "Min (mVdc)";
            header[1].Cells[5].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
            header[1].Cells[6].Value = "Actual";
            header[1].Cells[6].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
            header[1].Cells[7].Value = "Max (mVdc)";
            header[1].Cells[7].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);

            // Set vertical alignment for a specific column
            PdfStringFormat resultCellFormat = new PdfStringFormat
            {
                Alignment = PdfTextAlignment.Center
            };

            foreach (PdfGridRow gridRow in grid.Rows)
            {
                gridRow.Cells[1].Style.StringFormat = resultCellFormat;
                gridRow.Cells[2].Style.StringFormat = resultCellFormat;
                gridRow.Cells[3].Style.StringFormat = resultCellFormat;
                gridRow.Cells[5].Style.StringFormat = resultCellFormat;
                gridRow.Cells[6].Style.StringFormat = resultCellFormat;
                gridRow.Cells[7].Style.StringFormat = resultCellFormat;
            }

            // Set pass/fail colours
            for (int i = 0; i < grid.Rows.Count; i++)
            {
                // First column
                if (results[i].Average >= zeroTestStageValues[i,0] && results[i].Average <= zeroTestStageValues[i, 1])
                {
                    grid.Rows[i].Cells[2].Style.BackgroundBrush = new PdfSolidBrush(Color.LightGreen);
                }
                else
                {
                    grid.Rows[i].Cells[2].Style.BackgroundBrush = new PdfSolidBrush(Color.Red);
                    grid.Rows[i].Cells[2].Style.TextBrush = new PdfSolidBrush(Color.White);
                }

                // Second column
                int offsetValue = i + 5;
                if (results[offsetValue].Average >= zeroTestStageValues[offsetValue, 0] && results[offsetValue].Average <= zeroTestStageValues[offsetValue, 1])
                {
                    grid.Rows[i].Cells[6].Style.BackgroundBrush = new PdfSolidBrush(Color.LightGreen);
                }
                else
                {
                    grid.Rows[i].Cells[6].Style.BackgroundBrush = new PdfSolidBrush(Color.Red);
                    grid.Rows[i].Cells[6].Style.TextBrush = new PdfSolidBrush(Color.White);
                }
            }
            // Set the grid cell padding.
            grid.Style.CellPadding.All = 5;

            // Draw the table in page, below the line with a height gap of 20.
            layoutResult = grid.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 20));

            // Create the detailed results section
            textElement.Text = "Detailed Position Results";

            // Assign standard font to PdfTextElement.
            textElement.Font = new PdfStandardFont(PdfFontFamily.Helvetica, 14, PdfFontStyle.Bold);

            // Draw the header text on page, below the paragraph text with a height gap of 20 and maintain the position in PdfLayoutResult.
            layoutResult = textElement.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 20));

            // Initialize PdfLine with start point and end point for drawing the line.
            line = new PdfLine(new PointF(0, 0), new PointF(page.GetClientSize().Width, 0))
            {
                Pen = PdfPens.DarkGray
            };

            // Draw the line on page, below the header text with a height gap of 5 and maintain the position in PdfLayoutResult.
            layoutResult = line.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 5));

            // Detailed result output
            for (int i = 0; i < testStages.Length; i++)
            {

                // Load the paragraph text into PdfTextElement with initialized text font.
                PdfTextElement resultElement = new PdfTextElement(testStages[i].PadRight(10) + " - " + results[i].ToEngineeringString(), resultsFont, new PdfSolidBrush(Color.Black));

                // Draw the paragraph text on page and maintain the position in PdfLayoutResult.
                layoutResult = resultElement.Draw(page, new RectangleF(0, layoutResult.Bounds.Bottom + 5, page.GetClientSize().Width, page.GetClientSize().Height));
            }

            // Save the document.
            document.Save("Output"+DateTime.Now.ToLongTimeString().Replace(":", "-") + ".pdf");

            // Close the document.
            document.Close(true);
        }

        private static void CreeateAccuracyTestReport(StatisticalValues[] results)
        {
            // Based off the SyncFusion example code for PDF generation

            // Create a new PDF document.
            PdfDocument document = new PdfDocument();

            // Add a page to the document.
            PdfPage page = document.Pages.Add();

            // Create PDF graphics for the page.
            PdfGraphics graphics = page.Graphics;

            // Initialize our fonts font.
            PdfFont titleFont = new PdfStandardFont(PdfFontFamily.Helvetica, 20, PdfFontStyle.Bold);
            PdfFont textFont = new PdfStandardFont(PdfFontFamily.Helvetica, 10, PdfFontStyle.Bold);
            PdfFont resultsFont = new PdfStandardFont(PdfFontFamily.Courier, 10, PdfFontStyle.Bold);

            // Load the paragraph text into PdfTextElement with initialized standard font.
            PdfTextElement textElement = new PdfTextElement("Instrument Accuracy Test", titleFont, new PdfSolidBrush(Color.Blue));

            // Draw the paragraph text on page and maintain the position in PdfLayoutResult.
            PdfLayoutResult layoutResult = textElement.Draw(page, new RectangleF(0, 0, page.GetClientSize().Width, page.GetClientSize().Height));

            // Load the paragraph text into PdfTextElement with initialized standard font.
            textElement = new PdfTextElement("SPECIFICATION: ±1% of full scale on all ranges.", textFont, new PdfSolidBrush(Color.Black));

            // Draw the paragraph text on page and maintain the position in PdfLayoutResult.
            layoutResult = textElement.Draw(page, new RectangleF(0, layoutResult.Bounds.Bottom + 10, page.GetClientSize().Width, page.GetClientSize().Height));

            PdfImage image = PdfImage.FromImage(Properties.Resources.AccuracyTestSetup);

            // Calculate the scale factor based on the target width
            float scaleFactor = (float)page.GetClientSize().Width / image.Width;

            // Compute the new height while maintaining the aspect ratio
            int targetHeight = (int)(image.Height * scaleFactor);

            graphics.DrawImage(image, 0, layoutResult.Bounds.Bottom + 20, page.GetClientSize().Width, targetHeight);

            // Assign header text to PdfTextElement.
            textElement.Text = "Results";

            // Assign standard font to PdfTextElement.
            textElement.Font = new PdfStandardFont(PdfFontFamily.Helvetica, 14, PdfFontStyle.Bold);

            // Draw the header text on page, below the paragraph text with a height gap of 20 and maintain the position in PdfLayoutResult.
            layoutResult = textElement.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + targetHeight + 20));

            // Initialize PdfLine with start point and end point for drawing the line.
            PdfLine line = new PdfLine(new PointF(0, 0), new PointF(page.GetClientSize().Width, 0))
            {
                Pen = PdfPens.DarkGray
            };

            // Draw the line on page, below the header text with a height gap of 5 and maintain the position in PdfLayoutResult.
            layoutResult = line.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 5));

            // Initialize PdfGrid for drawing the table.
            PdfGrid grid = new PdfGrid();

            // accuracyTestStageValues

            List<ResultListRow> data = new List<ResultListRow>();
            data.Add(new ResultListRow("Fully CCW", ToEngineeringFormat.Convert(accuracyTestStageValues[0, 0], 4, "Vdc"), ToEngineeringFormat.Convert(results[0].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[0, 1], 4, "Vdc"), "5 Steps CW", ToEngineeringFormat.Convert(accuracyTestStageValues[5, 0], 4, "Vdc"), ToEngineeringFormat.Convert(results[5].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[5, 1], 4, "Vdc")));
            data.Add(new ResultListRow("1 Step CW", ToEngineeringFormat.Convert(accuracyTestStageValues[1,0], 4, "Vdc"), ToEngineeringFormat.Convert(results[1].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[1, 1], 4, "Vdc"), "6 Steps CW", ToEngineeringFormat.Convert(accuracyTestStageValues[6, 0], 4, "Vdc"), ToEngineeringFormat.Convert(results[6].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[6, 1], 4, "Vdc")));
            data.Add(new ResultListRow("2 Steps CW", ToEngineeringFormat.Convert(accuracyTestStageValues[2,0], 4, "Vdc"), ToEngineeringFormat.Convert(results[2].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[2, 1], 4, "Vdc"), "7 Steps CW", ToEngineeringFormat.Convert(accuracyTestStageValues[7, 0], 4, "Vdc"), ToEngineeringFormat.Convert(results[7].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[7, 1], 4, "Vdc")));
            data.Add(new ResultListRow("3 Steps CW", ToEngineeringFormat.Convert(accuracyTestStageValues[3,0], 4, "Vdc"), ToEngineeringFormat.Convert(results[3].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[3, 1], 4, "Vdc"), "8 Steps CW", ToEngineeringFormat.Convert(accuracyTestStageValues[8, 0], 4, "Vdc"), ToEngineeringFormat.Convert(results[8].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[8, 1], 4, "Vdc")));
            data.Add(new ResultListRow("4 Steps CW", ToEngineeringFormat.Convert(accuracyTestStageValues[4,0], 4, "Vdc"), ToEngineeringFormat.Convert(results[4].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[4, 1], 4, "Vdc"), "Fully CW", ToEngineeringFormat.Convert(accuracyTestStageValues[9, 0], 4, "Vdc"), ToEngineeringFormat.Convert(results[9].Average, 3, "Vdc"), ToEngineeringFormat.Convert(accuracyTestStageValues[9, 1], 4, "Vdc")));

            // Add list to IEnumerable.
            IEnumerable<object> dataTable = data;

            // Assign the DataTable as data source to grid.
            grid.DataSource = dataTable;

            //Using the Header collection
            grid.Headers[0].Cells[0].Value = "Range Switch Position";
            grid.Headers[0].Cells[1].Value = "Results";
            grid.Headers[0].Cells[4].Value = "Range Switch Position";
            grid.Headers[0].Cells[5].Value = "Results";

            // OSM Grid Layout
            // 8 columns, 8 rows, Header - Cell 0, rowspan 2 - Cell 1, columnspan 3 - Cell 4, rowspan 2 - Cell 5, columnspan 3
            grid.Headers[0].Cells[0].RowSpan = 2;
            grid.Headers[0].Cells[0].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
            grid.Headers[0].Cells[1].ColumnSpan = 3;
            grid.Headers[0].Cells[1].StringFormat = new PdfStringFormat(PdfTextAlignment.Center);
            grid.Headers[0].Cells[4].RowSpan = 2;
            grid.Headers[0].Cells[4].StringFormat = new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle);
            grid.Headers[0].Cells[5].ColumnSpan = 3;
            grid.Headers[0].Cells[5].StringFormat = new PdfStringFormat(PdfTextAlignment.Center);

            //Add new row to the header
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

            // Set vertical alignment for a specific column
            PdfStringFormat resultCellFormat = new PdfStringFormat
            {
                Alignment = PdfTextAlignment.Center
            };

            foreach (PdfGridRow gridRow in grid.Rows)
            {
                gridRow.Cells[1].Style.StringFormat = resultCellFormat;
                gridRow.Cells[2].Style.StringFormat = resultCellFormat;
                gridRow.Cells[3].Style.StringFormat = resultCellFormat;
                gridRow.Cells[5].Style.StringFormat = resultCellFormat;
                gridRow.Cells[6].Style.StringFormat = resultCellFormat;
                gridRow.Cells[7].Style.StringFormat = resultCellFormat;
            }

            // Set pass/fail colours
            for (int i = 0; i < grid.Rows.Count; i++)
            {
                // First column
                if (results[i].Average >= accuracyTestStageValues[i, 0] && results[i].Average <= accuracyTestStageValues[i, 1])
                {
                    grid.Rows[i].Cells[2].Style.BackgroundBrush = new PdfSolidBrush(Color.LightGreen);
                }
                else
                {
                    grid.Rows[i].Cells[2].Style.BackgroundBrush = new PdfSolidBrush(Color.Red);
                    grid.Rows[i].Cells[2].Style.TextBrush = new PdfSolidBrush(Color.White);
                }

                // Second column
                int offsetValue = i + 5;
                if (results[offsetValue].Average >= accuracyTestStageValues[offsetValue, 0] && results[offsetValue].Average <= accuracyTestStageValues[offsetValue, 1])
                {
                    grid.Rows[i].Cells[6].Style.BackgroundBrush = new PdfSolidBrush(Color.LightGreen);
                }
                else
                {
                    grid.Rows[i].Cells[6].Style.BackgroundBrush = new PdfSolidBrush(Color.Red);
                    grid.Rows[i].Cells[6].Style.TextBrush = new PdfSolidBrush(Color.White);
                }
            }
            // Set the grid cell padding.
            grid.Style.CellPadding.All = 5;

            // Draw the table in page, below the line with a height gap of 20.
            layoutResult = grid.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 20));

            // Create the detailed results section
            textElement.Text = "Detailed Position Results";

            // Assign standard font to PdfTextElement.
            textElement.Font = new PdfStandardFont(PdfFontFamily.Helvetica, 14, PdfFontStyle.Bold);

            // Draw the header text on page, below the paragraph text with a height gap of 20 and maintain the position in PdfLayoutResult.
            layoutResult = textElement.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 20));

            // Initialize PdfLine with start point and end point for drawing the line.
            line = new PdfLine(new PointF(0, 0), new PointF(page.GetClientSize().Width, 0))
            {
                Pen = PdfPens.DarkGray
            };

            // Draw the line on page, below the header text with a height gap of 5 and maintain the position in PdfLayoutResult.
            layoutResult = line.Draw(page, new PointF(0, layoutResult.Bounds.Bottom + 5));

            // Detailed result output
            for (int i = 0; i < testStages.Length; i++)
            {

                // Load the paragraph text into PdfTextElement with initialized text font.
                PdfTextElement resultElement = new PdfTextElement(testStages[i].PadRight(10) + " - " + results[i].ToEngineeringString(), resultsFont, new PdfSolidBrush(Color.Black));

                // Draw the paragraph text on page and maintain the position in PdfLayoutResult.
                layoutResult = resultElement.Draw(page, new RectangleF(0, layoutResult.Bounds.Bottom + 5, page.GetClientSize().Width, page.GetClientSize().Height));
            }

            // Save the document.
            document.Save("Output" + DateTime.Now.ToLongTimeString().Replace(":", "-") + ".pdf");

            // Close the document.
            document.Close(true);
        }

        private static void PrintMeasurementResults(string title, List<double> doubleList)
        {
            Console.WriteLine(title);
            //foreach (double value in doubleList)
            //{
            //    //var valueString = ToEngineeringFormat.Convert(double.Parse(QueryString("READ?")), 3, "Vdc");
            //    Console.WriteLine($"Reading: {ToEngineeringFormat.Convert(value, 3, "Vdc")}");
            //}
            Console.WriteLine($"Min Value: {ToEngineeringFormat.Convert(doubleList.Min(), 3, "Vdc")}");
            Console.WriteLine($"Max Value: {ToEngineeringFormat.Convert(doubleList.Max(), 3, "Vdc")}");
            Console.WriteLine($"Avg Value: {ToEngineeringFormat.Convert(doubleList.Average(), 3, "Vdc")}");
            Console.WriteLine($"SDev Value: {ToEngineeringFormat.Convert(StdDev(doubleList), 3, "Vdc")}");
        }

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
            SendCommand(":Display:Text:STate 1");
            Console.WriteLine(text);
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
            SendCommand(":Display:Text:STate 0");
        }
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

        public static List<double> ConvertStringToDoubleList(string input)
        {
            List<double> result = new List<double>();
            string[] values = input.Split(',');

            foreach (string value in values)
            {
                if (double.TryParse(value, out double doubleValue))
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

        public static void SRQHandler(object sender, Ivi.Visa.VisaEventArgs e)
        {
            /* Status Byte Bit Values
             * Bit 7 - OSB - Operation Summary Bit
             * Bit 6 - RQS - Request for Service
             * Bit 5 - ESB - Event Summary Bit
             * Bit 4 - MAV - Message Available
             * Bit 3 - QSB - Questionable Summary Bit
             * Bit 2 - EAV = Error Available
             * Bit 1 - Not Used
             * Bit 0 - MSB - Measurement Summary Bit
             */

            // Read the Status Byte
            // TODO: Check the status byte for the measurement summary bit
            var gbs = (GpibSession)sender;
            StatusByteFlags sb = gbs.ReadStatusByte();

            Console.WriteLine($"SRQHandler - Status Byte: {sb} - Note: User0 is the Measurement Summary Bit");

            gpibSession.DiscardEvents(EventType.ServiceRequest);

            srqWait.Release();
        }

        static private void SendCommand(string command)
        {
            gpibSession.FormattedIO.WriteLine(command);
        }

        static private string ReadResponse()
        {
            return gpibSession.FormattedIO.ReadLine();
        }

        static private string QueryString(string command)
        {
            SendCommand(command);
            return (ReadResponse());
        }
    }
}
