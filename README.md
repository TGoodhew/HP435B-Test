# HP435B-Test

A comprehensive automated test application for the HP 435B Power Meter that executes performance verification tests from the HP 435B Owner & Service Manual. This program uses an HP 34401A Digital Multimeter and HP 11683A Range Calibrator to perform automated testing and generate professional PDF reports.

## About the HP 435B Power Meter

The **HP 435B** is a precision analog power meter designed for accurate RF and microwave power measurements. This versatile instrument is commonly used in laboratory and field service applications for testing RF equipment, communication systems, and microwave devices.

### Key Features and Specifications:

- **Power Reference Output**: Internal 50 MHz oscillator with Type N Female connector on front panel provides 1.00 mW output (factory set to ±0.7% traceable to the National Bureau of Standards)
- **Measurement Range**: Wide dynamic range for power measurements from microwatts to hundreds of milliwatts
- **Frequency Coverage**: Operates across a broad frequency spectrum when used with appropriate thermistor or diode power sensors
- **Calibration Factor Switch**: 16-position switch (85% to 100% in 1% steps) normalizes meter readings to account for sensor calibration factor or effective efficiency
- **Range Settings**: 10 power ranges accessible via front-panel rotary switch for versatile measurement capability
- **Zero Function**: Front-panel ZERO switch enables zeroing in the most sensitive range for accurate low-level measurements
- **Zero Carryover Accuracy**: ±0.5% of full scale when zeroed in the most sensitive range
- **Instrumentation Accuracy**: ±1% of full scale on all ranges
- **Recorder Output**: Rear-panel BNC connector provides DC output voltage proportional to meter reading

The 435B's analog design provides reliable measurements and is valued for its durability, accuracy, and ease of use in professional test environments.

## Performance Tests

This application implements three critical performance tests from Section IV of the HP 435B Owner & Service Manual (Part No. 00435-90040) to verify the instrument meets its published specifications:

### Test 4-6: Zero Carryover Test

**Purpose**: Validates that zero carryover remains within specification across all 10 range switch positions.

**Specification**: ±0.5% of full scale when zeroed in the most sensitive range.

**Test Method**: 
- After zeroing the power meter in the most sensitive range (fully counter-clockwise position), the meter reading is monitored at the RECORDER OUTPUT as the instrument is stepped through all ranges
- The change in meter reading is measured using a Digital Voltmeter connected to the rear-panel RECORDER OUTPUT
- Readings account for both noise and drift, as these cannot be separated from zero carryover

**Test Limits** (in mVdc at RECORDER OUTPUT):

| Range Position | Minimum | Maximum |
|---------------|---------|---------|
| Fully CCW (most sensitive) | -15 mV | +15 mV |
| 1 Step CW | -17 mV | +17 mV |
| 2 Steps CW | -14 mV | +14 mV |
| 3 Steps CW | -11 mV | +11 mV |
| 4 Steps CW | -8 mV | +8 mV |
| 5 Steps CW | -5 mV | +5 mV |
| 6 Steps CW | -5 mV | +5 mV |
| 7 Steps CW | -5 mV | +5 mV |
| 8 Steps CW | -5 mV | +5 mV |
| Fully CW | -5 mV | +5 mV |

### Test 4-7: Instrumentation Accuracy Test with Calibrator

**Purpose**: Verifies the overall measurement accuracy of the instrument using the HP 11683A Range Calibrator as a calibrated reference source.

**Specification**: ±1% of full scale on all ranges.

**Test Method**:
- The HP 11683A Range Calibrator provides a precise full-scale reference input to the Power Meter on each range
- The RECORDER OUTPUT level is measured using a Digital Voltmeter
- Verification confirms that readings are within ±1% plus noise and drift
- The CAL ADJ control is used to set a 1000 mVdc reading at the reference position (5 steps from fully CCW)
- Each range is then tested by setting both the Power Meter and Calibrator RANGE switches to the same position

**Test Limits** (in mVdc at RECORDER OUTPUT with 1 mW calibrated input):

| Range Position | Minimum | Maximum |
|---------------|---------|---------|
| Fully CCW | 975 mV | 1025 mV |
| 1 Step CW | 978 mV | 1022 mV |
| 2 Steps CW | 981 mV | 1019 mV |
| 3 Steps CW | 984 mV | 1016 mV |
| 4 Steps CW | 987 mV | 1013 mV |
| 5 Steps CW | 998 mV | 1002 mV |
| 6 Steps CW | 990 mV | 1010 mV |
| 7 Steps CW | 990 mV | 1010 mV |
| 8 Steps CW | 990 mV | 1015 mV |
| Fully CW | 990 mV | 1015 mV |

### Test 4-8: Calibration Factor Test

**Purpose**: Confirms that the 16-position calibration factor switch correctly normalizes meter readings.

**Specification**: 16-position switch normalizes meter reading to account for calibration factor or effective efficiency. Range 85% to 100% in 1% steps.

**Test Method**:
- After zeroing the Power Meter on the most sensitive range, a 1 mW input level is applied using the HP 11683A Calibrator
- The CAL ADJ control is adjusted to obtain a 1.000 Vdc indication at the RECORDER OUTPUT
- The CAL FACTOR switch is then stepped through all 16 positions (100% down to 85%)
- The meter reading is monitored at each position to ensure proper indication

**Test Limits** (in Vdc at RECORDER OUTPUT with 1 mW input):

| CAL FACTOR Position | Minimum | Maximum |  | CAL FACTOR Position | Minimum | Maximum |
|---------------------|---------|---------|---|---------------------|---------|---------|
| 100% | 0.994 V | 1.006 V | | 92% | 1.081 V | 1.093 V |
| 99% | 1.004 V | 1.016 V | | 91% | 1.093 V | 1.105 V |
| 98% | 1.014 V | 1.026 V | | 90% | 1.105 V | 1.117 V |
| 97% | 1.025 V | 1.037 V | | 89% | 1.118 V | 1.130 V |
| 96% | 1.036 V | 1.048 V | | 88% | 1.130 V | 1.142 V |
| 95% | 1.047 V | 1.059 V | | 87% | 1.143 V | 1.155 V |
| 94% | 1.058 V | 1.070 V | | 86% | 1.157 V | 1.169 V |
| 93% | 1.069 V | 1.081 V | | 85% | 1.170 V | 1.182 V |

**Purpose of Calibration Factor**: The calibration factor switch allows the power meter to account for the effective efficiency of different power sensors. When a sensor has less than 100% efficiency (e.g., 92%), the calibration factor normalizes the reading so that the displayed power is the actual input power, not just the absorbed power.

## Syncfusion PDF Components

This application uses the [Syncfusion PDF library](https://www.syncfusion.com/pdf-framework/net) to generate professional test reports with tables, images, and formatted results.

### Getting a Community License

Syncfusion offers a **free community license** for individual developers and small businesses:

**Eligibility**:
- Individual developers
- Companies with less than $1 million USD in annual gross revenue
- Companies with 5 or fewer developers

**How to Obtain a Community License**:

1. Visit the [Syncfusion Community License page](https://www.syncfusion.com/products/communitylicense)
2. Click on **"Claim Free License"**
3. Sign up or log in to your Syncfusion account
4. Fill out the community license form confirming your eligibility
5. You will receive your license key via email within 24-48 hours

**Registering Your License**:

Once you receive your license key, register it in your application by adding the following code before any Syncfusion component is used:

```csharp
// Register Syncfusion license
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```

For this application, you would typically add this in the `Main` method of `Program.cs` before initializing any PDF components.

**Note**: Without a registered license, the application will work but generated PDFs will include a watermark and warning message.

**More Information**:
- [Syncfusion Licensing FAQ](https://www.syncfusion.com/products/communitylicense/faq)
- [Syncfusion PDF Library Documentation](https://help.syncfusion.com/file-formats/pdf/overview)

## Prerequisites

### Hardware Requirements

This test application requires the following equipment:

- **HP 435B Power Meter** - The device under test (DUT)
- **HP 34401A Digital Multimeter** - Used to measure the 435B's RECORDER OUTPUT voltage (GPIB address 14 by default)
- **HP 11683A Range Calibrator** - Provides calibrated 1 mW reference signals for accuracy and calibration factor testing
- **GPIB Interface** - National Instruments or compatible GPIB controller card/adapter (USB-GPIB adapters such as NI GPIB-USB-HS are recommended)
- **BNC Cables** - For connecting the 435B RECORDER OUTPUT to the DMM and the Calibrator to the 435B Power Sensor input
- **Power Sensor** - HP 8481A/H or 8482A/H thermistor power sensor (or equivalent) for connecting to the 435B

### Software Requirements

- **Operating System**: Windows 7 or later (required for GPIB drivers)
- **.NET Framework**: Version 4.7.2 or later
- **NI-VISA**: National Instruments VISA drivers for GPIB communication
  - Download from [National Instruments VISA Download](https://www.ni.com/en-us/support/downloads/drivers/download.ni-visa.html)
  - Version 25.5.0.13 or later recommended
- **Visual Studio 2019 or later** (for building from source)
  - Community Edition is sufficient

### NuGet Package Dependencies

The following packages are automatically restored when building:

- `Syncfusion.Pdf.Base` (v30.1462.37.0) - PDF generation library
- `Syncfusion.Licensing` (v30.1462.37.0) - License management for Syncfusion components
- `NationalInstruments.Visa` (v25.5.0.13) - GPIB communication with instruments
- `Spectre.Console` (0.50.0) - Rich terminal UI with colors, tables, and prompts
- `Ivi.Visa` (v8.0.0.0) - VISA interface library for instrument control

## Installation and Setup

### Building from Source

1. **Clone the repository**:
   ```bash
   git clone https://github.com/TGoodhew/HP435B-Test.git
   cd HP435B-Test
   ```

2. **Install NI-VISA drivers** (if not already installed):
   - Download and install NI-VISA from National Instruments
   - Restart your computer after installation
   - Verify installation by running NI MAX (Measurement & Automation Explorer)

3. **Open the solution**:
   - Open `HP435B-Test.sln` in Visual Studio

4. **Restore NuGet packages**:
   - Visual Studio should automatically restore packages on first build
   - Or manually: Right-click solution → Restore NuGet Packages

5. **Configure GPIB Address** (if different from default):
   - In `Program.cs`, locate the line:
     ```csharp
     private static int gpibIntAddress = 14;
     ```
   - Change `14` to match your HP 34401A's GPIB address
   - You can verify your DMM's GPIB address using NI MAX

6. **Build the solution**:
   - Build → Build Solution (or press Ctrl+Shift+B)
   - Ensure there are no build errors

7. **Run the application**:
   - Press F5 to run in debug mode, or Ctrl+F5 to run without debugging

### Hardware Setup

1. **Connect GPIB interface** to your computer (USB or PCIe)
2. **Connect HP 34401A** to the GPIB bus and verify/set its address (default: 14)
3. **Power on** all equipment and allow proper warm-up time:
   - HP 435B: Minimum 30 minutes warm-up required for valid performance tests
   - HP 34401A: 1 hour warm-up recommended for best accuracy
   - HP 11683A: 30 minutes warm-up recommended
4. **Connect test equipment** as follows:
   - Connect power sensor to HP 435B front-panel power input
   - Connect HP 11683A Calibrator output to power sensor input (for accuracy tests)
   - Connect HP 435B rear-panel RECORDER OUTPUT to HP 34401A input using BNC cable
   - Ensure proper 50Ω termination on all RF connections

5. **Verify GPIB Communication**:
   - Use NI MAX to verify that the HP 34401A is detected on the GPIB bus
   - Send a simple query (e.g., `*IDN?`) to confirm communication

## Usage

### Starting the Application

1. Launch the HP435B-Test application
2. The program will display a welcome screen with the HP435B Test title
3. Available tests are presented with their specifications

### Selecting a Test

The application provides an interactive menu with the following options:

- **Settings** - Configure GPIB address and other parameters
- **Zero Carryover** - Execute Test 4-6 from the HP manual
- **Instrument Accuracy with Calibrator** - Execute Test 4-7 from the HP manual
- **Calibration Factor** - Execute Test 4-8 from the HP manual
- **Exit** - Close the application

### Running a Test

1. Select a test from the menu
2. The program will display the instrument ID from the HP 34401A
3. Follow the on-screen prompts to:
   - Set the HP 435B RANGE switch to the specified position
   - Set the HP 11683A Calibrator controls (for applicable tests)
   - Press Enter when ready for each measurement
4. The application will:
   - Configure the HP 34401A DMM automatically
   - Trigger measurements and collect data
   - Display live results in a formatted table
   - Show pass/fail status for each test step
5. After test completion:
   - A PDF report is automatically generated
   - Choose whether to open the report immediately
   - Choose whether to run another test

### Understanding the Results

- **Green** text indicates passing results within specification limits
- **Red** text indicates failing results outside specification limits
- Results include statistical data: Min, Max, Average, and Standard Deviation
- Each stage of testing is clearly labeled (e.g., "Fully CCW", "1 Step CW", etc.)

### PDF Report Format

Generated reports include:

- Test title and specification
- Test setup diagram (embedded image)
- Date and time of test execution
- Detailed results table with:
  - Range/switch position
  - Specification limits (min/max)
  - Actual measured values
  - Pass/Fail indication
- Statistical summary for each test stage

Reports are saved in the application directory with descriptive filenames including timestamp.

## High-Level Overview

This program is a test automation tool for the HP 435B power meter, interfaced using GPIB (General Purpose Interface Bus). The program allows the user to select and run different performance verification tests, gathers measurement data from the HP 34401A DMM, processes the results against published specifications, and generates professional PDF reports summarizing the test outcomes.

## Key Components

### GPIB Communication Setup
- Uses `NationalInstruments.Visa` and `Ivi.Visa` libraries to communicate with the HP 34401A
- `GpibSession gpibSession` and `ResourceManager resManager` manage the hardware interface
- Implements Service Request (SRQ) handling for efficient measurement triggering

### Test Stages & Limits
- `testRangeStages`: Array defining the 10 range switch positions
- `zeroTestStageValues`: Specification limits for Zero Carryover Test (Test 4-6)
- `accuracyTestStageValues`: Specification limits for Instrumentation Accuracy Test (Test 4-7)
- `calibrationFactorTestStageValues`: Specification limits for Calibration Factor Test (Test 4-8)
- `testCalibrationStages`: Array populated at runtime with 16 calibration positions (100% to 85%)

### Results Handling
- `StatisticalValues` struct: Stores statistics (min, max, average, standard deviation) for each test stage
- `ResultListRow` class: Formats results into report tables with proper engineering notation
- `ToEngineeringFormat` class: Converts numerical values to engineering notation with SI prefixes

### Main Program Logic

1. **Initialize GPIB Connection**: Opens and configures the HP 34401A connection
2. **Prepare Calibration Array**: Fills calibration positions (100% down to 85%)
3. **User Interaction Loop**: 
   - Prompts user to select a test
   - Displays instrument ID
   - Configures the DMM for the selected test
   - Runs the test via `TestRun()` method with user prompts and data collection
   - Creates PDF report via `CreateTestReport()` method
   - Prompts user to open report and/or run another test
4. **Cleanup**: Closes GPIB session and resource manager

## Core Methods

- **TestRun**: Loops through each test stage, prompts the user to set controls, collects measurement data using `GetData()`, and updates a live results table in the terminal
- **SetupDMM**: Configures the HP 34401A to the correct measurement mode, range, trigger, and NPLC settings for accurate testing
- **GetData**: Triggers measurement, waits for service request (SRQ), fetches results from the DMM, and calculates statistical values
- **CreateTestReport**: Builds a professional PDF report using Syncfusion components with test results, setup diagrams, highlighting pass/fail states, and detailed statistics
- **SRQHandler**: Handles GPIB service requests (instrument signaling end of measurement) using semaphore synchronization
- **Utility Methods**: 
  - `SendCommand`, `ReadResponse`, `QueryString`: GPIB communication wrappers
  - `StdDev`, `ConvertStringToDoubleList`: Statistical calculation and data parsing

## User Experience

- The program runs entirely in the terminal using `Spectre.Console` for rich prompts, colored text, and formatted tables
- Step-by-step guidance through each test with clear instructions
- Real-time display of measurement results as they are collected
- Professional PDF reports generated automatically after each test
- Option to open reports immediately or save for later review

## Example Flow

1. User starts the program
2. Application displays welcome screen and available tests
3. User chooses a test (e.g., "Instrument Accuracy with Calibrator")
4. Program displays HP 34401A instrument ID
5. User follows prompts to:
   - Configure HP 435B RANGE switch
   - Configure HP 11683A Calibrator controls
   - Initiate each measurement
6. Results are displayed live in a formatted table with pass/fail indicators
7. PDF summary report is generated - [View Example File](https://github.com/TGoodhew/HP435B-Test/blob/master/AccuracyTestReport1-44-30%20PM.pdf)
8. User can open the report or run another test

## Troubleshooting

### GPIB Communication Issues

- **DMM not found**: Verify GPIB address matches the setting in code (default 14)
- **Timeout errors**: Check GPIB cables and connections
- **SRQ not received**: Ensure NI-VISA drivers are properly installed
- Use NI MAX (Measurement & Automation Explorer) to verify instrument connectivity

### Test Failures

- **Zero Carryover out of spec**: May indicate need for adjustment or repair of HP 435B
- **Accuracy out of spec**: Verify 30-minute warm-up has been completed; check HP 11683A calibrator accuracy
- **Calibration Factor out of spec**: May indicate issues with A4R66 or other factory-selected components

### Build Issues

- **NuGet restore failures**: Check internet connection; manually restore packages
- **Syncfusion license warnings**: Obtain and register a community license (see above)
- **VISA references not found**: Install NI-VISA runtime and development support

## References

- **HP 435B Operating and Service Manual**: Part No. 00435-90040 (included as 435B-Copilot.pdf)
- **HP 34401A User's Guide**: Available from Keysight Technologies
- **HP 11683A Range Calibrator Operating Guide**: Available from Keysight Technologies
- **NI-VISA Documentation**: [ni.com/visa](https://www.ni.com/en-us/support/documentation/supplemental/06/ni-visa-overview.html)

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## License

This project is licensed under the [MIT License](LICENSE.txt).

See the LICENSE.txt file for full license details.

## Acknowledgments

- HP/Agilent/Keysight Technologies for the original HP 435B and test equipment
- National Instruments for VISA drivers and GPIB support
- Syncfusion for the community-licensed PDF generation components
- Spectre.Console for the excellent terminal UI library
