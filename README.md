# HP435B-Test

This is a short program that uses a HP 34401A DMM and a HP 11683A Range Calibrator to execute 3 performance tests from the HP 435B Owner & Service Manual:

- 4-6 Zero Carry Over Test
- 4-7 Instrumentation Accuracy Test with Calibrator
- 4-8 Calibration Factor Test

It uses a [community licensed version of the Syncfusion components ](https://www.syncfusion.com/products/communitylicense) to create the PDF reports.

## High-Level Overview

This program is a test automation tool for an HP435B instrument, interfaced using GPIB (General Purpose Interface Bus), likely for electrical measurements. The program allows the user to select and run different tests, gathers measurement data from the instrument, processes the results, and generates a PDF report summarizing the results.

---

## Key Components

- **GPIB Communication Setup**
  - Uses `NationalInstruments.Visa` and `Ivi.Visa` libraries to communicate with the instrument.
  - `GpibSession gpibSession` and `ResourceManager resManager` manage the hardware interface.

- **Test Stages & Limits**
  - Arrays like `testRangeStages`, `zeroTestStageValues`, `accuracyTestStageValues`, and `calibrationFactorTestStageValues` define the steps and pass/fail limits for different tests.
  - `testCalibrationStages` is populated at runtime with calibration steps.

- **Results Handling**
  - `StatisticalValues` struct: Stores statistics (min, max, average, standard deviation) for each test stage.
  - `ResultListRow` class: Used for formatting results into report tables.

- **Main Program Logic (`Main` method)**
  1. **Initialize GPIB Connection**: Opens and configures the instrument connection.
  2. **Prepare Calibration Array**: Fills calibration positions for the calibration test.
  3. **User Interaction Loop**: 
     - Prompts the user to select a test.
     - For each test:
       - Displays instrument ID.
       - Configures the instrument.
       - Runs the selected test via `TestRun()`, which interacts with the user and collects measurement data.
       - Creates a PDF report of the results via `CreateTestReport()`.
       - Prompts the user to open the report and/or run another test.
  4. **Cleanup**: Closes the GPIB session and resource manager.

---

## Core Methods

- **TestRun**  
  Loops through each test stage, prompts the user to set the device, collects measurement data, and updates a live results table in the terminal.

- **SetupDMM**  
  Configures the Digital Multimeter (DMM) to the correct mode and range for testing.

- **GetData**  
  Triggers measurement, waits for service request, fetches results, and calculates statistics.

- **CreateTestReport**  
  Builds a PDF report with the test results, highlighting pass/fail states, and detailed statistics for each step.

- **SRQHandler**  
  Handles GPIB service requests (instrument signaling end of measurement).

- **Utility Methods**  
  - `SendCommand`, `ReadResponse`, `QueryString`: For sending commands and receiving responses from the instrument.
  - `StdDev`, `ConvertStringToDoubleList`: For data processing and statistics.

---

## User Experience

- The program runs entirely in the terminal, using `Spectre.Console` for rich prompts and tables.
- It guides the user through each test step-by-step, requesting manual actions and instrument adjustments.
- After each test, it generates a PDF report and optionally opens it for the user.

---

## Example Flow

1. User starts the program.
2. Chooses a test (e.g., "Zero Carryover").
3. Follows prompts to configure the instrument and DUT (Device Under Test) for each stage.
4. Results are displayed live in the terminal.
5. A PDF summary report is generated - [View Example File](https://github.com/TGoodhew/HP435B-Test/blob/master/AccuracyTestReport1-44-30%20PM.pdf)
6. User can open the report or run another test.

## License

This project is licensed under the [MIT License](LICENSE.txt).

See the LICENSE.txt file for full license details.
