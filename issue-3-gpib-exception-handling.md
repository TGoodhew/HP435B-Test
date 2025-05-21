# Issue: Lack of Exception Handling for GPIB Operations

GPIB communication (e.g., `SendCommand`, `QueryString`, session creation) is not wrapped in try/catch blocks. Device errors could cause the program to crash. Add exception handling to improve robustness.
