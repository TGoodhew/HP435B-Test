# Issue: Windows-Specific File Open Command

The code uses `explorer.exe` to open PDF reports, which is specific to Windows. This will not work on Linux or macOS. Consider using a cross-platform approach, such as detecting the OS and using `xdg-open` on Linux or `open` on macOS.
