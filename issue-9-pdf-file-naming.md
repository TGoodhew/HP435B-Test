# Issue: PDF File Naming

The PDF file name uses `DateTime.Now.ToLongTimeString()` with colons replaced by dashes, but this could still result in file name collisions if tests are run within the same second. Consider using a more robust timestamp or unique identifier.
