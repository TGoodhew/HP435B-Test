# Issue: Thread Safety

Static fields for GPIB session and semaphore could cause issues if the program is extended to multi-threaded or multi-session use. Consider refactoring to avoid static state or ensure thread safety.
