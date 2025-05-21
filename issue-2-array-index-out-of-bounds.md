# Issue: Potential Array Index Out of Bounds

The `results` array is always initialized with 16 elements, but some tests (e.g., "Zero Carryover") use only 10 stages. This could cause index out of range exceptions in report generation or result assignment. Ensure array sizes match the number of test stages.
