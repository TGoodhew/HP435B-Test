# Issue: Resource Disposal in Finally Block

The `finally` block disposes `gpibSession` and `resManager` without null checks after assignment. If initialization fails, this could throw a `NullReferenceException`. Add null checks before disposing resources.
