This repository demonstrates a subtle bug in VBScript's `IsEmpty` function.  The `IsEmpty` function, when used with a `Null` value, unexpectedly returns `False`. This can lead to logic errors in code handling optional parameters or database values that may contain `Null`.

The `bug.vbs` file contains code that showcases this issue. The `bugSolution.vbs` file offers a corrected approach using more robust null checks.