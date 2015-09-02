
# XlEnableCancelKey Enumeration (Excel)

 **Last modified:** July 28, 2015

Specifies how Microsoft Office Excel 2007 handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure.


|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlDisabled**|0|Cancel key trapping is completely disabled.|
| **xlErrorHandler**|2|The interrupt is sent to the running procedure as an error, trappable by an error handler set up with an On Error GoTo statement. The trappable error code is 18.|
| **xlInterrupt**|1|The current procedure is interrupted, and the user can debug or end the procedure.|
