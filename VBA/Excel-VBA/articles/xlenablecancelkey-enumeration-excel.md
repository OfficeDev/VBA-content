---
title: XlEnableCancelKey Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlEnableCancelKey
ms.assetid: ccf1a7d1-c2fe-7a7e-16d8-ebb4ebf5ba6b
ms.date: 06/08/2017
---


# XlEnableCancelKey Enumeration (Excel)

Specifies how Microsoft Office Excel 2007 handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlDisabled**|0|Cancel key trapping is completely disabled.|
| **xlErrorHandler**|2|The interrupt is sent to the running procedure as an error, trappable by an error handler set up with an On Error GoTo statement. The trappable error code is 18.|
| **xlInterrupt**|1|The current procedure is interrupted, and the user can debug or end the procedure.|

