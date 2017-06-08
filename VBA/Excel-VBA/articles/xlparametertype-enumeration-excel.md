---
title: XlParameterType Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlParameterType
ms.assetid: f6774f89-4992-2b7c-2dce-791fecafc1df
ms.date: 06/08/2017
---


# XlParameterType Enumeration (Excel)

Specifies how to determine the value of the parameter for the specified query table.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlConstant**|1|Uses the value specified by the  **Value** argument.|
| **xlPrompt**|0|Displays a dialog box that prompts the user for the value. The  **Value** argument specifies the text shown in the dialog box.|
| **xlRange**|2|Uses the value of the cell in the upper-left corner of the range. The  **Value** argument specifies a **Range** object.|

