---
title: Application.Windows Property (Excel)
keywords: vbaxl10.chm183114
f1_keywords:
- vbaxl10.chm183114
ms.prod: excel
api_name:
- Excel.Application.Windows
ms.assetid: 07e54620-c2f5-2354-f313-9756a0f17425
ms.date: 06/08/2017
---


# Application.Windows Property (Excel)

Returns a  **[Windows](windows-object-excel.md)** collection that represents all the windows in all the workbooks. Read-only **Windows** object.


## Syntax

 _expression_ . **Windows**

 _expression_ A variable that represents an **Application** object.


## Remarks

Using this property without an object qualifier is equivalent to using  `Application.Windows`.

This property returns a collection of both visible and hidden windows.


## Example

This example closes the first open or hidden window in Microsoft Excel.


```vb
Application.Windows(1).Close
```

This example names window one in the active workbook "Consolidated Balance Sheet." This name is then used as the index to the  **Windows** collection.




```vb
ActiveWorkbook.Windows(1).Caption = "Consolidated Balance Sheet" 
ActiveWorkbook.Windows("Consolidated Balance Sheet") _ 
 .ActiveSheet.Calculate
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

