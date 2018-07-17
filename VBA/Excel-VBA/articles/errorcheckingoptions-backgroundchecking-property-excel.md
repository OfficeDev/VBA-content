---
title: ErrorCheckingOptions.BackgroundChecking Property (Excel)
keywords: vbaxl10.chm698073
f1_keywords:
- vbaxl10.chm698073
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.BackgroundChecking
ms.assetid: 427b3b32-c099-dc71-1255-7f73805a4db8
ms.date: 06/08/2017
---


# ErrorCheckingOptions.BackgroundChecking Property (Excel)

Alerts the user for all cells that violate enabled error-checking rules. When this property is set to  **True** (default), the **AutoCorrect Options** button appears next to all cells that violate enabled errors. **False** disables background checking for errors. Read/write **Boolean** .


## Syntax

 _expression_ . **BackgroundChecking**

 _expression_ A variable that represents an **ErrorCheckingOptions** object.


## Remarks

Refer to the  **[ErrorCheckingOptions](errorcheckingoptions-object-excel.md)** object to view a list of its members that can be enabled.


## Example

In the following example, when the user selects cell A1 (which contains a formula referring to empty cells), the  **AutoCorrect Options** button appears.


```vb
Sub CheckBackground() 
 
 ' Simulate an error by referring to empty cells. 
 Application.ErrorCheckingOptions.BackgroundChecking = True 
 Range("A1").Select 
 ActiveCell.Formula = "=A2/A3" 
 
End Sub
```


## See also


#### Concepts


[ErrorCheckingOptions Object](errorcheckingoptions-object-excel.md)

