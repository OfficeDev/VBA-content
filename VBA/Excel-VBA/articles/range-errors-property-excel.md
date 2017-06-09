---
title: Range.Errors Property (Excel)
keywords: vbaxl10.chm144235
f1_keywords:
- vbaxl10.chm144235
ms.prod: excel
api_name:
- Excel.Range.Errors
ms.assetid: 88dcc606-d412-a9ce-82bc-5fbba8baae87
ms.date: 06/08/2017
---


# Range.Errors Property (Excel)

Allows the user to to access error checking options.


## Syntax

 _expression_ . **Errors**

 _expression_ A variable that represents a **Range** object.


## Remarks

Reference the  **[Errors](errors-object-excel.md)** object to view a list of index values associated with error checking options.


## Example

In this example, a number written as text is placed in cell A1. Microsoft Excel then determines if the number is written as text in cell A1 and notifies the user accordingly.


```vb
Sub CheckForErrors() 
 
 Range("A1").Formula = "'12" 
 
 If Range("A1").Errors.Item(xlNumberAsText).Value = True Then 
 MsgBox "The number is written as text." 
 Else 
 MsgBox "The number is not written as text." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

