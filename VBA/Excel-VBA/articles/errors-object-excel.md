---
title: Errors Object (Excel)
keywords: vbaxl10.chm699072
f1_keywords:
- vbaxl10.chm699072
ms.prod: excel
api_name:
- Excel.Errors
ms.assetid: d2b50bbf-2685-fc5f-74c5-fa8bb9955f2a
ms.date: 06/08/2017
---


# Errors Object (Excel)

Represents the various spreadsheet errors for a range.


## Remarks

Use the  **[Errors](range-errors-property-excel.md)** property of the **[Range](range-object-excel.md)** object to return an **Errors** object.


## Example

Once an  **Errors** object is returned, you can use the **Value** property of the **[Error](error-object-excel.md)** object to check for particular error-checking conditions. The following example places a number as text in cell A1 and then notifies the user when the value of cell A1 contains a number as text.


```vb
Sub ErrorValue() 
 
 ' Place a number written as text in cell A1. 
 Range("A1").Formula = "'1" 
 
 If Range("A1").Errors.Item(xlNumberAsText).Value = True Then 
 MsgBox "Cell A1 has a number as text." 
 Else 
 MsgBox "Cell A1 is a number." 
 End If 
 
End Sub
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


