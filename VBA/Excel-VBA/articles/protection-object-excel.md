---
title: Protection Object (Excel)
keywords: vbaxl10.chm719072
f1_keywords:
- vbaxl10.chm719072
ms.prod: excel
api_name:
- Excel.Protection
ms.assetid: dc13a9dd-bd19-daa2-5093-7182917d5bde
ms.date: 06/08/2017
---


# Protection Object (Excel)

Represents the various types of protection options available for a worksheet.


## Remarks

Use the  **[Protection](worksheet-protection-property-excel.md)** property of the **[Worksheet](worksheet-object-excel.md)** object to return a **Protection** object.

Once a  **Protection** object is returned, you can use its following properties, to set or return protection options.


-  **[AllowDeletingColumns](protection-allowdeletingcolumns-property-excel.md)**
    
-  **[AllowDeletingRows](protection-allowdeletingrows-property-excel.md)**
    
-  **[AllowFiltering](protection-allowfiltering-property-excel.md)**
    
-  **[AllowFormattingCells](protection-allowformattingcells-property-excel.md)**
    
-  **[AllowFormattingColumns](protection-allowformattingcolumns-property-excel.md)**
    
-  **[AllowFormattingRows](protection-allowformattingrows-property-excel.md)**
    
-  **[AllowInsertingColumns](protection-allowinsertingcolumns-property-excel.md)**
    
-  **[AllowInsertingHyperlinks](protection-allowinsertinghyperlinks-property-excel.md)**
    
-  **[AllowInsertingRows](protection-allowinsertingrows-property-excel.md)**
    
-  **[AllowSorting](protection-allowsorting-property-excel.md)**
    
-  **[AllowUsingPivotTables](protection-allowusingpivottables-property-excel.md)**
    

## Example

The following example demonstrates how to use the  **[AllowInsertingColumns](protection-allowinsertingcolumns-property-excel.md)** property of the **Protection** object, placing three numbers in the top row and protecting the worksheet. Then this example checks to see if the protection setting for allowing the insertion of columns is **False** and sets it to **True**, if necessary. Finally, it notifies the user to insert a column.


```
Sub SetProtection() 
 
 Range("A1").Formula = "1" 
 Range("B1").Formula = "3" 
 Range("C1").Formula = "4" 
 ActiveSheet.Protect 
 
 ' Check the protection setting of the worksheet and act accordingly. 
 If ActiveSheet.Protection.AllowInsertingColumns = False Then 
 ActiveSheet.Protect AllowInsertingColumns:=True 
 MsgBox "Insert a column between 1 and 3" 
 Else 
 MsgBox "Insert a column between 1 and 3" 
 End If 
 
End Sub
```


## Properties



|**Name**|
|:-----|
|[AllowDeletingColumns](protection-allowdeletingcolumns-property-excel.md)|
|[AllowDeletingRows](protection-allowdeletingrows-property-excel.md)|
|[AllowEditRanges](protection-alloweditranges-property-excel.md)|
|[AllowFiltering](protection-allowfiltering-property-excel.md)|
|[AllowFormattingCells](protection-allowformattingcells-property-excel.md)|
|[AllowFormattingColumns](protection-allowformattingcolumns-property-excel.md)|
|[AllowFormattingRows](protection-allowformattingrows-property-excel.md)|
|[AllowInsertingColumns](protection-allowinsertingcolumns-property-excel.md)|
|[AllowInsertingHyperlinks](protection-allowinsertinghyperlinks-property-excel.md)|
|[AllowInsertingRows](protection-allowinsertingrows-property-excel.md)|
|[AllowSorting](protection-allowsorting-property-excel.md)|
|[AllowUsingPivotTables](protection-allowusingpivottables-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
