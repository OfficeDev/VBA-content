---
title: ListObject.TotalsRowRange Property (Excel)
keywords: vbaxl10.chm734094
f1_keywords:
- vbaxl10.chm734094
ms.prod: excel
api_name:
- Excel.ListObject.TotalsRowRange
ms.assetid: 80f22712-5113-30d9-a0ea-1158a563d17b
ms.date: 06/08/2017
---


# ListObject.TotalsRowRange Property (Excel)

 Returns a **[Range](range-object-excel.md)** representing the Total row, if any, from a specified **ListObject** object. Read-only.


## Syntax

 _expression_ . **TotalsRowRange**

 _expression_ A variable that represents a **ListObject** object.


## Example

The following sample code returns the address of the Total row in the default list in Sheet1 of the active workbook. The code displays the Total row if it is not displayed already.


```vb
Sub DisplayTotalsRowAddress() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet2") 
 Set objListObj = wrksht.ListObjects(1) 
 objListObj.ShowTotals = True 
 MsgBox objListObj.TotalsRowRange.Address 
End Sub
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

