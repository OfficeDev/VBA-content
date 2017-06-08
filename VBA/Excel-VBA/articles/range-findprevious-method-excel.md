---
title: Range.FindPrevious Method (Excel)
keywords: vbaxl10.chm144130
f1_keywords:
- vbaxl10.chm144130
ms.prod: excel
api_name:
- Excel.Range.FindPrevious
ms.assetid: c03f2e17-d28c-8b0d-b8c8-024863523c99
ms.date: 06/08/2017
---


# Range.FindPrevious Method (Excel)

Continues a search that was begun with the  **[Find](range-find-method-excel.md)** method. Finds the previous cell that matches those same conditions and returns a **[Range](range-object-excel.md)** object that represents that cell. Doesn?t affect the selection or the active cell.


## Syntax

 _expression_ . **FindPrevious**( **_After_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _After_|Optional| **Variant**|The cell before which you want to search. This corresponds to the position of the active cell when a search is done from the user interface. Note that  _After_ must be a single cell in the range. Remember that the search begins before this cell; the specified cell isn?t searched until the method wraps back around to this cell. If this argument isn?t specified, the search starts before the upper- left cell in the range.|

### Return Value

Range


## Remarks

When the search reaches the beginning of the specified search range, it wraps around to the end of the range. To stop a search when this wraparound occurs, save the address of the first found cell, and then test each successive found-cell address against this saved address.


## Example

This example shows how the  **FindPrevious** method is used with the **Find** and **[FindNext](range-findnext-method-excel.md)** methods. Before running this example, make sure that Sheet1 contains at least two occurrences of the word ?Phoenix? in column B.


```vb
Sub FindTest() 
 Dim fc As Range 
 Set fc = Worksheets("Sheet1").Columns("B").Find(what:="Phoenix") 
 MsgBox "The first occurrence is in cell " &; fc.Address 
 Set fc = Worksheets("Sheet1").Columns("B").FindNext(after:=fc) 
 MsgBox "The next occurrence is in cell " &; fc.Address 
 Set fc = Worksheets("Sheet1").Columns("B").FindPrevious(after:=fc) 
 MsgBox "The previous occurrence is in cell " &; fc.Address 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

