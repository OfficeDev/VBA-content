---
title: Application.Goto Method (Excel)
keywords: vbaxl10.chm133144
f1_keywords:
- vbaxl10.chm133144
ms.prod: excel
api_name:
- Excel.Application.Goto
ms.assetid: ce60e6d4-18e5-056c-229e-8c0b730109ae
ms.date: 06/08/2017
---


# Application.Goto Method (Excel)

Selects any range or Visual Basic procedure in any workbook, and activates that workbook if it?s not already active.


## Syntax

 _expression_ . **Goto**( **_Reference_** , **_Scroll_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Reference_|Optional| **Variant**|The destination. Can be a  **[Range](range-object-excel.md)** object, a string that contains a cell reference in R1C1-style notation, or a string that contains a Visual Basic procedure name. If this argument is omitted, the destination is the last range you used the **Goto** method to select.|
| _Scroll_|Optional| **Variant**| **True** to scroll through the window so that the upper-left corner of the range appears in the upper-left corner of the window. **False** to not scroll through the window. The default is **False** .|

## Remarks

This method differs from the  **[Select](range-select-method-excel.md)** method in the following ways:


- If you specify a range on a sheet that?s not on top, Microsoft Excel will switch to that sheet before selecting. (If you use  **Select** with a range on a sheet that?s not on top, the range will be selected but the sheet won?t be activated).
    
- This method has a  **_Scroll_** argument that lets you scroll through the destination window.
    
- When you use the  **Goto** method, the previous selection (before the **Goto** method runs) is added to the array of previous selections (for more information, see the **[PreviousSelections](application-previousselections-property-excel.md)** property). You can use this feature to quickly jump between as many as four selections.
    
- The  **Select** method has a _Replace_ argument; the **Goto** method doesn?t.
    

## Example

This example selects cell A154 on Sheet1 and then scrolls through the worksheet to display the range.


```vb
Application.Goto Reference:=Worksheets("Sheet1").Range("A154"), _ 
 scroll:=True
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

