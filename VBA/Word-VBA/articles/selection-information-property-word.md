---
title: Selection.Information Property (Word)
keywords: vbawd10.chm158663057
f1_keywords:
- vbawd10.chm158663057
ms.prod: word
api_name:
- Word.Selection.Information
ms.assetid: 73028751-6339-47e6-9629-9584cc4c59ec
ms.date: 06/08/2017
---


# Selection.Information Property (Word)

Returns information about the specified selection. Read-only  **Variant** .


## Syntax

 _expression_ . **Information**( **_Type_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **WdInformation**|The information type.|

## Example

This example displays the current page number and the total number of pages in the active document.


```vb
MsgBox "The selection is on page " &; _ 
 Selection.Information(wdActiveEndPageNumber) &; " of page " _ 
 &; Selection.Information(wdNumberOfPagesInDocument)
```

If the selection is in a table, this example selects the table.




```vb
If Selection.Information(wdWithInTable) Then _ 
 Selection.Tables(1).Select
```

This example displays a message that indicates the current section number.




```
Selection.Collapse Direction:=wdCollapseStart 
MsgBox "The insertion point is in section " &; _ 
 Selection.Information(wdActiveEndSectionNumber)
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

