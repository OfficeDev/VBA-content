---
title: LineNumbering.CountBy Property (Word)
keywords: vbawd10.chm158466151
f1_keywords:
- vbawd10.chm158466151
ms.prod: word
api_name:
- Word.LineNumbering.CountBy
ms.assetid: 7cb90bfb-84a9-d52f-f406-7bef835744d3
ms.date: 06/08/2017
---


# LineNumbering.CountBy Property (Word)

Returns or sets the numeric increment for line numbers. Read/write  **Long** .


## Syntax

 _expression_ . **CountBy**

 _expression_ A variable that represents a **[LineNumbering](linenumbering-object-word.md)** object.


## Remarks

If the  **CountBy** property is set to 5, every fifth line will display the line number. Line numbers are only displayed in print layout view and print preview. This property has no effect unless the **[Active](linenumbering-active-property-word.md)** property of the **LineNumbering** object is set to **True** .


## Example

This example turns on line numbering for the active document. The line number is displayed on every fifth line and line numbering starts over for each new section.


```vb
With ActiveDocument.PageSetup.LineNumbering 
 .Active = True 
 .CountBy = 5 
 .RestartMode = wdRestartSection 
End With
```


## See also


#### Concepts


[LineNumbering Object](linenumbering-object-word.md)

