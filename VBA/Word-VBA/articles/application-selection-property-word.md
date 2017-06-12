---
title: Application.Selection Property (Word)
keywords: vbawd10.chm158334981
f1_keywords:
- vbawd10.chm158334981
ms.prod: word
api_name:
- Word.Application.Selection
ms.assetid: d2362378-06a1-3a1a-2bd0-358f190eb6f3
ms.date: 06/08/2017
---


# Application.Selection Property (Word)

Returns the  **[Selection](selection-object-word.md)** object that represents a selected range or the insertion point. Read-only.


## Syntax

 _expression_ . **Selection**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example displays the selected text.


```vb
If Selection.Type = wdSelectionNormal Then MsgBox Selection.Text
```

This example applies the Arial font and bold formatting to the selection.




```vb
With Selection.Font 
 .Bold = True 
 .Italic = False 
 .Name = "Arial" 
End With
```

If the insertion point isn't located in a table, the selection is moved to the next table.




```vb
If Selection.Information(wdWithInTable) = False Then 
 Selection.GoToNext What:=wdGoToTable 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

