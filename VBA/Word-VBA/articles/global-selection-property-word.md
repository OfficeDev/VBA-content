---
title: Global.Selection Property (Word)
keywords: vbawd10.chm163119109
f1_keywords:
- vbawd10.chm163119109
ms.prod: word
api_name:
- Word.Global.Selection
ms.assetid: 71938a78-36ae-07ba-496b-911bef746444
ms.date: 06/08/2017
---


# Global.Selection Property (Word)

Returns a  **Selection** object that represents a selected range or the insertion point. Read-only.


## Syntax

 _expression_ . **Selection**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


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


[Global Object](global-object-word.md)

