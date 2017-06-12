---
title: TaskPane Object (Word)
keywords: vbawd10.chm1250
f1_keywords:
- vbawd10.chm1250
ms.prod: word
api_name:
- Word.TaskPane
ms.assetid: 57367e56-2de5-37bd-a9ca-f1fcb6b8c465
ms.date: 06/08/2017
---


# TaskPane Object (Word)

Represents a single task pane available to Microsoft Word, which contains common tasks that users perform. The  **TaskPane** object is a member of the **TaskPanes** collection.


## Remarks

Use the  **TaskPanes** property to return a **TaskPane** object. Use the **Visible** property to display an individual task pane. This example displays the formatting task pane.


```vb
Sub FormattingPane() 
 Application.TaskPanes(wdTaskPaneFormatting).Visible = True 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


