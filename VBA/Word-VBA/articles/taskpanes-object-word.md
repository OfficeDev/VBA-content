---
title: TaskPanes Object (Word)
ms.prod: word
api_name:
- Word.TaskPanes
ms.assetid: a560a41b-a1d7-175a-b475-af742c9fa1f8
ms.date: 06/08/2017
---


# TaskPanes Object (Word)

A collection of  **TaskPane** objects that contains commonly performed tasks in Microsoft Word.


## Remarks

Use the  **TaskPanes** property to return the **TaskPanes** collection. Use the **Item** method with a **[WdTaskPanes](wdtaskpanes-enumeration-word.md)** constant to refer to a specific task pane. The example below displays the formatting task pane.


```vb
Sub FormattingPane() 
 Application.TaskPanes(wdTaskPaneFormatting).Visible = True 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


