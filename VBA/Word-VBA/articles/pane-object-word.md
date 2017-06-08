---
title: Pane Object (Word)
keywords: vbawd10.chm2400
f1_keywords:
- vbawd10.chm2400
ms.prod: word
api_name:
- Word.Pane
ms.assetid: 4a0c2690-d9d2-4e34-fef4-cc41365f5251
ms.date: 06/08/2017
---


# Pane Object (Word)

Represents a window pane. The  **Pane** object is a member of the **Panes** collection. The **[Panes](panes-object-word.md)** collection includes all the window panes for a single window.


## Remarks

Use  **Panes** (Index), where Index is the index number, to return a single **Pane** object. The following example closes the active pane.


```vb
If ActiveDocument.ActiveWindow.Panes.Count >= 2 Then _ 
 ActiveDocument.ActiveWindow.ActivePane.Close
```

Use the  **Add** method or the **Split** property to add a window pane. The following example splits the active window at 20 percent of the current window size.




```vb
ActiveDocument.ActiveWindow.Panes.Add SplitVertical:=20
```

The following example splits the active window in half.




```vb
ActiveDocument.ActiveWindow.Split = True
```

You can use the  **SplitSpecial** property to show comments, footnotes, or endnotes in a separate pane.

A window has more than one pane if the window is split or the view is not print layout view and information such as footnotes or comments are displayed. The following example displays the comments pane in normal view and then prompts to close the pane.




```vb
ActiveDocument.ActiveWindow.View.Type = wdNormalView 
If ActiveDocument.Comments.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneComments 
 response = _ 
 MsgBox("Do you want to close the comments pane?", vbYesNo) 
 If response = vbYes Then _ 
 ActiveDocument.ActiveWindow.ActivePane.Close 
End If
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


