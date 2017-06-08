---
title: Window.DisplayRulers Property (Word)
keywords: vbawd10.chm157417484
f1_keywords:
- vbawd10.chm157417484
ms.prod: word
api_name:
- Word.Window.DisplayRulers
ms.assetid: 4e1f2dd1-641b-4fe7-c801-febba26372ec
ms.date: 06/08/2017
---


# Window.DisplayRulers Property (Word)

 **True** if rulers are displayed for the specified window or pane. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayRulers**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Remarks

This property is equivalent to the  **Ruler** command on the **View** menu. If **DisplayRulers** is **False** , the horizontal and vertical rulers won't be displayed, regardless of the state of the **DisplayVerticalRuler** property.


## Example

This example toggles the ruler display for the active window.


```vb
ActiveDocument.ActiveWindow.DisplayRulers = _ 
 Not ActiveDocument.ActiveWindow.DisplayRulers
```

This example switches the window to print layout view and displays the horizontal and vertical rulers.




```vb
With ActiveDocument.ActiveWindow 
 .View.Type = wdPrintView 
 .DisplayVerticalRuler = True 
 .DisplayRulers = True 
End With
```


## See also


#### Concepts


[Window Object](window-object-word.md)

