---
title: View.ShowAllHeadings Method (Word)
keywords: vbawd10.chm161808487
f1_keywords:
- vbawd10.chm161808487
ms.prod: word
api_name:
- Word.View.ShowAllHeadings
ms.assetid: 294aa5f0-9821-faed-aa82-ff59f7a84eb6
ms.date: 06/08/2017
---


# View.ShowAllHeadings Method (Word)

Switches between showing all text (headings and body text) and showing only headings.


## Syntax

 _expression_ . **ShowAllHeadings**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


## Remarks

This method generates an error if the view isn't outline view or master document view.


## Example

This example uses the  **ShowHeading** method to show all headings (without any body text) and then switches the display to show all text (headings and body text) in outline view.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdOutlineView 
 .ShowHeading 9 
 .ShowAllHeadings 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

