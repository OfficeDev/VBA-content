---
title: View.ShowHeading Method (Word)
keywords: vbawd10.chm161808488
f1_keywords:
- vbawd10.chm161808488
ms.prod: word
api_name:
- Word.View.ShowHeading
ms.assetid: b459e936-13fa-f2f4-33e6-f25d21a6f77c
ms.date: 06/08/2017
---


# View.ShowHeading Method (Word)

Shows all headings up to the specified heading level and hides subordinate headings and body text.


## Syntax

 _expression_ . **ShowHeading**( **_Level_** )

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Level_|Required| **Long**|The outline heading level (a number from 1 to 9).|

## Remarks

This method generates an error if the view isn't outline view or master document view.


## Example

This example switches the active window to outline view and displays all text that's formatted with the Heading 1 style. Body text and all other types of headings are hidden.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdOutlineView 
 .ShowHeading 1 
End With
```

This example switches the window for Document1 to outline view and displays all text that's formatted with the Heading 1, Heading 2, or Heading 3 style.




```vb
With Windows("Document1").View 
 .Type = wdOutlineView 
 .ShowHeading 3 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

