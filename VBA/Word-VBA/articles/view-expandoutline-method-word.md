---
title: View.ExpandOutline Method (Word)
keywords: vbawd10.chm161808486
f1_keywords:
- vbawd10.chm161808486
ms.prod: word
api_name:
- Word.View.ExpandOutline
ms.assetid: 46286501-3583-e931-71a6-cf5d091f0b15
ms.date: 06/08/2017
---


# View.ExpandOutline Method (Word)

Expands the text under the selection by one heading level.


## Syntax

 _expression_ . **ExpandOutline**( **_Range_** )

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Optional| **Range**|The range of paragraphs to be expanded. If this argument is omitted, the entire selection is expanded.|

## Remarks

If the document isn't in outline or master document view, an error occurs.


## Example

This example expands every heading in the document by one level.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdOutlineView 
 .ExpandOutline Range:=ActiveDocument.Content 
End With
```

This example expands the active paragraph in the Document2 window.




```vb
With Windows("Document2") 
 .Activate 
 .View.Type = wdOutlineView 
 .View.ExpandOutline 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

