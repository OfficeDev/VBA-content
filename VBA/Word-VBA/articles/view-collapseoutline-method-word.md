---
title: View.CollapseOutline Method (Word)
keywords: vbawd10.chm161808485
f1_keywords:
- vbawd10.chm161808485
ms.prod: word
api_name:
- Word.View.CollapseOutline
ms.assetid: b22ac567-ef40-e47e-f0fc-311263675045
ms.date: 06/08/2017
---


# View.CollapseOutline Method (Word)

Collapses the text under the selection or the specified range by one heading level.


## Syntax

 _expression_ . **CollapseOutline**( **_Range_** )

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Optional| **Range object**|The range of paragraphs to be collapsed. If this argument is omitted, the entire selection is collapsed.|

## Remarks

If the document isn't in outline or master document view, an error occurs.


## Example

This example applies the Heading 2 style to the second paragraph in the active document, switches the active window to outline view, and collapses the text under the second paragraph in the document.


```vb
ActiveDocument.Paragraphs(2).Style = wdStyleHeading2 
With ActiveDocument.ActiveWindow.View 
 .Type = wdOutlineView 
 .CollapseOutline Range:=ActiveDocument.Paragraphs(2).Range 
End With
```

This example collapses every heading in the document by one level.




```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdOutlineView 
 .CollapseOutline Range:=ActiveDocument.Content 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

