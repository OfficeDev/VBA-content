---
title: View.NextHeaderFooter Method (Word)
keywords: vbawd10.chm161808490
f1_keywords:
- vbawd10.chm161808490
ms.prod: word
api_name:
- Word.View.NextHeaderFooter
ms.assetid: 48b52b41-cee4-fa85-7229-86af61607556
ms.date: 06/08/2017
---


# View.NextHeaderFooter Method (Word)

Moves to the next header or footer, depending on whether a header or footer is displayed in the view.


## Syntax

 _expression_ . **NextHeaderFooter**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.


## Remarks

If the view displays a header, this method moves to the next header within the current section (for example, from an odd header to an even header) or to the first header in the following section. If the view displays a footer, this method moves to the next footer. 


 **Note**  If the view displays the last header or footer in the last section of the document, or if it is not displaying a header or footer at all, an error occurs.


## Example

This example displays the first page header in the active document and then switches to the next header. The document needs to be at least two pages long.


```vb
ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = True 
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView =wdSeekFirstPageHeader 
 .NextHeaderFooter 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

