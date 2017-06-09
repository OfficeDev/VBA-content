---
title: PageNumbers.RestartNumberingAtSection Property (Word)
keywords: vbawd10.chm159776774
f1_keywords:
- vbawd10.chm159776774
ms.prod: word
api_name:
- Word.PageNumbers.RestartNumberingAtSection
ms.assetid: b960fc7d-78f7-ec23-d429-7ee9820e8340
ms.date: 06/08/2017
---


# PageNumbers.RestartNumberingAtSection Property (Word)

 **True** if page numbering starts at 1 again at the beginning of the specified section. Read/write **Boolean** .


## Syntax

 _expression_ . **RestartNumberingAtSection**

 _expression_ An expression that returns a **[PageNumbers](pagenumbers-object-word.md)** collection object.


## Remarks

If set to  **False** , the **RestartNumberingAtSection** property will override the **[StartingNumber](pagenumbers-startingnumber-property-word.md)** property so that page numbering can continue from the previous section.


## Example

This example adds page numbers to the headers in the active document, and then it sets page numbering to start at 1 again at the beginning of each section.


```vb
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary) _ 
 .PageNumbers.Add Pagenumberalignment:=wdAlignPageNumberCenter 
For Each s In ActiveDocument.Sections 
 With s.Headers(wdHeaderFooterPrimary).PageNumbers 
 .RestartNumberingAtSection = True 
 .StartingNumber = 1 
 End With 
Next s
```


## See also


#### Concepts


[PageNumbers Collection Object](pagenumbers-object-word.md)

