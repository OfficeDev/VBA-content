---
title: ParagraphFormat.KashidaPercentage Property (Publisher)
keywords: vbapb10.chm5439513
f1_keywords:
- vbapb10.chm5439513
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.KashidaPercentage
ms.assetid: d62aa512-cce6-2e78-657f-51ff1b2cbcf8
ms.date: 06/08/2017
---


# ParagraphFormat.KashidaPercentage Property (Publisher)

Returns or sets a  **Long** indicating the percentage by which kashidas are to be lengthened for the specified paragraphs. Valid values are from 0 to 100. Read/write.


## Syntax

 _expression_. **KashidaPercentage**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Long


## Remarks

The  **[Alignment](paragraphformat-alignment-property-publisher.md)** property of the specified paragraphs must be set to **pbParagraphAlignmentKashida** or the **KashidaPercentage** property is ignored.


## Example

The following example sets the paragraphs in shape one on page one of the active publication to kashida alignment and specifies that kashidas are to be lengthened by 20 percent.


```vb
With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
 .Alignment = pbParagraphAlignmentKashida 
 .KashidaPercentage = 20 
End With
```


