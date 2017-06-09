---
title: ParagraphFormat.TextDirection Property (Publisher)
keywords: vbapb10.chm5439507
f1_keywords:
- vbapb10.chm5439507
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.TextDirection
ms.assetid: b96c634d-0e7e-dba8-2bf4-e5baf3afa3d1
ms.date: 06/08/2017
---


# ParagraphFormat.TextDirection Property (Publisher)

Returns or sets a  **PbTextDirection** constant indicating the direction in which text flows in the specified paragraph. Read/write.


## Syntax

 _expression_. **TextDirection**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

PbTextDirection


## Remarks

This property is meant to be used in conjunction with documents that have text in both left-to-right and right-to-left languages. Setting the property to a value that is not in accordance with the text direction dictated by the language in use may have unpredictable results.

The  **TextDirection** property value can be one of the **PbTextDirection** constants declared in the Microsoft Publisher type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **pbTextDirectionLeftToRight**| Text flows from left to right.|
| **pbTextDirectionMixed**|Return value indicating a range containing some left-to-right text and some right-to-left text.|
| **pbTextDirectionRightToLeft**|Text flows from right to left.|

## Example

The following example changes the text direction of the first shape on page one so that it flows from right-to-left.


```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.TextDirection = pbTextDirectionRightToLeft
```


