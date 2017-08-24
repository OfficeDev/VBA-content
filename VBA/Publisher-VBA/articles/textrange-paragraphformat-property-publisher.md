---
title: TextRange.ParagraphFormat Property (Publisher)
keywords: vbapb10.chm5308439
f1_keywords:
- vbapb10.chm5308439
ms.prod: publisher
api_name:
- Publisher.TextRange.ParagraphFormat
ms.assetid: 475da411-9292-a12d-addd-1bbe822ec09e
ms.date: 06/08/2017
---


# TextRange.ParagraphFormat Property (Publisher)

Returns a  **[ParagraphFormat](paragraphformat-object-publisher.md)** object representing the paragraph formatting for the specified text range or text style.


## Syntax

 _expression_. **ParagraphFormat**

 _expression_A variable that represents a  **TextRange** object.


## Example

The following example removes all the tab stops from the text in the first shape on page one of the active publication.


```vb
Dim pfTemp As ParagraphFormat 
 
Set pfTemp = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
 
pfTemp.Tabs.ClearAll
```


