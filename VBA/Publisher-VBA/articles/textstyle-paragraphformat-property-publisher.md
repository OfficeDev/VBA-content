---
title: TextStyle.ParagraphFormat Property (Publisher)
keywords: vbapb10.chm5963781
f1_keywords:
- vbapb10.chm5963781
ms.prod: publisher
api_name:
- Publisher.TextStyle.ParagraphFormat
ms.assetid: 5ab0a2ec-d7a9-f3af-29e7-5421427ee783
ms.date: 06/08/2017
---


# TextStyle.ParagraphFormat Property (Publisher)

Returns a  **[ParagraphFormat](paragraphformat-object-publisher.md)** object representing the paragraph formatting for the specified text range or text style.


## Syntax

 _expression_. **ParagraphFormat**

 _expression_A variable that represents a  **TextStyle** object.


## Example

The following example removes all the tab stops from the text in the first shape on page one of the active publication.


```vb
Dim pfTemp As ParagraphFormat 
 
Set pfTemp = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
 
pfTemp.Tabs.ClearAll
```


