---
title: Paragraph.TabStops Property (Word)
keywords: vbawd10.chm156697679
f1_keywords:
- vbawd10.chm156697679
ms.prod: word
api_name:
- Word.Paragraph.TabStops
ms.assetid: e1739724-c236-e934-4e10-512d19cb8989
ms.date: 06/08/2017
---


# Paragraph.TabStops Property (Word)

Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraph. Read/write.


## Syntax

 _expression_ . **TabStops**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the tab stops for every paragraph in the document to match the tab stops in the first paragraph.


```vb
Set para1Tabs = ActiveDocument.Paragraphs(1).TabStops 
ActiveDocument.Paragraphs.TabStops = para1Tabs
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

