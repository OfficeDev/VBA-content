---
title: Section.PageSetup Property (Word)
keywords: vbawd10.chm156828749
f1_keywords:
- vbawd10.chm156828749
ms.prod: word
api_name:
- Word.Section.PageSetup
ms.assetid: ef198acd-1bb6-8e9b-64db-b162ad61f8c1
ms.date: 06/08/2017
---


# Section.PageSetup Property (Word)

Returns a  **PageSetup** object that is associated with the specified section.


## Syntax

 _expression_ . **PageSetup**

 _expression_ A variable that represents a **[Section](section-object-word.md)** object.


## Example

This example sets the gutter for the first section in Summary.doc to 36 points (0.5 inch).


```
Documents("Summary.doc").Sections(1).PageSetup.Gutter = 36
```


## See also


#### Concepts


[Section Object](section-object-word.md)

