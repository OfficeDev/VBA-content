---
title: TabStop Object (PowerPoint)
keywords: vbapp10.chm574000
f1_keywords:
- vbapp10.chm574000
ms.prod: powerpoint
api_name:
- PowerPoint.TabStop
ms.assetid: 73be0eee-d42e-fa84-416d-0ecd30c9c2c3
ms.date: 06/08/2017
---


# TabStop Object (PowerPoint)

Represents a single tab stop. The  **TabStop** object is a member of the **[TabStops](tabstops-object-powerpoint.md)** collection. The **TabStops** collection represents all the tab stops on one ruler.


## Example

Use  **TabStops** (index), where index is the tab stop index number, to return a single **TabStop** object. The following example clears tab stop one for the text in shape two on slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Shapes(2).TextFrame _
    .Ruler.TabStops(1).Clear
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

