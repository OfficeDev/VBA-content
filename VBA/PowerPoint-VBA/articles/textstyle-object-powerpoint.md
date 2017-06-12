---
title: TextStyle Object (PowerPoint)
keywords: vbapp10.chm579000
f1_keywords:
- vbapp10.chm579000
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyle
ms.assetid: 59cf79e2-7212-4928-d966-6340c9021a6d
ms.date: 06/08/2017
---


# TextStyle Object (PowerPoint)

Represents one of three text styles: title text, body text, or default text. Each text style contains a  **[TextFrame](textframe-object-powerpoint.md)** object that describes how text is placed within the text bounding box, a **[Ruler](ruler-object-powerpoint.md)** object that contains tab stops and outline indent formatting information, and a **[TextStyleLevels](textstylelevels-object-powerpoint.md)** collection that contains outline text formatting information. The **TextStyle** object is a member of the **[TextStyles](textstyles-object-powerpoint.md)** collection.


## Example

Use  **TextStyles** (index), where index is either **ppBodyStyle**, **ppDefaultStyle**, or **ppTitleStyle**, to return a single **TextStyle** object. The following example sets the font name and font size for level-one body text on all the slides in the active presentation.


```vb
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Levels(1)
    With .Font
        .Name = "Arial"
        .Size = 36
    End With
End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

