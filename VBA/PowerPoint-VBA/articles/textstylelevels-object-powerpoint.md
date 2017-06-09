---
title: TextStyleLevels Object (PowerPoint)
keywords: vbapp10.chm580000
f1_keywords:
- vbapp10.chm580000
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyleLevels
ms.assetid: dc61e97f-e92e-d214-47af-5830c14b1b09
ms.date: 06/08/2017
---


# TextStyleLevels Object (PowerPoint)

A collection of all the outline text levels. This collection always contains five members, each of which is represented by a  **[TextStyleLevel](textstylelevel-object-powerpoint.md)** object.


## Example

Use  **Levels** (index), where index is a number from 1 through 5 that corresponds to the outline level, to return a single **TextStyleLevel** object. The following example sets the font name and font size for level-one body text on all the slides in the active presentation.


```vb
With ActivePresentation.SlideMaster _
        .TextStyles(ppBodyStyle).Levels(1)
    With .Font
        .Name = "Arial"
        .Size = 36
    End With
End With
```

The following example sets the font size for text at each outline level for the notes body area on all the notes pages in the active presentation.




```vb
With ActivePresentation.NotesMaster.TextStyles(ppBodyStyle).Levels

    .Item(1).Font.Size = 34

    .Item(2).Font.Size = 30

    .Item(3).Font.Size = 25

    .Item(4).Font.Size = 20

    .Item(5).Font.Size = 15

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

