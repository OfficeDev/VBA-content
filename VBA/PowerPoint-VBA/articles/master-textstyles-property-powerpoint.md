---
title: Master.TextStyles Property (PowerPoint)
keywords: vbapp10.chm533011
f1_keywords:
- vbapp10.chm533011
ms.prod: powerpoint
api_name:
- PowerPoint.Master.TextStyles
ms.assetid: 713b6f60-5c20-6ddf-9660-4f5f2d27546d
ms.date: 06/08/2017
---


# Master.TextStyles Property (PowerPoint)

Returns a  **[TextStyles](textstyles-object-powerpoint.md)** collection that represents three text styles — title text, body text, and default text — for the specified slide master. Read-only.


## Syntax

 _expression_. **TextStyles**

 _expression_ A variable that represents a **Master** object.


### Return Value

TextStyles


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](return-objects-from-collections.md).


## Example

This example sets the font name and font size for level-one body text on slides in the active presentation.


```vb
With ActivePresentation.SlideMaster_

        .TextStyles(ppBodyStyle).Levels(1)

    With .Font

        .Name = "arial"

        .Size = 36

    End With

End With
```


## See also


#### Concepts


[Master Object](master-object-powerpoint.md)

