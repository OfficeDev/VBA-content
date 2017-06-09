---
title: Presentation.Saved Property (PowerPoint)
keywords: vbapp10.chm583027
f1_keywords:
- vbapp10.chm583027
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Saved
ms.assetid: 52798ca6-e181-cf82-d397-647404235cb9
ms.date: 06/08/2017
---


# Presentation.Saved Property (PowerPoint)

Determines whether changes have been made to a presentation since it was last saved. Read/write.


## Syntax

 _expression_. **Saved**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

MsoTriState


## Remarks

If the  **Saved** property of a modified presentation is set to **msoTrue**, the user won't be prompted to save changes when closing the presentation, and all changes made to it since it was last saved will be lost.

The value of the  **Saved** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Changes have been made to a presentation since it was last saved.|
|**msoTrue**| No changes have been made to a presentation since it was last saved.|

## Example

This example saves the active presentation if it is been changed since the last time it was saved.


```vb
With Application.ActivePresentation

    If Not .Saved And .Path <> "" Then .Save

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

