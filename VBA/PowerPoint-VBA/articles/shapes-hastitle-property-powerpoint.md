---
title: Shapes.HasTitle Property (PowerPoint)
keywords: vbapp10.chm543018
f1_keywords:
- vbapp10.chm543018
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.HasTitle
ms.assetid: 0754bda8-7e19-6dd1-55a3-2b19541480b9
ms.date: 06/08/2017
---


# Shapes.HasTitle Property (PowerPoint)

Returns whether the collection of objects on the specified slide contains a title placeholder. Read-only.


## Syntax

 _expression_. **HasTitle**

 _expression_ A variable that represents a **Shapes** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HasTitle** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The collection of objects on the specified slide does not contain a title placeholder.|
|**msoTrue**| The collection of objects on the specified slide contains a title placeholder.|

## Example

This example restores the title placeholder to slide one in the active presentation if this placeholder has been deleted. The text of the restored title is "Restored title."


```vb
With ActivePresentation.Slides(1)

    If .Layout <> ppLayoutBlank Then
        With .Shapes
            If Not .HasTitle Then
                .AddTitle.TextFrame.TextRange _
                    .Text = "Restored title"
            End If
        End With
    End If

End With
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

