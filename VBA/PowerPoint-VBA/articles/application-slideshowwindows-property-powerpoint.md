---
title: Application.SlideShowWindows Property (PowerPoint)
keywords: vbapp10.chm502006
f1_keywords:
- vbapp10.chm502006
ms.prod: powerpoint
api_name:
- PowerPoint.Application.SlideShowWindows
ms.assetid: 4beed51c-bb67-6208-c2b1-f1d5b6425d9b
ms.date: 06/08/2017
---


# Application.SlideShowWindows Property (PowerPoint)

Returns a  **[SlideShowWindows](slideshowwindows-object-powerpoint.md)** collection that represents all open slide show windows. Read-only.


## Syntax

 _expression_. **SlideShowWindows**

 _expression_ A variable that represents a **Application** object.


### Return Value

SlideShowWindows


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](return-objects-from-collections.md).


## Example

This example runs a slide show in a window and sets the height and width of the slide show window.


```vb
With Application

    .Presentations(1).SlideShowSettings.Run

    With .SlideShowWindows(1)

        .Height = 250

        .Width = 250

    End With

End With
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

