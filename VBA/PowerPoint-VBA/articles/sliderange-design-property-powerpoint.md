---
title: SlideRange.Design Property (PowerPoint)
keywords: vbapp10.chm532033
f1_keywords:
- vbapp10.chm532033
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Design
ms.assetid: 7960f99a-fa5a-1ba0-e39a-fe3afe579621
ms.date: 06/08/2017
---


# SlideRange.Design Property (PowerPoint)

Returns a  **Design** object representing a design.


## Syntax

 _expression_. **Design**

 _expression_ A variable that represents a **SlideRange** object.


### Return Value

Design


## Example

The following example adds a title master.


```vb
Sub AddDesignMaster

    ActivePresentation.Slides(1).Design.AddTitleMaster

End Sub
```


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

