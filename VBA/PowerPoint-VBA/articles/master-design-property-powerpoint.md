---
title: Master.Design Property (PowerPoint)
keywords: vbapp10.chm533014
f1_keywords:
- vbapp10.chm533014
ms.prod: powerpoint
api_name:
- PowerPoint.Master.Design
ms.assetid: 78035fbd-e2f3-9089-2263-c04ce72394db
ms.date: 06/08/2017
---


# Master.Design Property (PowerPoint)

Returns a  **Design** object representing a design.


## Syntax

 _expression_. **Design**

 _expression_ A variable that represents a **Master** object.


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


[Master Object](master-object-powerpoint.md)

