---
title: View.Zoom Property (PowerPoint)
keywords: vbapp10.chm512004
f1_keywords:
- vbapp10.chm512004
ms.prod: powerpoint
api_name:
- PowerPoint.View.Zoom
ms.assetid: 83624f62-0da8-ad96-d887-7f87cb4cacd2
ms.date: 06/08/2017
---


# View.Zoom Property (PowerPoint)

Returns or sets the zoom setting of the specified view as a percentage of normal size. Read/write.


## Syntax

 _expression_. **Zoom**

 _expression_ A variable that represents a **View** object.


### Return Value

Integer


## Remarks

The  **Zoom** property value can be from 10 to 400 percent.


## Example

The following example sets the zoom to 30 percent for the view in document window one.


```
Windows(1).View.Zoom = 30
```


## See also


#### Concepts


[View Object](view-object-powerpoint.md)

