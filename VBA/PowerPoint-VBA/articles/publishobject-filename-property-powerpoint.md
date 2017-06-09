---
title: PublishObject.FileName Property (PowerPoint)
keywords: vbapp10.chm635009
f1_keywords:
- vbapp10.chm635009
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.FileName
ms.assetid: 21bb55c1-1e0c-5ca5-5b44-668642b013a9
ms.date: 06/08/2017
---


# PublishObject.FileName Property (PowerPoint)

Returns or sets the path and file name of the Web presentation created when all or part of the active presentation is published. Read/write.


## Syntax

 _expression_. **FileName**

 _expression_ A variable that represents a **PublishObject** object.


### Return Value

String


## Remarks

The  **FileName** property generates an error if a folder in the specified path does not exist.


## Example

This example publishes slides three through five of the active presentation to HTML. It names the published presentation Mallard.htm and saves it in the Test folder.


```vb
With ActivePresentation.PublishObjects(1)

    .FileName = "C:\Test\Mallard.htm"

    .SourceType = ppPublishSlideRange

    .RangeStart = 3

    .RangeEnd = 5

    .Publish

End With
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-powerpoint.md)

