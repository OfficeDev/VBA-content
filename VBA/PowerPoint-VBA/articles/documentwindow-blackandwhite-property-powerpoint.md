---
title: DocumentWindow.BlackAndWhite Property (PowerPoint)
keywords: vbapp10.chm511007
f1_keywords:
- vbapp10.chm511007
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.BlackAndWhite
ms.assetid: 1363b7df-8de5-955f-60a7-682cd6b4c848
ms.date: 06/08/2017
---


# DocumentWindow.BlackAndWhite Property (PowerPoint)

Determines whether the document window display is black and white. Read/write.


## Syntax

 _expression_. **BlackAndWhite**

 _expression_ A variable that represents a **DocumentWindow** object.


### Return Value

MsoTriState


## Remarks

The value of the  **BlackAndWhite** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. The document window display is not black and white. |
|**msoTrue**| The document window display is black and white.|

## Example

This example changes the display in window one to black and white.


```vb
Application.Windows(1).BlackAndWhite = msoTrue
```


## See also


#### Concepts


[DocumentWindow Object](documentwindow-object-powerpoint.md)


