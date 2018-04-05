---
title: Line.LineSlant Property (Access)
keywords: vbaac10.chm10329
f1_keywords:
- vbaac10.chm10329
ms.prod: access
api_name:
- Access.Line.LineSlant
ms.assetid: 336f66fe-2b15-f3d0-6cf2-5b48992ddafc
ms.date: 06/08/2017
---


# Line.LineSlant Property (Access)

You use the  **LineSlant** property to specify whether a[line control](line-control.md)slants from upper left to lower right or from upper right to lower left. Read/write  **Boolean**.


## Syntax

 _expression_. **LineSlant**

 _expression_ A variable that represents a **Line** object.


## Remarks

The  **LineSlant** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|**\**|**False**|(Default) Upper left to lower right|
|**/**|**True**|Upper right to lower left|
Use the  **LineSlant** property to change a line's direction. To position and size the line on your form or report, use the mouse.


## Example

The following example slants a line on a form from upper right to lower left.


```vb
Forms("Purchase Orders").Controls("Section Separator").LineSlant = True 

```


## See also


#### Concepts


[Line Object](line-object-access.md)

