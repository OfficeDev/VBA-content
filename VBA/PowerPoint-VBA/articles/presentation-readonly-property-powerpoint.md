---
title: Presentation.ReadOnly Property (PowerPoint)
keywords: vbapp10.chm583023
f1_keywords:
- vbapp10.chm583023
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.ReadOnly
ms.assetid: d0d69c81-baa0-9b33-5ee3-d8e581508a88
ms.date: 06/08/2017
---


# Presentation.ReadOnly Property (PowerPoint)

Returns whether the specified presentation is read-only. Read-only.


## Syntax

 _expression_. **ReadOnly**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

MsoTriState


## Remarks

The value of the  **ReadOnly** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The specified presentation is not read-only.|
|**msoTrue**| The specified presentation is read-only.|

## Example

If the active presentation is read-only, this example saves it as newfile.ppt.


```vb
With Application.ActivePresentation

    If .ReadOnly Then .SaveAs FileName:="newfile"

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

