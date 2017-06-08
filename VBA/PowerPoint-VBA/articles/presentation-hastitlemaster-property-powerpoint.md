---
title: Presentation.HasTitleMaster Property (PowerPoint)
keywords: vbapp10.chm583005
f1_keywords:
- vbapp10.chm583005
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.HasTitleMaster
ms.assetid: 93b5932c-c03f-451a-c7f9-30683c01bcfa
ms.date: 06/08/2017
---


# Presentation.HasTitleMaster Property (PowerPoint)

 **MsoTrue** if the specified presentation has a title master. Read-only.


## Syntax

 _expression_. **HasTitleMaster**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HasTitleMaster** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified presentation does not have a title master.|
|**msoTrue**| The specified presentation has a title master.|

## Example

This example adds a title master to the active presentation if it doesn't already have one.


```vb
With Application.ActivePresentation

    If Not .HasTitleMaster Then .AddTitleMaster

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

