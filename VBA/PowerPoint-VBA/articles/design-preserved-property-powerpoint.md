---
title: Design.Preserved Property (PowerPoint)
keywords: vbapp10.chm644009
f1_keywords:
- vbapp10.chm644009
ms.prod: powerpoint
api_name:
- PowerPoint.Design.Preserved
ms.assetid: c7620e5a-49f5-49bc-307b-230ead112cf6
ms.date: 06/08/2017
---


# Design.Preserved Property (PowerPoint)

Represents whether a design master is preserved from changes. Read/write.


## Syntax

 _expression_. **Preserved**

 _expression_ A variable that represents a **Design** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Preserved** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The design master is not preserved and can be edited.|
|**msoTrue**| The design master is preserved and cannot be edited.|

## Example

The following line of code locks and preserves the first design master.


```vb
Sub PreserveMaster

    ActivePresentation.Designs(1).Preserved = msoTrue

End Sub
```


## See also


#### Concepts


[Design Object](design-object-powerpoint.md)

