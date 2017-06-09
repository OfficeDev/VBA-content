---
title: Hyperlink.ShowAndReturn Property (PowerPoint)
keywords: vbapp10.chm526010
f1_keywords:
- vbapp10.chm526010
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.ShowAndReturn
ms.assetid: 5d08a3ff-8352-0523-2d8c-629f996b296a
ms.date: 06/08/2017
---


# Hyperlink.ShowAndReturn Property (PowerPoint)

Determines if and under what circumstances Microsoft PowerPoint returns to the initiating slide show. Read/write.


## Syntax

 _expression_. **ShowAndReturn**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

MsoTriState


## Remarks

The value of the  **ShowAndReturn** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Default. PowerPoint doesn't return to the initiating slide show from the deactivated custom slide show.|
|**msoTrue**| PowerPoint returns to the initiating slide show from a deactivated custom slide show that was activated by using the **[Hyperlink](hyperlink-object-powerpoint.md)** object of the initiating presentation.|

## See also


#### Concepts


[Hyperlink Object](hyperlink-object-powerpoint.md)

