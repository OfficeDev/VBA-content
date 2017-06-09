---
title: ValidationIssue.TargetShape Property (Visio)
keywords: vis_sdr.chm18662685
f1_keywords:
- vis_sdr.chm18662685
ms.prod: visio
api_name:
- Visio.ValidationIssue.TargetShape
ms.assetid: 93cc256d-6763-064c-392e-46232033b6dc
ms.date: 06/08/2017
---


# ValidationIssue.TargetShape Property (Visio)

Returns the  **[Shape](shape-object-visio.md)** object that is associated with the validation issue. Read-only.


## Syntax

 _expression_ . **TargetShape**

 _expression_ A variable that represents a **[ValidationIssue](validationissue-object-visio.md)** object.


### Return Value

 **Shape**


## Remarks

 If the error is not associated with a specific shape, the **TargetShape** property returns **Nothing** .


