---
title: Document.SnapExtensions Property (Visio)
keywords: vis_sdr.chm10550885
f1_keywords:
- vis_sdr.chm10550885
ms.prod: visio
api_name:
- Visio.Document.SnapExtensions
ms.assetid: 8b5aad7a-335a-dc8c-aa58-42947ffdc53e
ms.date: 06/08/2017
---


# Document.SnapExtensions Property (Visio)

Determines the shape extensions that are active in a document. Read/write.


## Syntax

 _expression_ . **SnapExtensions**

 _expression_ A variable that represents a **Document** object.


### Return Value

VisSnapExtensions


## Remarks

You can also set this value by checking options in the  **Shape extension options** box on the **Advanced** tab of the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab).

The  **SnapExtensions** property can be any combination of the following **VisSnapExtensions** constants, which are declared in the Visio type library. The default is to show center axes and linear extensions, &;H22 (34).



|**Constant **|**Value **|
|:-----|:-----|
| **visSnapExtNone**|&;H0|
| **visSnapExtAlignmentBoxExtension**|&;H1|
| **visSnapExtCenterAxes**|&;H2|
| **visSnapExtCurveTangent**|&;H4|
| **visSnapExtEndpoint**|&;H8|
| **visSnapExtMidpoint**|&;H10|
| **visSnapExtLinearExtension**|&;H20|
| **visSnapExtCurveExtension**|&;H40|
| **visSnapExtEndpointPerpendicular**|&;H80|
| **visSnapExtMidpointPerpendicular**|&;H100|
| **visSnapExtEndpointHorizontal**|&;H200|
| **visSnapExtEndpointVertical**|&;H400|
| **visSnapExtEllipseCenter**|&;H800|
| **visSnapExtIsometricAngles**|&;H1000|

