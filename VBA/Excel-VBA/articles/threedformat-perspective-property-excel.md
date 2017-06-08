---
title: ThreeDFormat.Perspective Property (Excel)
keywords: vbaxl10.chm119008
f1_keywords:
- vbaxl10.chm119008
ms.prod: excel
api_name:
- Excel.ThreeDFormat.Perspective
ms.assetid: 9f31508e-c723-e55a-07a9-cef1bc526136
ms.date: 06/08/2017
---


# ThreeDFormat.Perspective Property (Excel)

Returns or sets an  **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** value that determines whether the extrusion appears in perspective.


## Syntax

 _expression_ . **Perspective**

 _expression_ A variable that represents a **ThreeDFormat** object.


## Remarks

This property can be set to one of the following  **MsoTriState** constants:



| **msoCTrue** Does not apply to this property.|
| **msoFalse** The extrusion is a parallel, or orthographic, projection—that is, the walls don't narrow toward a vanishing point.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** The extrusion appears in perspective—that is, the walls of the extrusion narrow toward a vanishing point **.**|

## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

