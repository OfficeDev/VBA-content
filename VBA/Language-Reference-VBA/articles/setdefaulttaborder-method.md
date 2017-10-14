---
title: SetDefaultTabOrder Method
keywords: fm20.chm2000420
f1_keywords:
- fm20.chm2000420
ms.prod: office
api_name:
- Office.SetDefaultTabOrder
ms.assetid: fd4661ee-a995-1872-509b-edffa6dbbf80
ms.date: 06/08/2017
---


# SetDefaultTabOrder Method



Sets the  **TabIndex** property of each control on a form, using a default top-to-bottom, left-to-right[tab order](vbe-glossary.md).
 **Syntax**
 _object_. **SetDefaultTabOrder**
The  **SetDefaultTabOrder** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
Microsoft Forms sets the tab order beginning with controls in the upper left corner of the form and moving to the right. It places controls closest to the left edge of the form earlier in the tab order. If more than one control is the same distance from the left edge of the form, tab order values are assigned from top to bottom.

