---
title: ToolbarItem.TypeSpecific2 Property (Visio)
keywords: vis_sdr.chm13514605
f1_keywords:
- vis_sdr.chm13514605
ms.prod: visio
api_name:
- Visio.ToolbarItem.TypeSpecific2
ms.assetid: cdd33e96-bb18-6476-ccac-70797d2df4c1
ms.date: 06/08/2017
---


# ToolbarItem.TypeSpecific2 Property (Visio)

Gets or sets the type of a toolbar item. Read/write.


## Syntax

 _expression_ . **TypeSpecific2**

 _expression_ A variable that represents a **ToolbarItem** object.


### Return Value

Integer


## Remarks

The value of an object's  **TypeSpecific2** property depends on the value of its **CntrlType** property.



|**CntrlType value **|**TypeSpecific1 value **|
|:-----|:-----|
| **visCtrlTypeBUTTON**|The  **TypeSpecific2** property is not used.|
| **visCtrlTypeCOMBOBOX**|The current width of the control expressed in pixels.|
| **visCtrlTypeEDITBOX**|The current width of the control expressed in pixels.|
| **visCtrlTypeLABEL**|The  **TypeSpecific2** property is not used.|

