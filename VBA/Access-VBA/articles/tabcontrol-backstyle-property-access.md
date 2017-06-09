---
title: TabControl.BackStyle Property (Access)
keywords: vbaac10.chm12085
f1_keywords:
- vbaac10.chm12085
ms.prod: access
api_name:
- Access.TabControl.BackStyle
ms.assetid: f20c9406-7608-89ea-d0ff-d54bf4d1c526
ms.date: 06/08/2017
---


# TabControl.BackStyle Property (Access)

You can use the  **BackStyle** property to specify whether a control will be transparent. Read/write **Byte**.


## Syntax

 _expression_. **BackStyle**

 _expression_ A variable that represents a **TabControl** object.


## Remarks

The  **BackStyle** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Normal|1|(Default for all controls except option group) The control has its interior color set by the BackColor property.|
|Transparent|0|(Default for option group) The control is transparent. The color of the form or report behind the control is visible.|
You can set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

If the  **Transparent** button on the **Back Color** button palette is selected, the **BackStyle** property is set to Transparent; otherwise the **BackStyle** property is set to Normal.

To make a command button invisible, set its  **Transparent** property to Yes.


## See also


#### Concepts


[TabControl Object](tabcontrol-object-access.md)

