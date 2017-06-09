---
title: Attachment.BorderWidth Property (Access)
keywords: vbaac10.chm13930
f1_keywords:
- vbaac10.chm13930
ms.prod: access
api_name:
- Access.Attachment.BorderWidth
ms.assetid: e72672a1-3b17-ad1b-ff7d-96e3652a9f35
ms.date: 06/08/2017
---


# Attachment.BorderWidth Property (Access)

You can use the  **BorderWidth** property to specify the width of a control's border. Read/write **Byte**.


## Syntax

 _expression_. **BorderWidth**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

The  **BorderWidth** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Hairline|0|(Default) The narrowest border possible on your system.|
|1 pt to 6 pt|1 to 6|The width as indicated in points.|
You can set the default for this property by using the control's default control style or the  **DefaultControl** property in Visual Basic.

To use the  **BorderWidth** property, the **SpecialEffect** property must be set to Flat or Shadowed and the **BorderStyle** property must not be set to Transparent. If the **SpecialEffect** property is set to any other value and/or the **BorderStyle** property is set to Transparent, and you set the **BorderWidth** property, the **SpecialEffect** property is automatically reset to Flat and the **BorderStyle** property is automatically reset to Solid.

The exact border width depends on your computer and printer. On some systems, the hairline and 1-point widths appear the same.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

