---
title: Attachment.Left Property (Access)
keywords: vbaac10.chm13920
f1_keywords:
- vbaac10.chm13920
ms.prod: access
api_name:
- Access.Attachment.Left
ms.assetid: cee21215-a0b0-9247-976d-9f7899287e54
ms.date: 06/08/2017
---


# Attachment.Left Property (Access)

You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.


## Syntax

 _expression_. **Left**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

In Visual Basic, use a numeric expression to set the value of this property. Values are expressed in twips

For reports, you can set these properties only by using a macro or event procedure in Visual Basic while the report is in Print Preview or being printed.

For reports, the  **Left** property setting is the amount the current section is offset from the left of the page. This property is expressed in twips. You can use this property to specify how far down the page you want a section to print in the section's **Format** event procedure.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

