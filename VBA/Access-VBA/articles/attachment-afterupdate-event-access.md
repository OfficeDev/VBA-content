---
title: Attachment.AfterUpdate Event (Access)
keywords: vbaac10.chm14020
f1_keywords:
- vbaac10.chm14020
ms.prod: access
api_name:
- Access.Attachment.AfterUpdate
ms.assetid: 09dfe871-0e56-38fc-46d2-c517ea795907
ms.date: 06/08/2017
---


# Attachment.AfterUpdate Event (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **AfterUpdate** event occurs. Read/write **String**.


## Syntax

 _expression_. **AfterUpdate**

 _expression_ An expression that returns an **Attachment** object.


## Remarks

The  **AfterUpdate** event applies only to controls on a form, not controls on a report.

To run a macro or event procedure when this event occurs, set the  **AfterUpdate** property to the name of the macro or to [Event Procedure].

The  **AfterUpdate** event is triggered when a control or record is updated. Within a record, changed data in each control is updated when the control loses the focus or when the user presses ENTER or TAB.

 **AfterUpdate** macros and event procedures run only if you change the data in a control. This event does not occur when a value changes in a calculated control. **AfterUpdate** macros and event procedures for a form run only if you change the data in one or more controls in the record.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

