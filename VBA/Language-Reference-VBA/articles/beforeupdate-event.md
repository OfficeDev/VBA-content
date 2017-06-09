---
title: BeforeUpdate Event
keywords: fm20.chm2000050
f1_keywords:
- fm20.chm2000050
ms.prod: office
api_name:
- Office.BeforeUpdate
ms.assetid: ccf0fa5d-a069-cba6-5725-072b141fa80b
ms.date: 06/08/2017
---


# BeforeUpdate Event



Occurs before data in a control is changed.
 **Syntax**
 **Private Sub**_object_ _**BeforeUpdate( ByVal**_Cancel_**As MSForms.ReturnBoolean)**
The  **BeforeUpdate** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Cancel_|Required. Event status.  **False** indicates that the control should handle the event (default). **True** cancels the update and indicates the application should handle the event.|
 **Remarks**
The BeforeUpdate event occurs regardless of whether the control is [bound](glossary-vba.md) (that is, when the **RowSource** property specifies a[data source](glossary-vba.md) for the control). This event occurs before the AfterUpdate and Exit events for the control (and before the Enter event for the next control that receives[focus](vbe-glossary.md)).
If you set the  _Cancel_ argument to **True**, the focus remains on the control and neither the AfterUpdate event nor the Exit event occurs.

