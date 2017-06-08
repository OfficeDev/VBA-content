---
title: AfterUpdate Event
keywords: fm20.chm5224934
f1_keywords:
- fm20.chm5224934
ms.prod: office
api_name:
- Office.AfterUpdate
ms.assetid: 3d15efd4-06c8-136f-c315-7efc44db35b1
ms.date: 06/08/2017
---


# AfterUpdate Event



Occurs after data in a control is changed through the user interface.
 **Syntax**
 **Private Sub**_object_ _**AfterUpdate( )**
The  **AfterUpdate** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
The AfterUpdate event occurs regardless of whether the control is [bound](glossary-vba.md) (that is, when the **RowSource** property specifies a[data source](glossary-vba.md) for the control). This event cannot be canceled. If you want to cancel the update (to restore the previous value of the control), use the BeforeUpdate event and set the _Cancel_ argument to **True**.
The AfterUpdate event occurs after the BeforeUpdate event and before the Exit event for the current control and before the Enter event for the next control in the [tab order](vbe-glossary.md).

