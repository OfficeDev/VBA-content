---
title: Report.Resize Event (Access)
keywords: vbaac10.chm13885
f1_keywords:
- vbaac10.chm13885
ms.prod: access
api_name:
- Access.Report.Resize
ms.assetid: cd2c1c2a-959b-a5d0-9f75-a7443a9a57f1
ms.date: 06/08/2017
---


# Report.Resize Event (Access)

The  **Resize** event occurs when a report is opened and whenever the size of a report changes.


## Syntax

 _expression_. **Resize**

 _expression_ A variable that represents a **Report** object.


## Remarks

To run a macro or event procedure when this event occurs, set the  **[OnResize](form-onresize-property-access.md)** property to the name of the macro or to [Event Procedure].

This event occurs if you change the size of the report in a macro or event procedure—for example, if you use the MoveSize action in a macro to resize the report.

By running a macro or an event procedure when a  **Resize** event occurs, you can move or resize a control when the report it's on is resized. You can also use a **Resize** event to recalculate variables or reset properties that may depend on the size of the report.

When you first open a report, the following events occur in this order:

Open → Load → Resize → Activate → Current


 **Note**  You need to be careful if you use a  **MoveSize**, **Maximize**, **Minimize**, or **Restore** action (or the corresponding methods of the **DoCmd** object) in a **Resize** macro or event procedure. These actions can trigger a **Resize** event for the report, and thus cause a cascading event.


## See also


#### Concepts


[Report Object](report-object-access.md)

