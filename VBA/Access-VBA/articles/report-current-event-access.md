---
title: Report.Current Event (Access)
keywords: vbaac10.chm13883
f1_keywords:
- vbaac10.chm13883
ms.prod: access
api_name:
- Access.Report.Current
ms.assetid: adfdbda0-c3e9-c3c6-8768-415b4bd270d5
ms.date: 06/08/2017
---


# Report.Current Event (Access)

Occurs when the focus moves to a record, making it the current record, or when the report is refreshed or requeried.


## Syntax

 _expression_. **Current**

 _expression_ A variable that represents a **Report** object.


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnCurrent** property to the name of the macro or to [Event Procedure].

This event occurs both when a report is opened and whenever the focus leaves one record and moves to another. Microsoft Access runs the  **Current** macro or event procedure before the first or next record is displayed.

By running a macro or event procedure when a form's  **Current** event occurs, you can display a message or perform an operation related to the current record.

The  **Current** event also occurs when you refresh a report or requery the report's underlying table or query— for example, when you use the Requery action in a macro or the **Requery** method in Visual Basic code.

When you first open a report, the following events occur in this order:

Open → Load → Resize → Activate → Current


## See also


#### Concepts


[Report Object](report-object-access.md)

