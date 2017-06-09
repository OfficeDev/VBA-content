---
title: Report.Load Event (Access)
keywords: vbaac10.chm13884
f1_keywords:
- vbaac10.chm13884
ms.prod: access
api_name:
- Access.Report.Load
ms.assetid: 966527a0-4c61-9f5e-50ca-791d39bd24ac
ms.date: 06/08/2017
---


# Report.Load Event (Access)

Occurs when a report is opened and its records are displayed.


## Syntax

 _expression_. **Load**

 _expression_ A variable that represents a **Report** object.


## Remarks

To run a macro or event procedure when these events occur, set the  **OnLoad** property to the name of the macro or to [Event Procedure].

The  **Load** event is caused by user actions such as:


- Starting an application. 
    
- Opening a report by clicking  **Open** in the Database window.
    
- Running the OpenReport action in a macro.
    
By running a macro or an event procedure when a report's  **Load** event occurs, you can specify default settings for controls, or display calculated data that depends on the data in the report's records.

By running a macro or an event procedure when a report's  **Unload** event occurs, you can verify that the report should be unloaded or specify actions that should take place when the report is unloaded.

When you first open a report, the following events occur in this order:

Open → Load → Resize → Activate → Current

If you're trying to decide whether to use the  **Open** or **Load** event for your macro or event procedure, one significant difference is that the **Open** event can be canceled, but the **Load** event cannot. For example, if you're dynamically building a record source for a report in an event procedure for the report's **Open** event, you can cancel opening the report if there are no records to display.

When you close a report, the following events occur in this order:

Unload → Deactivate → Close

The  **Unload** event occurs before the **Close** event. The **Unload** event can be canceled, but the **Close** event cannot.


 **Note**  When you create macros or event procedures for events related to the  **Load** event, such as **Activate** and **GotFocus**, be sure that they don't conflict (for example, make sure you don't cause something to happen in one macro or procedure that is canceled in another) and that they don't cause cascading events.


## See also


#### Concepts


[Report Object](report-object-access.md)

