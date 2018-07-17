---
title: Report.Unload Event (Access)
keywords: vbaac10.chm13886
f1_keywords:
- vbaac10.chm13886
ms.prod: access
api_name:
- Access.Report.Unload
ms.assetid: 05f0d51e-8fa0-9547-6b22-e7711754d1a5
ms.date: 06/08/2017
---


# Report.Unload Event (Access)

The  **Unload** event occurs after a report is closed but before it's removed from the screen.


## Syntax

 _expression_. **Unload**( ** _Cancel_**, )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**||

## Remarks

To run a macro or event procedure when these events occur, set the  **OnUnload** property to the name of the macro or to [Event Procedure].

The  **Unload** event is caused by user actions such as:


- Closing the report.
    
- Running the Close action in a macro.
    
- Quitting an application by right-clicking the application's taskbar button and then clicking  **Close**.
    
- Quitting Windows while an application is running.
    
By running a macro or an event procedure when a report's  **Unload** event occurs, you can verify that the report should be unloaded or specify actions that should take place when the report is unloaded. You can also open another report or display a dialog box requesting the user's name to make a log entry indicating who used the report.

When you close a report, the following events occur in this order:

**Unload** → **Deactivate** → **Close**

The  **Unload** event occurs before the **Close** event. The **Unload** event can be canceled, but the **Close** event cannot.


 **Note**  When you create macros or event procedures for events related to the  **Unload** event, such as **Deactivate** and **LostFocus**, be sure that they don't conflict (for example, make sure you don't cause something to happen in one macro or procedure that is canceled in another) and that they don't cause cascading events.


## See also


#### Concepts


[Report Object](report-object-access.md)

