---
title: Report.Timer Event (Access)
keywords: vbaac10.chm13897
f1_keywords:
- vbaac10.chm13897
ms.prod: access
api_name:
- Access.Report.Timer
ms.assetid: 52e3db7f-a61c-8144-e39b-0f9daf61bd98
ms.date: 06/08/2017
---


# Report.Timer Event (Access)

The  **Timer** event occurs for a report at regular intervals as specified by the report's **[TimerInterval](report-timerinterval-property-access.md)** property.


## Syntax

 _expression_. **Timer**

 _expression_ A variable that represents a **Report** object.


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnTimer** property to the name of the macro or to [Event Procedure].

By running a macro or event procedure when a  **Timer** event occurs, you can control what Microsoft Access does at every timer interval. For example, you might want to requery underlying records or repaint the screen at specified intervals.

The  **TimerInterval** property setting of the report specifies the interval, in milliseconds, between **Timer** events. The interval can be between 0 and 2,147,483,647 milliseconds. Setting the **TimerInterval** property to 0 prevents the **Timer** event from occurring.


## See also


#### Concepts


[Report Object](report-object-access.md)

