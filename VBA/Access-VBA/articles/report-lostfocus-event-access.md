---
title: Report.LostFocus Event (Access)
keywords: vbaac10.chm13888
f1_keywords:
- vbaac10.chm13888
ms.prod: access
api_name:
- Access.Report.LostFocus
ms.assetid: 8b80c2bc-8be4-1842-4011-0e6475b3a865
ms.date: 06/08/2017
---


# Report.LostFocus Event (Access)

The  **LostFocus** event occurs when the specified object loses the focus.


## Syntax

 _expression_. **LostFocus**

 _expression_ A variable that represents a **Report** object.


## Remarks

To run a macro or event procedure when these events occur, set the  **OnLostFocus** property to the name of the macro or to [Event Procedure].

This event occurs when the focus moves in response to a user action, such as pressing the TAB key or clicking the object, or when you use the  **SetFocus** method in Visual Basic or the SelectObject, GoToRecord, GoToControl, or GoToPage action in a macro.


## See also


#### Concepts


[Report Object](report-object-access.md)

