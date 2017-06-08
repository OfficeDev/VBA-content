---
title: Report.GotFocus Event (Access)
keywords: vbaac10.chm13887
f1_keywords:
- vbaac10.chm13887
ms.prod: access
api_name:
- Access.Report.GotFocus
ms.assetid: 667b4798-4407-f60f-af3a-7788a0501761
ms.date: 06/08/2017
---


# Report.GotFocus Event (Access)

The  **GotFocus** event occurs when the report receives the focus.


## Syntax

 _expression_. **GotFocus**

 _expression_ A variable that represents a **Report** object.


## Remarks

To run a macro or event procedure when these events occur, set the  **OnGotFocus** property to the name of the macro or to [Event Procedure].

These events occur when the focus moves in response to a user action, such as pressing the TAB key or clicking the object, or when you use the  **SetFocus** method in Visual Basic or the SelectObject, GoToRecord, GoToControl, or GoToPage action in a macro.


## See also


#### Concepts


[Report Object](report-object-access.md)

