---
title: Report.OnMouseMove Property (Access)
keywords: vbaac10.chm13864
f1_keywords:
- vbaac10.chm13864
ms.prod: access
api_name:
- Access.Report.OnMouseMove
ms.assetid: 260c8b56-2985-1da4-7c3f-1398b54666b3
ms.date: 06/08/2017
---


# Report.OnMouseMove Property (Access)

Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.


## Syntax

 _expression_. **OnMouseMove**

 _expression_ A variable that represents a **Report** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **MouseMove** event occurs when the user moves the mouse over the report.

The  **OnMouseMove** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Mouse Move** box in the report's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Mouse Move** box is blank, the property value is an empty string.


## See also


#### Concepts


[Report Object](report-object-access.md)

