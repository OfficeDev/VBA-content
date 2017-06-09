---
title: Report.OnPage Property (Access)
keywords: vbaac10.chm13768
f1_keywords:
- vbaac10.chm13768
ms.prod: access
api_name:
- Access.Report.OnPage
ms.assetid: d72bab5d-fdb8-99f5-5d27-8227bc0136ec
ms.date: 06/08/2017
---


# Report.OnPage Property (Access)

Sets or returns the value of the  **On Page** box in the **Properties** window of a report. Read/write **String**.


## Syntax

 _expression_. **OnPage**

 _expression_ A variable that represents a **Report** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Page** event occurs after Microsoft Access formats a page of a report for printing, but before the page is printed.

The  **OnPage** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Page** box in the report's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Page** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnPage** property in the Immediate window for the "Purchase Order" report.


```vb
Debug.Print Reports("Purchase Order").OnPage
```


## See also


#### Concepts


[Report Object](report-object-access.md)

