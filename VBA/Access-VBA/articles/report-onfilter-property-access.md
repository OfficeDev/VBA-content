---
title: Report.OnFilter Property (Access)
keywords: vbaac10.chm13869
f1_keywords:
- vbaac10.chm13869
ms.prod: access
api_name:
- Access.Report.OnFilter
ms.assetid: 72af402e-8e37-328e-b0f4-89f54f59bce0
ms.date: 06/08/2017
---


# Report.OnFilter Property (Access)

Sets or returns the value of the  **On Filter** box in the **Properties** window of a report. Read/write **String**.


## Syntax

 _expression_. **OnFilter**

 _expression_ A variable that represents a **Report** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Filter** event occurs when a report is opened and its records are displayed.

The  **OnFilter** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Filter** box in the report's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Filter** box is blank, the property value is an empty string.


## Example

The following example associates the  **Filter** property for the "Catalog" report to the event "Report_Filter".


```vb
Reports("Catalog").OnFilter = "[Event Procedure]"
```


## See also


#### Concepts


[Report Object](report-object-access.md)

