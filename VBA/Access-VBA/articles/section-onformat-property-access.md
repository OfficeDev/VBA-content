---
title: Section.OnFormat Property (Access)
keywords: vbaac10.chm12204,vbaac10.chm4089
f1_keywords:
- vbaac10.chm12204,vbaac10.chm4089
ms.prod: access
api_name:
- Access.Section.OnFormat
ms.assetid: 061652a9-0253-8dc2-a8c0-02daa40d132d
ms.date: 06/08/2017
---


# Section.OnFormat Property (Access)

Sets or returns the value of the  **On Format** box in the **Properties** window of a report section. Read/write **String**.


## Syntax

 _expression_. **OnFormat**

 _expression_ A variable that represents a **Section** object.


## Remarks

This property is helpful for programmatically changing the action Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The  **Format** event occurs when Microsoft Access determines which data belongs in a report section, but before Access formats the section for previewing or printing.

The  **OnFormat** value will be one of the following, depending on the selection chosen in the **Choose Builder** window (accessed by clicking the **Build** button next to the **On Format** box in the report section's **Properties** window):


- If Expression Builder is chosen, the value will be "= _expression_ ", where _expression_ is the expression from the Expression Builder window.
    
- If Macro Builder is chosen, the value is the name of the macro. 
    
- If Code Builder is chosen, the value will be "[Event Procedure]". 
    
If the  **On Format** box is blank, the property value is an empty string.


## Example

The following example prints the value of the  **OnFormat** property in the Immediate window for the "GroupHeader0" section in the "Purchase Order" report.


```vb
Debug.Print Reports("Purchase Order").Section("GroupHeader0").OnFormat
```


## See also


#### Concepts


[Section Object](section-object-access.md)

