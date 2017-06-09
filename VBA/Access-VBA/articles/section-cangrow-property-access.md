---
title: Section.CanGrow Property (Access)
keywords: vbaac10.chm12197
f1_keywords:
- vbaac10.chm12197
ms.prod: access
api_name:
- Access.Section.CanGrow
ms.assetid: ea3feebb-8193-2fc3-758b-533b65d0e5b0
ms.date: 06/08/2017
---


# Section.CanGrow Property (Access)

Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data the control contains. Read/write  **Boolean**.


## Syntax

 _expression_. **CanGrow**

 _expression_ A variable that represents a **Section** object.


## Remarks

The  **CanGrow** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True** (?1)|The section or control grows vertically so that all data it contains can be printed or previewed.|
|No|**False** (0)|(Default) The section or control doesn't grow. Data that doesn't fit within the fixed size of the section or control won't be printed or previewed.|
This property setting is read-only in a macro or Visual Basic in any view but Design view.

You can use this property to control the appearance of printed forms and reports. When you set the property to Yes, the object automatically adjusts so any amount of data can be printed. When a control grows, the controls below it move down the page.

If you set a control's  **CanGrow** property to Yes, Microsoft Access automatically sets the **CanGrow** property of the section containing the control to Yes.

Sections grow vertically across their entire width. To grow the data independently, you can place two subform or subreport controls side by side, and set their  **CanGrow** property to Yes.

When you use the  **CanGrow** property, remember that:


- The property settings don't affect the horizontal spacing between controls; they affect only the vertical space the controls occupy.
    
- Overlapping controls can't grow.
    

 **Note**  


## See also


#### Concepts


[Section Object](section-object-access.md)

