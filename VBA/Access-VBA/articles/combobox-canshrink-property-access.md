---
title: ComboBox.CanShrink Property (Access)
keywords: vbaac10.chm11497
f1_keywords:
- vbaac10.chm11497
ms.prod: access
api_name:
- Access.ComboBox.CanShrink
ms.assetid: 6f74e442-0b65-1d15-b247-6e12b9a08f1e
ms.date: 06/08/2017
---


# ComboBox.CanShrink Property (Access)

Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data the control contains. Read/write  **Boolean**.


## Syntax

 _expression_. **CanShrink**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **CanShrink** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The control shrinks vertically so that the data it contains can be printed or previewed without leaving blank lines.|
|No|**False**|(Default) The control doesn't shrink.|
This property setting is read-only in a macro or Visual Basic in any view but Design view.

You can use this property to control the appearance of printed forms and reports. When you set the property to Yes, the object automatically adjusts so any amount of data can be printed. When a control shrinks, the controls below it move up the page.

If you set a control's  **CanShrink** property to Yes, Microsoft Access does not set the section's **CanShrink** property to Yes.

Sections shrink vertically across their entire width. For example, suppose a form has two text boxes side by side in a section, and each control has its  **CanShrink** property set to Yes. If one text box contains one line of data and the other text box contains two lines of data, both text boxes will be two lines long because the section is sized across its entire width. To shrink the data independently, you can place two subform or subreport controls side by side, and set their **CanShrink** property to Yes.

When you use the  **CanShrink** property, remember that:


- The property settings don't affect the horizontal spacing between controls; they affect only the vertical space the controls occupy.
    
- Overlapping controls can't shrink.
    
- The height of a large control can prevent controls beside it from shrinking. For example, if several short controls are on the left side of a report's detail section and one tall control, such as an unbound object frame, is on the right side, the controls on the left won't shrink, even if they contain no data.
    

 **Note**  


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

