---
title: CommandButton.Width Property (Access)
keywords: vbaac10.chm10469
f1_keywords:
- vbaac10.chm10469
ms.prod: access
api_name:
- Access.CommandButton.Width
ms.assetid: 03729218-4c70-8312-ab61-be3cf4b7a029
ms.date: 06/08/2017
---


# CommandButton.Width Property (Access)

Gets or sets the width of the specified object in twips. Read/write  **Integer**.


## Syntax

 _expression_. **Width**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

For report controls, you can set the  **Width** property when you print or preview a report only by using a macro or an event procedure specified in a section's **OnFormat** event property setting.

You can't set this property for an object once the print process has started.

Microsoft Access automatically sets the  **Width** property when you create or size a control or when you size a window in form Design View or report Design view.

The width of forms and reports is measured from the inside of their borders. The width of controls is measured from the center of their borders so controls with different border widths align correctly. The margins for forms and reports are set in the  **Page Setup** dialog box, available by clicking **Page Setup** on the **File** menu.


## Example

The following code resizes a command button to a 1-inch by 1-inch square button (the default unit of measurement in Visual Basic is twips; 1440 twips equals one inch):


```vb
Me!cmdSizeButton.Height = 1440 ' 1440 twips = 1 inch. 
Me!cmdSizeButton.Width = 1440
```


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

