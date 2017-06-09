---
title: TextBox.Value Property (Access)
keywords: vbaac10.chm11039
f1_keywords:
- vbaac10.chm11039
ms.prod: access
api_name:
- Access.TextBox.Value
ms.assetid: 4cb4c33f-dd96-0309-f30b-8e445d123756
ms.date: 06/08/2017
---


# TextBox.Value Property (Access)

Determines or specifies the text in the text box. Read/write  **Variant**.


## Syntax

 _expression_. **Value**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

The  **Text** property returns the formatted string. The **Text** property may be different than the **Value** property for a text box control. The **Text** property is the current contents of the control. The **Value** property is the saved value of the text box control. The **Text** property is always current while the control has the focus.

The  **Value** property returns or sets a control's default property, which is the property that is assumed when you don't explicitly specify a property name.


 **Note**   The **Value** property is not the same as the **DefaultValue** property, which specifies the value that a property is assigned when a new record is created.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

