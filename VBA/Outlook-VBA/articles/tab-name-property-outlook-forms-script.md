---
title: Tab.Name Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 270b9d06-fdba-44a4-ba4c-b6b1a57a80d1
ms.date: 06/08/2017
---


# Tab.Name Property (Outlook Forms Script)

Returns or sets a  **String** that specifies the name of a control. Read/write.


## Syntax

 _expression_. **Name**

 _expression_A variable that represents a  **Tab** object.


## Remarks

Guidelines for assigning a string to  **Name**, such as the maximum length of the name, vary from one application to another.

For objects, the default value of  **Name** consists of the object's class name followed by an integer. For example, the default name for the first **[TextBox](textbox-object-outlook-forms-script.md)** you place on a form is TextBox1. The default name for the second **TextBox** is TextBox2.

You can set the  **Name** property for a control from the control's property sheet or, for controls added at run time, by using program statements. If you add a control at design time, you cannot modify its **Name** property at run time.

Each control added to a form at design time must have a unique name.


