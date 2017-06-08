---
title: Name Property (Outlook Controls)
ms.prod: outlook
ms.assetid: 5abf1af8-4914-6b76-99e6-9f78b46bae73
ms.date: 06/08/2017
---


# Name Property (Outlook Controls)

Returns or sets a  **String** that identifies the control. Read/write.


## Syntax

 _expression_. **Name**

 _expression_A variable that represents an Outlook control object.


## Remarks

Guidelines for assigning a string to  **Name**, such as the maximum length of the name, vary from one application to another.

For objects, the default value of  **Name** consists of the object's class name followed by an integer. For example, the default name for the first **[OlkTextBox](olktextbox-object-outlook.md)** you place on a form is OlkTextBox1. The default name for the second **OlkTextBox** is OlkTextBox2.

You can set the  **Name** property for a control from the control's property sheet or, for controls added at run time, by using program statements. If you add a control at design time, you cannot modify its **Name** property at run time.

Each control added to a form at design time must have a unique name.


