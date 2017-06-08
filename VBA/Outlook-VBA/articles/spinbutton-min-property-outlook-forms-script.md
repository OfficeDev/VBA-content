---
title: SpinButton.Min Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: bc44e375-0eab-bc9d-b8c6-618c62b5fd2f
ms.date: 06/08/2017
---


# SpinButton.Min Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the maximum and minimum acceptable values for the **[Value](spinbutton-value-property-outlook-forms-script.md)** property of a **[SpinButton](spinbutton-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **Min**

 _expression_A variable that represents a  **SpinButton** object.


## Remarks

Clicking a  **SpinButton** changes the **Value** property of the control.

Any integer is an acceptable setting for this property. The recommended range of values is from -32,767 to +32,767. The default value is 1.


