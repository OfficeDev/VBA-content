---
title: SpinButton.SmallChange Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 77e920e5-87ad-cad0-0c14-ac63cf5aa118
ms.date: 06/08/2017
---


# SpinButton.SmallChange Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the amount of movement that occurs when the user clicks either scroll arrow in a **[SpinButton](spinbutton-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **SmallChange**

 _expression_A variable that represents a  **SpinButton** object.


## Remarks

The  **SmallChange** property specifies the amount of change to the **[Value](spinbutton-value-property-outlook-forms-script.md)** property.

The  **SmallChange** property does not have units.

Any integer is an acceptable setting for this property. The recommended range of values is from -32,767 to +32,767. The default value is 1.


