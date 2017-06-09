---
title: ScrollBar.SmallChange Property (Outlook Forms Script)
keywords: olfm10.chm2001940
f1_keywords:
- olfm10.chm2001940
ms.prod: outlook
ms.assetid: cd8b6b7f-118a-1cda-00af-11ab74f6617a
ms.date: 06/08/2017
---


# ScrollBar.SmallChange Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the amount of movement that occurs when the user clicks either scroll arrow in a **[ScrollBar](scrollbar-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **SmallChange**

 _expression_A variable that represents a  **ScrollBar** object.


## Remarks

The  **SmallChange** property specifies the amount of change to the **[Value](scrollbar-value-property-outlook-forms-script.md)** property.

The  **SmallChange** property does not have units.

Any integer is an acceptable setting for this property. The recommended range of values is from -32,767 to +32,767. The default value is 1.


