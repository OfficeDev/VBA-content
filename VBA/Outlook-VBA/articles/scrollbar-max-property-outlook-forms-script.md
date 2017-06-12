---
title: ScrollBar.Max Property (Outlook Forms Script)
keywords: olfm10.chm2001510
f1_keywords:
- olfm10.chm2001510
ms.prod: outlook
ms.assetid: e87f11c7-2e94-2b29-39ba-b2f2121efed6
ms.date: 06/08/2017
---


# ScrollBar.Max Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the maximum and minimum acceptable values for the **[Value](scrollbar-value-property-outlook-forms-script.md)** property of a **[ScrollBar](scrollbar-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **Max**

 _expression_A variable that represents a  **ScrollBar** object.


## Remarks

Moving the scroll box in a  **ScrollBar** changes the **Value** property of the control.

The value for the  **Max** property corresponds to the lowest position of a vertical **ScrollBar** or the rightmost position of a horizontal **ScrollBar**.

Any integer is an acceptable setting for this property. The recommended range of values is from -32,767 to +32,767. The default value is 1.

 **Max** refer to locations, not to relative values, on the **ScrollBar**. That is, the value of  **Max** could be less than the value of **[Min](scrollbar-min-property-outlook-forms-script.md)**. If this is the case, moving toward the  **Max** (bottom) position means decreasing **Value**; moving toward the  **Min** (top) position means increasing **Value**.


