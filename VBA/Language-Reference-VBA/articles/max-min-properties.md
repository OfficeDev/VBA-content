---
title: Max, Min Properties
keywords: fm20.chm5225063
f1_keywords:
- fm20.chm5225063
ms.prod: office
ms.assetid: 499fb814-b84c-d7a6-1467-9d13afae97e9
ms.date: 06/08/2017
---


# Max, Min Properties



Specify the maximum and minimum acceptable values for the  **Value** property of a **ScrollBar** or **SpinButton**.
 **Syntax**
 _object_. **Max** [= _Long_ ]
 _object_. **Min** [= _Long_ ]
The  **Max** and **Min** property syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. A numeric expression specifying the maximum or minimum  **Value** property setting.|
 **Remarks**
Clicking a  **SpinButton** or moving the scroll box in a **ScrollBar** changes the **Value** property of the control.
The value for the  **Max** property corresponds to the lowest position of a vertical **ScrollBar** or the rightmost position of a horizontal **ScrollBar**. The value for the **Min** property corresponds to the highest position of a vertical **ScrollBar** or the leftmost position of a horizontal **ScrollBar**.
Any integer is an acceptable setting for this property. The recommended range of values is from -32,767 to +32,767. The default value is 1.

 **Note**   **Min** and **Max** refer to locations, not to relative values, on the **ScrollBar**. That is, the value of **Max** could be less than the value of **Min**. If this is the case, moving toward the **Max** (bottom) position means decreasing **Value**; moving toward the **Min** (top) position means increasing **Value**.


