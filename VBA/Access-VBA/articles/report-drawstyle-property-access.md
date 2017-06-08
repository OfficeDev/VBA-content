---
title: Report.DrawStyle Property (Access)
keywords: vbaac10.chm13754
f1_keywords:
- vbaac10.chm13754
ms.prod: access
api_name:
- Access.Report.DrawStyle
ms.assetid: 0dd2afb9-d310-3637-6ed7-e66c9ad3460d
ms.date: 06/08/2017
---


# Report.DrawStyle Property (Access)

You can use the  **DrawStyle** property to specify the line style when using the **[Line](report-line-method-access.md)** and **[Circle](report-circle-method-access.md)** methods to print lines on reports. Read/write **Integer**.


## Syntax

 _expression_. **DrawStyle**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **DrawStyle** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|(Default) Solid line, transparent interior|
|1|Dash, transparent interior|
|2|Dot, transparent interior|
|3|Dash-dot, transparent interior|
|4|Dash-dot-dot, transparent interior|
|5|Invisible line, transparent interior|
|6|Invisible line, solid interior|

 **Note**  You can set this property in an event procedure specified by a section's **OnPrint** property setting.

The  **DrawStyle** property produces the results described in the preceding table if the **[DrawWidth](report-drawwidth-property-access.md)** property is set to 1. If the **DrawWidth** property setting is greater than 3, the **DrawStyle** property settings 1 through 4 produce a solid line (the **DrawStyle** property value isn't changed).


## See also


#### Concepts


[Report Object](report-object-access.md)

