---
title: Report.DrawWidth Property (Access)
keywords: vbaac10.chm13755
f1_keywords:
- vbaac10.chm13755
ms.prod: access
api_name:
- Access.Report.DrawWidth
ms.assetid: 1bda5387-9244-f150-2165-8dba1684ca25
ms.date: 06/08/2017
---


# Report.DrawWidth Property (Access)

You can use the  **DrawWidth** property to specify the line width for the **[Line](report-line-method-access.md)**, **[Circle](report-circle-method-access.md)**, and **[Pset](report-pset-method-access.md)** methods to print lines on reports. Read/write **Integer**.


## Syntax

 _expression_. **DrawWidth**

 _expression_ A variable that represents a **Report** object.


## Remarks

You can set the  **DrawWidth** property to an **Integer** value of 1 through 32,767. This value represents the width of the line in pixels. The default is 1, or 1 pixel wide.

You can set this property in an event procedure specified by a section's **OnPrint** property setting.

Increase the value of this property to increase the width of the line. If the  **DrawWidth** property setting is greater than 3, **[DrawMode](report-drawmode-property-access.md)** property settings 1 through 4 produce a solid line (the **DrawStyle** property setting isn't changed). Setting the **DrawWidth** property to 1 enables the **DrawStyle** property to produce the results shown in the setting table of the **DrawStyle** property.


## See also


#### Concepts


[Report Object](report-object-access.md)

