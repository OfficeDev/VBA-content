---
title: TabFixedHeight, TabFixedWidth Properties
keywords: fm20.chm2002000
f1_keywords:
- fm20.chm2002000
ms.prod: office
ms.assetid: 3b8f8ea4-4916-66ba-1e18-87d56d42efe3
ms.date: 06/08/2017
---


# TabFixedHeight, TabFixedWidth Properties



Sets or returns the fixed height or width of the tabs in [points](vbe-glossary.md).
 **Syntax**
 _object_. **TabFixedHeight** [= _Single_ ]
 _object_. **TabFixedWidth** [= _Single_ ]
The  **TabFixedHeight** and **TabFixedWidth** property syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Single_|Optional. The number of points of the height or width of the tabs on a  **TabStrip** or **MultiPage**.|
 **Settings**
If the value is 0, tab widths are automatically adjusted so that each tab is wide enough to accommodate its contents and each row of tabs spans the width of the control.
If the value is greater than 0, all tabs have an identical width as specified by this property.
 **Remarks**
The minimum size is 4 points.

