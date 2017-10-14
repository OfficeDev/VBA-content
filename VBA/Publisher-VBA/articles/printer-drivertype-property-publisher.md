---
title: Printer.DriverType Property (Publisher)
keywords: vbapb10.chm8978435
f1_keywords:
- vbapb10.chm8978435
ms.prod: publisher
api_name:
- Publisher.Printer.DriverType
ms.assetid: 99c3b4e5-a55a-0f8d-3767-d035d9d6e4df
ms.date: 06/08/2017
---


# Printer.DriverType Property (Publisher)

Specifies the type of driver supported by the printer. Read-only.


## Syntax

 _expression_. **DriverType**

 _expression_A variable that represents a  **Printer** object.


### Return Value

 **PbDriverType**


## Remarks

Possible values for the  **DriverType** property are declared in the **PbDriverType** enumeration and shown in the following table.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **pbDriverTypeNonPostScript**|1|Non PostScript|
| **pbDriverTypePostScript1**|2|PostScript 1|
| **pbDriverTypePostScript2**|3|PostScript 2|
| **pbDriverTypePostScript3**|4|PostScript 3|

