---
title: PjEnableCancelKey Enumeration (Project)
ms.prod: project-server
api_name:
- Project.PjEnableCancelKey
ms.assetid: a50ff9ef-7462-a414-8680-a127b1bdc9a3
ms.date: 06/08/2017
---


# PjEnableCancelKey Enumeration (Project)

Contains constants that specify how to handle the ** CTRL+BREAK** key combination.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**pjDisabled**|0|**CTRL+BREAK** is ignored.|
|**pjErrorHandler**|2|Sends the interrupt to the macro as a trappable error. The error code is 18.|
|**pjInterrupt**|1|**CTRL+BREAK** interrupts the macro.|

