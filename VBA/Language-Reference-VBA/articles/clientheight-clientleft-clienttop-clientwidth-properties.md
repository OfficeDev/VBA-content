---
title: ClientHeight, ClientLeft, ClientTop, ClientWidth Properties
keywords: fm20.chm2000910
f1_keywords:
- fm20.chm2000910
ms.prod: office
ms.assetid: d0754b52-156b-f8a4-3b28-9ce3020bc5f7
ms.date: 06/08/2017
---


# ClientHeight, ClientLeft, ClientTop, ClientWidth Properties



Define the dimensions and location of the display area of a  **TabStrip**.
 **Syntax**
 _object_. **ClientHeight** [ = _Single_ ]
 _object_. **ClientLeft** [ = _Single_ ]
 _object_. **ClientTop** [ = _Single_ ]
 _object_. **ClientWidth** [ = _Single_ ]
The  **ClientHeight**, **ClientLeft**, **ClientTop**, and **ClientWidth** property syntaxes have these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                             |
|:----------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                                                                                                                                                                                                                                                                |
| <em>Single</em>       | Optional. For  <strong>ClientHeight</strong> and <strong>ClientWidth</strong>, specifies the height or width, in points, of the display area. For <strong>ClientLeft</strong> and <strong>ClientTop</strong>, specifies the distance, in points, from the top or left edge of the <strong>TabStrip's</strong> container. |

 **Remarks**
At [run time](vbe-glossary.md),  **ClientLeft**, **ClientTop**, **ClientHeight**, and **ClientWidth** automatically store the coordinates and dimensions of the **TabStrip's** internal area, which is shared by objects in the **TabStrip**.

