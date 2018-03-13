---
title: Report.MinMaxButtons Property (Access)
keywords: vbaac10.chm13802
f1_keywords:
- vbaac10.chm13802
ms.prod: access
api_name:
- Access.Report.MinMaxButtons
ms.assetid: 8aee0247-804a-e9ee-e11a-11c9c5d37ed6
ms.date: 06/08/2017
---


# Report.MinMaxButtons Property (Access)

You can use the  **MinMaxButtons** property to specify whether the **Maximize** and **Minimize** buttons will be visible on a report. Read/write **Byte**.


## Syntax

 _expression_. **MinMaxButtons**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **MinMaxButtons** property uses the following settings.



| <strong>Setting</strong> | <strong>Visual Basic</strong> | <strong>Description</strong>                                                                     |
|:-------------------------|:------------------------------|:-------------------------------------------------------------------------------------------------|
| None                     | 0                             | The  <strong>Maximize</strong> and <strong>Minimize</strong> buttons aren't visible.             |
| Min Enabled              | 1                             | Only the  <strong>Minimize</strong> button is visible.                                           |
| Max Enabled              | 2                             | Only the  <strong>Maximize</strong> button is visible.                                           |
| Both Enabled             | 3                             | (Default) Both the  <strong>Minimize</strong> and <strong>Maximize</strong> buttons are visible. |

You can set the  **MinMaxButtons** property only in form Design view.


## See also


#### Concepts


[Report Object](report-object-access.md)

