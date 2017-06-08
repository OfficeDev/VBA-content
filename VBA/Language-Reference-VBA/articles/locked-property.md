---
title: Locked Property
keywords: fm20.chm5225059
f1_keywords:
- fm20.chm5225059
ms.prod: office
api_name:
- Office.Locked
ms.assetid: 08bf09c4-0445-0749-daf2-a0fab8787ea8
ms.date: 06/08/2017
---


# Locked Property



Specifies whether a control can be edited.
 **Syntax**
 _object_. **Locked** [= _Boolean_ ]
The  **Locked** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the control can be edited.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|You can't edit the value.|
|**False**|You can edit the value (default).|
 **Remarks**
When a control is locked and enabled, it can still initiate events and can still receive the [focus](vbe-glossary.md).

