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


| <strong>Part</strong> | <strong>Description</strong>                 |
|:----------------------|:---------------------------------------------|
| <em>object</em>       | Required. A valid object.                    |
| <em>Boolean</em>      | Optional. Whether the control can be edited. |

 **Settings**
The settings for  _Boolean_ are:


| <strong>Value</strong> | <strong>Description</strong>      |
|:-----------------------|:----------------------------------|
| <strong>True</strong>  | You can't edit the value.         |
| <strong>False</strong> | You can edit the value (default). |

 **Remarks**
When a control is locked and enabled, it can still initiate events and can still receive the [focus](vbe-glossary.md).

