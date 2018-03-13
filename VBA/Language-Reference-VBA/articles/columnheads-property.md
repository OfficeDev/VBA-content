---
title: ColumnHeads Property
keywords: fm20.chm5225021
f1_keywords:
- fm20.chm5225021
ms.prod: office
api_name:
- Office.ColumnHeads
ms.assetid: 55cd26ad-8ef3-8e65-f655-315af620658d
ms.date: 06/08/2017
---


# ColumnHeads Property



Displays a single row of column headings for list boxes, combo boxes, and objects that accept column headings.
 **Syntax**
 _object_. **ColumnHeads** [= _Boolean_ ]
The  **ColumnHeads** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                   |
|:----------------------|:---------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                      |
| <em>Boolean</em>      | Optional. Specifies whether the column headings are displayed. |

 **Settings**
The settings for  _Boolean_ are:


| <strong>Value</strong> | <strong>Description</strong>              |
|:-----------------------|:------------------------------------------|
| <strong>True</strong>  | Display column headings.                  |
| <strong>False</strong> | Do not display column headings (default). |

Headings in combo boxes appear only when the list drops down.
 **Remarks**
When the system uses the first row of data items as column headings, they can't be selected.

