---
title: SelectionMargin Property
keywords: fm20.chm2001860
f1_keywords:
- fm20.chm2001860
ms.prod: office
api_name:
- Office.SelectionMargin
ms.assetid: 1e86e761-7427-e6a2-0b66-887bf89f9fa7
ms.date: 06/08/2017
---


# SelectionMargin Property



Specifies whether the user can select a line of text by clicking in the region to the left of the text.
 **Syntax**
 _object_. **SelectionMargin** [= _Boolean_ ]
The  **SelectionMargin** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                     |
|:----------------------|:-----------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                        |
| <em>Boolean</em>      | Optional. Whether clicking in the margin selects a line of text. |

 **Settings**
The settings for  _Boolean_ are:


| <strong>Value</strong> | <strong>Description</strong>                           |
|:-----------------------|:-------------------------------------------------------|
| <strong>True</strong>  | Clicking in margin causes selection of text (default). |
| <strong>False</strong> | Clicking in margin does not cause selection of text.   |

 **Remarks**
When the  **SelectionMargin** property is **True**, the selection margin occupies a thin strip along the left edge of a control's edit region. When set to **False**, the entire edit region can store text.
If the  **SelectionMargin** property is set to **True** when a control is printed, the selection margin also prints.

