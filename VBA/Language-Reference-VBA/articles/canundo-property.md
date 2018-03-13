---
title: CanUndo Property
keywords: fm20.chm5225016
f1_keywords:
- fm20.chm5225016
ms.prod: office
api_name:
- Office.CanUndo
ms.assetid: e96f23c1-5a82-0f94-4bef-aaf9767db719
ms.date: 06/08/2017
---


# CanUndo Property



Indicates whether the last user action can be undone.
 **Syntax**
 _object_. **CanUndo**
The  **CanUndo** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong> |
|:----------------------|:-----------------------------|
| <em>object</em>       | Required. A valid object.    |

 **Return Values**
The  **CanUndo** property return values are:


| <strong>Value</strong> | <strong>Description</strong>                  |
|:-----------------------|:----------------------------------------------|
| <strong>True</strong>  | The most recent user action can be undone.    |
| <strong>False</strong> | The most recent user action cannot be undone. |

 **Remarks**
 **CanUndo** is read-only.
Many user actions can be undone with the Undo command. The  **CanUndo** property indicates whether the most recent action can be undone.

