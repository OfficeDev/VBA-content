---
title: MatchFound Property
keywords: fm20.chm5225061
f1_keywords:
- fm20.chm5225061
ms.prod: office
api_name:
- Office.MatchFound
ms.assetid: db350684-1758-a849-c9e1-34714a00f1c3
ms.date: 06/08/2017
---


# MatchFound Property



Indicates whether the text that a user has typed into a combo box matches any of the entries in the list.
 **Syntax**
 _object_. **MatchFound**
The  **MatchFound** property syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong> |
|:----------------------|:-----------------------------|
| <em>object</em>       | Required. A valid object.    |

 **Return Values**
The  **MatchFound** property return values are:


| <strong>Value</strong> | <strong>Description</strong>                                                                     |
|:-----------------------|:-------------------------------------------------------------------------------------------------|
| <strong>True</strong>  | The contents of the  <strong>Value</strong> property matches one of the records in the list.     |
| <strong>False</strong> | The contents of  <strong>Value</strong> does not match any of the records in the list (default). |

 **Remarks**
The  **MatchFound** property is read-only. It is not applicable when the **MatchEntry** property is set to **fmMatchEntryNone**.

