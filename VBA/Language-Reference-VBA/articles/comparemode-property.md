---
title: CompareMode Property
keywords: vblr6.chm2181931
f1_keywords:
- vblr6.chm2181931
ms.prod: office
api_name:
- Office.CompareMode
ms.assetid: 75893886-8bed-4685-b483-18b3d39569da
ms.date: 06/08/2017
---


# CompareMode Property



 **Description**
Sets and returns the comparison mode for comparing string keys in a  **Dictionary** object.
 **Syntax**
 _object_. **CompareMode** [ = _compare_ ]
The  **CompareMode** property has the following parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                             |
|:----------------------|:-----------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>Dictionary</strong> object.                                                                      |
| <em>compare</em>      | Optional. If provided,  <em>compare</em> is a value representing the comparison mode used by functions such as <strong>StrComp</strong>. |

 **Settings**
The  _compare_ argument can have the following values:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>                                                               |
|:------------------------------------|:-----------------------|:-------------------------------------------------------------------------------------------|
| <strong>vbUseCompareOption</strong> | -1                     | Performs a comparison using the setting of the  <strong>Option Compare</strong> statement. |
| <strong>vbBinaryCompare</strong>    | 0                      | Performs a binary comparison.                                                              |
| <strong>vbTextCompare</strong>      | 1                      | Performs a textual comparison.                                                             |
| <strong>vbDatabaseCompare</strong>  | 2                      | Microsoft Access only. Performs a comparison based on information in your database.        |

 **Remarks**
An error occurs if you try to change the comparison mode of a  **Dictionary** object that already contains data.
The  **CompareMode** property uses the same values as the _compare_ argument for the **StrComp** function. Values greater than 2 can be used to refer to comparisons using specific Locale IDs (LCID).

