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


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **Dictionary** object.|
| _compare_|Optional. If provided,  _compare_ is a value representing the comparison mode used by functions such as **StrComp**.|
 **Settings**
The  _compare_ argument can have the following values:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison using the setting of the  **Option Compare** statement.|
|**vbBinaryCompare**| 0|Performs a binary comparison.|
|**vbTextCompare**| 1|Performs a textual comparison.|
|**vbDatabaseCompare**| 2|Microsoft Access only. Performs a comparison based on information in your database.|
 **Remarks**
An error occurs if you try to change the comparison mode of a  **Dictionary** object that already contains data.
The  **CompareMode** property uses the same values as the _compare_ argument for the **StrComp** function. Values greater than 2 can be used to refer to comparisons using specific Locale IDs (LCID).

