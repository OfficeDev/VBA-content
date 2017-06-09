---
title: CodeMask.Add Method (Project)
keywords: vbapj.chm131648
f1_keywords:
- vbapj.chm131648
ms.prod: project-server
api_name:
- Project.CodeMask.Add
ms.assetid: 78a7afaa-1a19-6d64-1341-63955aaff7e3
ms.date: 06/08/2017
---


# CodeMask.Add Method (Project)

Returns a  **[CodeMaskLevel ](codemasklevel-object-project.md)** object.


## Syntax

 _expression_. **Add**( ** _Sequence_**, ** _Length_**, ** _Separator_** )

 _expression_ A variable that represents a **CodeMask** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sequence_|Optional|**Long**|Specifies the type of sequence in the code mask. Can be one of the  **[PjCustomOutlineCodeSequence](pjcustomoutlinecodesequence-enumeration-project.md)** constants. The default value is **pjCustomOutlineCodeNumbers**.|
| _Length_|Optional|**Variant**|Specifies the length for a given level in the code mask. Can be the string "Any" or an integer value between 1 and 255. |
| _Separator_|Optional|**String**|The character that separates the level of a code mask from the next code mask. Can be one of the following characters: ".", "-", "+", or "/". |

### Return Value

 **CodeMaskLevel**


