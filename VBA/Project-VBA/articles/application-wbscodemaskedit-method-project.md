---
title: Application.WBSCodeMaskEdit Method (Project)
keywords: vbapj.chm630
f1_keywords:
- vbapj.chm630
ms.prod: project-server
api_name:
- Project.Application.WBSCodeMaskEdit
ms.assetid: 37ade035-5235-54ab-92fa-962c4172dcdc
ms.date: 06/08/2017
---


# Application.WBSCodeMaskEdit Method (Project)

Edits the work breakdown structure (WBS) code mask.


## Syntax

 _expression_. **WBSCodeMaskEdit**( ** _CodePrefix_**, ** _Level_**, ** _Sequence_**, ** _Length_**, ** _Separator_**, ** _CodeGenerate_**, ** _VerifyUniqueness_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CodePrefix_|Optional|**String**|The WBS code prefix for the project.|
| _Level_|Optional|**Long**|The level of code mask to edit. The default value is one greater than the highest level currently specified for the field.|
| _Sequence_|Optional|**Long**|The sequence for the level specified in Level. Can be one of the following  **[PjWBSSequence](pjwbssequence-enumeration-project.md)** constants: **pjWBSOrderedNumbers**, **pjWBSOrderedLowercaseLetters**, **pjWBSOrderedUppercaseLetters**, or **pjWBSUnorderedCharacters**. The default value is **pjWBSOrderedNumbers**.|
| _Length_|Optional|**Variant**|Specifies the length for a given level. Can be the string "Any" or an integer value 1-255. The default value is "Any".|
| _Separator_|Optional|**String**|The character that separates the level specified in Level from the next. Can be one of ".", "-", "+", or "/". The default value is ".".|
| _CodeGenerate_|Optional|**Boolean**|**True** if a new WBS code is generated whenever a new task is created.|
| _VerifyUniqueness_|Optional|**Boolean**|**True** if new WBS codes are verified to be unique.|

### Return Value

 **Boolean**


## Remarks

Using the  **WBSCodeMaskEdit** method without specifying any arguments brings up the **WBS Code Definition** dialog box.


## Example

The following example creates a two-level mask for WBS codes. Using the default values for the method, the first line specifies that the first level uses two-digit codes, sorted by number, and is separated from the next level by the "-" character. The second line specifies that uppercase letters, sorted alphabetically, are used for the second level and are separated from the next level by the default "." character. By default, new codes using the mask are generated for each new task and are verified for uniqueness within the project.

Possible results would be in the order 01-A.1, 01-A.2, 01-B.1, 01-B.2, 02-A.1, 02-A.2, 02-B.1, 02-B.2, and so on.




```vb
Sub SetNewWBSCode() 
 Application.WBSCodeMaskEdit Length:=2, Separator:="-" 
 Application.WBSCodeMaskEdit Length:=1, Sequence:=pjWBSOrderedUppercaseLetters 
End Sub
```


