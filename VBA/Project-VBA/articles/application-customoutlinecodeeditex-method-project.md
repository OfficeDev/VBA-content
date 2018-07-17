---
title: Application.CustomOutlineCodeEditEx Method (Project)
keywords: vbapj.chm631
f1_keywords:
- vbapj.chm631
ms.prod: project-server
api_name:
- Project.Application.CustomOutlineCodeEditEx
ms.assetid: fc0f60a6-18bf-a8e6-9376-1222a126a64a
ms.date: 06/08/2017
---


# Application.CustomOutlineCodeEditEx Method (Project)

Edits a local outline code custom field definition.


## Syntax

 _expression_. **CustomOutlineCodeEditEx**( ** _FieldID_**, ** _Level_**, ** _Sequence_**, ** _Length_**, ** _Separator_**, ** _OnlyLookUpTableCodes_**, ** _OnlyCompleteCodes_**, ** _OnlyLeaves_**, ** _MatchGeneric_**, ** _RequiredCode_**, ** _LookupDefault_**, ** _DefaultValue_**, ** _SortOrder_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**PjCustomField**|Specifies the custom outline code to edit. Can be one of the non-enterprise  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _Level_|Optional|**Long**|The level of code mask to edit. The default value is one greater than the highest level currently specified for the field.|
| _Sequence_|Optional|**PjCustomOutlineCodeSequence**|The sequence for the level specified in Level. Can be one of the  **[PjCustomOutlineCodeSequence](pjcustomoutlinecodesequence-enumeration-project.md)** constants. The default value is **pjCustomOutlineCodeNumbers**.|
| _Length_|Optional|**Variant**|Specifies the length for a given level. Can be the string "Any" or an integer value from 1 through 255. The default value is "Any".|
| _Separator_|Optional|**String**|The character that separates the level specified in Level from the next. Can be one of ".", "-", "+", or "/". The default value is ".".|
| _OnlyLookUpTableCodes_|Optional|**Boolean**|**True** if only codes listed in the lookup table can be used; otherwise **false**. The default value is **false**, which allows additional items to be added.|
| _OnlyCompleteCodes_|Optional|**Boolean**|**True** if only codes with values at all levels of the mask can be used; otherwise **false**. The default value is **false**.|
| _LookupTableLink_|Optional|**Long**|Deprecated in Project and later versions. Local outline codes cannot share lookup tables.
 **Caution**  Do not use  _LookupTableLink_ with the enterprise global or enterprise lookup tables. Data corruption can result.

|
| _OnlyLeaves_|Optional|**Boolean**|**True** if only outline code values without subordinate values can be selected; otherwise **false**. The default is **false**.|
| _MatchGeneric_|Optional|**Boolean**|**True** if Project uses the outline code in the Resource Substitution Wizard; otherwise **false**. The default is **false**.|
| _RequiredCode_|Optional|**Boolean**|**True** if the outline code must be present before save is allowed; otherwise **false**. The default is **false**.|
| _LookupDefault_|Optional|**Boolean**|**True** if the outline code has a default value; otherwise **false**. The default is **false**.|
| _DefaultValue_|Optional|**String**|Specifies the default value of the outline code.|
| _SortOrder_|Optional|**Long**|Specifies whether sorting is ascending, descending, or the lookup table row order. Can be one of the  **[PjListOrder](pjlistorder-enumeration-project.md)** constants. The default is **pjListOrderDefault**.|

### Return Value

 **Boolean**


## Remarks

If only the  _FieldID_ argument is specified, the **CustomOutlineCodeEditEx** method displays the **Code Mask Definition** dialog box for the specified outline code.

The  _OnlyLeaves_,  _MatchGeneric_, and  _RequiredCode_ arguments are available only in Project Professional.


## Example

The following example edits an existing  **Outline Code 1** for tasks, in which the only code mask is "*" for the first level. With default values in the **CustomOutlineCodeEditEx** method, the first command in the example specifies that the second level uses two-digit codes, sorted by number, and is separated from the third level by the "-" character. The second command specifies that the third level uses a single uppercase letter. It also specifies that only codes that contain all three levels can be used.

To use the example, the original  **Outline Code 1** contains the characters "oc1" in the first level. After running the code, the code mask is "*.11-A". A user can edit the lookup table and add, for example, "23" in the level under "oc1" and "X" in the third level. When setting the value of **Outline Code 1**, the user can choose  **oc1.23-X**, but can not choose  **oc1.23**.




```vb
Sub EditCustOutlineCode() 
    CustomOutlineCodeEditEx pjCustomTaskOutlineCode1, Length:=2, _ 
        Separator:="-" 
    CustomOutlineCodeEditEx pjCustomTaskOutlineCode1, Length:=1, _ 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, OnlyCompleteCodes:=True 
End Sub
```

In the following example, the task  **Outline Code 3** contains the lookup table values "a", "b", and "c". Running the example changes the order the user sees when setting the value, to "c", "b", and "a", with the default value "b".




```vb
Sub ChangeOCDefaults() 
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode3, SortOrder:=pjListOrderDescending 
     
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode3, LookupDefault:=True, DefaultValue:="b" 
End Sub
```


