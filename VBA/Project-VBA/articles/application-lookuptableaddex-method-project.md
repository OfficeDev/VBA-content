---
title: Application.LookUpTableAddEx Method (Project)
keywords: vbapj.chm635
f1_keywords:
- vbapj.chm635
ms.prod: project-server
api_name:
- Project.Application.LookUpTableAddEx
ms.assetid: 5f316f1e-de4b-2fe4-6d3e-84a9944adaed
ms.date: 06/08/2017
---


# Application.LookUpTableAddEx Method (Project)

Appends items to the lookup table of a custom outline code definition.


## Syntax

 _expression_. **LookUpTableAddEx**( ** _FieldID_**, ** _Level_**, ** _Code_**, ** _Description_**, ** _Phonetic_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|Specifies the custom outline code to edit. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _Level_|Optional|**Long**|Specifies the level of the new code. The default value is the level of the last item in the lookup table.|
| _Code_|Optional|**String**|The code to be added to the lookup table.|
| _Description_|Optional|**String**|A description for the field specified in the Code argument.|
| _Phonetic_|Optional|**String**|The phonetic spelling of the Code argument, used for sorting order in Japanese. For languages other than Japanese, Phonetic is ignored.|

### Return Value

 **Boolean**


## Remarks

If only the FieldID argument is specified, the  **LookUpTableAddEx** method displays the **Lookup Table** dialog box for the specified custom outline code.


## Example

This example shows how it is possible to create an invalid entry in a lookup table. The first line correctly adds a new code to the second level of a two-level code mask. The second line, however, causes a problem in the lookup table because the appended code doesn't match the mask for the code; that is, it adds the new code at the third level of a two-level mask.


```vb
Sub LookupTableProblem() 
 Application.LookUpTableAddEx pjCustomTaskOutlineCode1, Level:=2, Code:="Q" 
 Application.LookUpTableAddEx pjCustomTaskOutlineCode1, Level:=3, Code:="Z" 
End Sub
```


