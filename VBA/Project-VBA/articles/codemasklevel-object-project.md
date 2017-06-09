---
title: CodeMaskLevel Object (Project)
ms.prod: project-server
api_name:
- Project.CodeMaskLevel
ms.assetid: cef1b15f-c7f1-3b95-49a1-00854a74d9da
ms.date: 06/08/2017
---


# CodeMaskLevel Object (Project)

Represents a level in the code mask of an outline code definition. The  **CodeMaskLevel** object is a member of the **[CodeMask](codemask-object-project.md)** collection.
 


## Example

The following example adds three levels to a code mask.
 

 

```
Sub DefineLocationCodeMask(objCodeMask As CodeMask) 
 
 objCodeMask.Add _ 
 Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
 Length:=2, Separator:="." 
 
 objCodeMask.Add _ 
 Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
 Separator:="." 
 
 objCodeMask.Add _ 
 Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
 Length:=3, Separator:="." 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](codemasklevel-delete-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](codemasklevel-application-property-project.md)|
|[Index](codemasklevel-index-property-project.md)|
|[Length](codemasklevel-length-property-project.md)|
|[Level](codemasklevel-level-property-project.md)|
|[Parent](codemasklevel-parent-property-project.md)|
|[Separator](codemasklevel-separator-property-project.md)|
|[Sequence](codemasklevel-sequence-property-project.md)|

