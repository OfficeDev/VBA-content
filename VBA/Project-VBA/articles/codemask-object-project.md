---
title: CodeMask Object (Project)
ms.prod: project-server
ms.assetid: 4d0a22f4-fee9-8f4b-a0c0-7bc817ad3f6a
ms.date: 06/08/2017
---


# CodeMask Object (Project)

The  **CodeMask** object is a collection of **[CodeMaskLevel](codemasklevel-object-project.md)** objects that define the code mask for an outline code in Project.
 


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
|[Add](codemask-add-method-project.md)|

## Properties



|**Name**|
|:-----|
|[Application](codemask-application-property-project.md)|
|[Count](codemask-count-property-project.md)|
|[Item](codemask-item-property-project.md)|
|[Parent](codemask-parent-property-project.md)|

