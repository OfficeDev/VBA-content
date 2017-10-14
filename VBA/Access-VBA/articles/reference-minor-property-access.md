---
title: Reference.Minor Property (Access)
keywords: vbaac10.chm12633
f1_keywords:
- vbaac10.chm12633
ms.prod: access
api_name:
- Access.Reference.Minor
ms.assetid: 7c227db9-9b75-92e5-d32d-e3fda027c145
ms.date: 06/08/2017
---


# Reference.Minor Property (Access)

The  **Minor** property of a **[Reference](reference-object-access.md)** object returns a **Long** value indicating the minor version number of the application to which you have set a reference.


## Syntax

 _expression_. **Minor**

 _expression_ A variable that represents a **Reference** object.


## Remarks

The  **Minor** property returns the value to the right of the decimal point in a version number. For example, if you've set a reference to an application whose version number is 2.5, the **Minor** property returns 5.


## Example

The following example displays a message with information about all the references in the current project.


```vb
Dim r As Reference 
Dim strInfo As String 
 
For Each r In Application.References 
 strInfo = strInfo &; r.Name &; " " &; r.Major &; "." &; r.Minor &; vbCrLf 
Next 
 
 
MsgBox "Current References: " &; vbCrLf &; strInfo
```


## See also


#### Concepts


[Reference Object](reference-object-access.md)

