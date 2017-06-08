---
title: Resource.MaterialLabel Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.MaterialLabel
ms.assetid: 802fd00b-3f0e-9ecf-6cb9-a8858f0137a0
ms.date: 06/08/2017
---


# Resource.MaterialLabel Property (Project)

Gets or sets the label for a material resource. Read/write  **String**.


## Syntax

 _expression_. **MaterialLabel**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **MaterialLabel** property does not return any meaningful information for non-material resources, such as people or machines. Setting a value returns a trappable error (error code 1101) when applied to non-material resources.


## Example

This example goes through the list of resources in the current project and sets the material label for all material resources to "pallet." (The error trapping in this example is only to illustrate how you might handle an expected exception. In a real-life example, you would probably include a test such as the following: 


```
If InStr(R.Name, "bricks") <> 0 Then...
```

The test would ensure that you only assign the material label to paving bricks, red bricks, and so on.




```vb
Sub FixLabels() 
 Dim R As Resource 
 
 On Error GoTo ErrTrap: 
 
 For Each R In ActiveProject.Resources 
 If R.MaterialLabel <> "pallet" Then R.MaterialLabel = "pallet" 
 Next R 
 
 Exit Sub 
 
ErrTrap: 
 If Err.Number = 1101 Then 
 Err.Clear 
 Resume Next 
 Else 
 MsgBox Err.Description, vbExclamation, "Error" 
 End If 
End Sub
```


