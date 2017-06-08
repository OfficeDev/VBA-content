---
title: Resource.OvertimeRate Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.OvertimeRate
ms.assetid: 889226c3-8493-3d61-d31d-56cccab8c07c
ms.date: 06/08/2017
---


# Resource.OvertimeRate Property (Project)

Gets or sets the overtime rate of a resource. Read/write  **Variant**.


## Syntax

 _expression_. **OvertimeRate**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **OvertimeRate** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.


## Example

The following example sets the current overtime rate of each resource in the active project to 1.5 times its standard rate.


```vb
Sub SetOverTimeRate() 
 
 Dim R As Resource ' Resource object used in For Each loop 
 Dim StdRate As Double ' Numeric value of resource's standard rate 
 Dim Count As Integer ' Counter used in For Next loop 
 Dim FirstNumber As Integer ' Position of the first number 
 
 For Each R In ActiveProject.Resources 
 ' Find the first character that is a number 
 For Count = 1 To Len(R.StandardRate) 
 If IsNumeric(Mid(R.StandardRate, Count, 1)) Then 
 FirstNumber = Count - 1 
 Exit For 
 End If 
 Next Count 
 
 ' Strip off any leading currency symbol and then use the 
 ' Val function to ignore any characters that follow the number 
 StdRate = Val(Right$(R.StandardRate, Len(R.StandardRate) - FirstNumber)) 
 
 ' Set the overtime rate 
 R.OvertimeRate = 1.5 * StdRate 
 Next R 
 
End Sub
```


