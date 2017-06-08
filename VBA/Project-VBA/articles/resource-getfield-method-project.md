---
title: Resource.GetField Method (Project)
ms.prod: project-server
api_name:
- Project.Resource.GetField
ms.assetid: 36fbbc13-272e-72f4-ebbe-2c13f67abbe7
ms.date: 06/08/2017
---


# Resource.GetField Method (Project)

Returns the value of the specified resource custom field.


## Syntax

 _expression_. **GetField**( ** _FieldID_** )

 _expression_ A variable that represents a **Resource** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|For a local custom field, can be one of the  **[PjField](pjfield-enumeration-project.md)** constants for resource custom fields. For an enterprise custom field, use the **[FieldNameToFieldConstant](application-fieldnametofieldconstant-method-project.md)** method to get the FieldID.|

### Return Value

 **String**


## Example

The following example displays the value of a local resource custom field specified by the user.


```vb
Sub DisplayField() 
    Dim Temp As String 
 
    Temp = InputBox$("Enter the name of the field you want to see:") 
    Temp = LCase(Temp) 
 
    Select Case Temp 
        Case "name" 
            MsgBox (ActiveCell.Resource.GetField(FieldID:=pjResourceName)) 
        Case "initials" 
            MsgBox (ActiveCell.Resource.GetField(FieldID:=pjResourceInitials)) 
        Case "standard rate" 
            MsgBox (ActiveCell.Resource.GetField(FieldID:=pjResourceStandardRate)) 
        Case "" 
            End 
        Case Else 
            MsgBox "You entered an invalid field. Please try again." 
            End 
    End Select 
End Sub
```

For an example that uses an enterprise resource custom field, see the  **[SetField](resource-setfield-method-project.md)** method.


