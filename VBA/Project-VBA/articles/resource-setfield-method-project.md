---
title: Resource.SetField Method (Project)
ms.prod: project-server
api_name:
- Project.Resource.SetField
ms.assetid: 9ac1e770-8716-2954-4459-7f5ff090e2ed
ms.date: 06/08/2017
---


# Resource.SetField Method (Project)

Sets the value of the specified resource custom field.


## Syntax

 _expression_. **SetField**( ** _FieldID_**, ** _Value_** )

 _expression_ A variable that represents a **Resource** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|For a local custom field, can be one of the  **[PjField](pjfield-enumeration-project.md)** constants for resource custom fields. For an enterprise custom field, use the **[FieldNameToFieldConstant](application-fieldnametofieldconstant-method-project.md)** method to get the FieldID.|
| _Value_|Required|**String**|The value of the field.|

## Example

The following example shows how to access an enterprise resource custom field by using the  **SetField** method and the **GetField** method for the **Resource** object together with the **FieldNameToFieldConstant** and **FieldConstantToFieldName** methods.


1. To use the example, use Project Web App to create an enterprise resource text custom field named, for example,  **TestEntResText**. 
    
2. Restart Project Professional with a Project Server profile, so that it includes the new custom field.
    
3. Create a project, build the team from enterprise resources, and then assign a resource to the first task.
    
4. The  **TestEnterpriseResourceCF** macro uses the **FieldNameToFieldConstant** method to find the resourceField number, for example, 205553667.
    
5. The macro shows the number and text value in a message box, by using the  **GetField** method.
    
6. The macro sets a new value for the custom field, by using the  **SetField** method.
    
7. The macro gets the field name by using the  **FieldConstantToFieldName** method, and then shows the field name and new value in another message box.
    





```vb
Sub TestEnterpriseResourceCF() 
    Dim resourceField As Long 
    Dim resourceFieldName As String 
    Dim resourceFieldValue As String 
    Dim message As String 
 
    resourceField = FieldNameToFieldConstant("TestEntResText", pjResource) 
 
    ' Show the enterprise resource field number and old value. 
    message = "Enterprise resource field number: " &; resourceField &; vbCrLf 
    resourceFieldValue = ActiveProject.Tasks(1).Assignments(1).Resource.GetField(resourceField) 

    If resourceFieldValue = "" Then resourceFieldValue = "[No value]" 
    MsgBox message &; "Field value: " &; resourceFieldValue 
 
    ' Set a value for the enterprise resource custom field. 
    ' You can use either the Resources collection or the Assignments collection 
    ' to access the resource custom field. 
    ' Here, use the Assignments collection. 
    ActiveProject.Tasks(1).Assignments(1).Resource.SetField _
        FieldID:=resourceField, Value:="This is a new value." 
 
    ' For a demonstration, get the field name from the number, 
    ' and then verify the new value. 
    resourceFieldName = FieldConstantToFieldName(resourceField) 
 
    ' Here, use the Resources collection to access the custom field. 
    resourceFieldValue = ActiveProject.Resources(1).GetField(resourceField) 
 
    message = "New value for field: " &; resourceFieldName &; vbCrLf 
    MsgBox message &; "Field value: " &; resourceFieldValue 
End Sub
```

For an example that uses a local resource custom field, see the  **[GetField](resource-getfield-method-project.md)** method.


