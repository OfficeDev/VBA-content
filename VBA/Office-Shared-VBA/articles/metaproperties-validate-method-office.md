---
title: MetaProperties.Validate Method (Office)
keywords: vbaof11.chm274004
f1_keywords:
- vbaof11.chm274004
ms.prod: office
api_name:
- Office.MetaProperties.Validate
ms.assetid: 658532c6-c8c0-ff01-3736-4161a09af2bb
ms.date: 06/08/2017
---


# MetaProperties.Validate Method (Office)

Validates all of the properties in a  **MetaProperties** collection object according to a schema.


## Syntax

 _expression_. **Validate**

 _expression_ An expression that returns a **MetaProperties** object.


### Return Value

String


## Remarks

If any of the properties is invalid, the test fails and an error message is returned. The schema used for validation is stored as part of the document's Microsoft SharePoint Foundation profile.


## Example

In the following example, a  **MetaProperties** object is passed to a validation function. The function then validates all of the properties of the object and returns the result.


```
Function ValidateMetaProperties(ByVal metaProps As MetaProperties) As String 
Dim result As String 
 
result = metaProps.Validate 
ValidateMetaProperties = result 
End Function
```


## See also


#### Concepts


[MetaProperties Object](metaproperties-object-office.md)
#### Other resources


[MetaProperties Object Members](metaproperties-members-office.md)

