---
title: MetaProperty.Validate Method (Office)
keywords: vbaof11.chm275007
f1_keywords:
- vbaof11.chm275007
ms.prod: office
api_name:
- Office.MetaProperty.Validate
ms.assetid: e8037c82-a9bd-936f-fbf1-03c35d83685b
ms.date: 06/08/2017
---


# MetaProperty.Validate Method (Office)

Validates a  **MetaProperty** object representing a single property value according to a schema.


## Syntax

 _expression_. **Validate**

 _expression_ An expression that returns a **MetaProperty** object.


### Return Value

String


## Remarks

If the property is invalid, the test fails and an error message is returned. The schema used for validation is stored as part of the document's Microsoft SharePoint Foundation profile.


## Example

In the following example, a  **MetaProperties** object is passed to a validation function. The function then validates the value of a single property represented by its index and returns the result.


```
Function ValidateMetaProperty(ByVal metaProps As MetaProperties) As String 
Dim result As String 
 
result = metaProps(1).Validate 
ValidateMetaProperty = result 
End Function
```


## See also


#### Concepts


[MetaProperty Object](metaproperty-object-office.md)
#### Other resources


[MetaProperty Object Members](metaproperty-members-office.md)

