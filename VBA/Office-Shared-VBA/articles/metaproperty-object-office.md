---
title: MetaProperty Object (Office)
keywords: vbaof11.chm275000
f1_keywords:
- vbaof11.chm275000
ms.prod: office
api_name:
- Office.MetaProperty
ms.assetid: 4379d183-9b80-92d8-1dd0-ac9be400e366
ms.date: 06/08/2017
---


# MetaProperty Object (Office)

Represents a single property in a collection of properties describing the metadata stored in a document.


## Example

In the following example, a  **MetaProperties** object is passed to a validation function. The function then validates the value of a single property represented by its index and returns the result.


```
Function ValidateMetaProperty(ByVal metaProps As MetaProperties) As String 
Dim result As String 
 
result = metaProps(1).Validate 
ValidateMetaProperty = result 
End Function 

```


## Methods



|**Name**|
|:-----|
|[Validate](metaproperty-validate-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](metaproperty-application-property-office.md)|
|[Creator](metaproperty-creator-property-office.md)|
|[Id](metaproperty-id-property-office.md)|
|[IsReadOnly](metaproperty-isreadonly-property-office.md)|
|[IsRequired](metaproperty-isrequired-property-office.md)|
|[Name](metaproperty-name-property-office.md)|
|[Parent](metaproperty-parent-property-office.md)|
|[Type](metaproperty-type-property-office.md)|
|[Value](metaproperty-value-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
