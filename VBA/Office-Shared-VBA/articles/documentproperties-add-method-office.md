---
title: DocumentProperties.Add Method (Office)
keywords: vbaof11.chm250014
f1_keywords:
- vbaof11.chm250014
ms.prod: office
api_name:
- Office.DocumentProperties.Add
ms.assetid: 80738562-8b0b-33f1-3dfa-0d66b1844ef7
ms.date: 06/08/2017
---


# DocumentProperties.Add Method (Office)

Creates a new custom document property. You can add a new document property only to the custom **DocumentProperties** collection.


## Syntax

_expression_. **Add** (**_Name_**, **_LinkToContent_**, **_Type_**, **_Value_**, **_LinkSource_**)

_expression_ Required. A variable that represents a **[DocumentProperties](documentproperties-object-office.md)** object. The custom **DocumentProperties** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The string of the [Name](http://msdn.microsoft.com/library/b609c38e-71ca-e019-9852-fc7811dc798f.md%28Office.15%29.aspx) of the property.|
| _LinkToContent_|Required|**Boolean**|Specifies whether the [LinkToContent](http://msdn.microsoft.com/library/062df6df-cdee-81fc-3244-e229dacaa64e.md%28Office.15%29.aspx) property is linked to the contents of the container document. If this argument is **True**, the _LinkSource_ argument is required; if it's **False**, the value argument is required.|
| _Type_|Optional|**Variant**|The data type of the [Type](documentproperty-type-property-office.md) property. Can be one of the following **MsoDocProperties** constants: **msoPropertyTypeBoolean**, **msoPropertyTypeDate**, **msoPropertyTypeFloat**, **msoPropertyTypeNumber**, or **msoPropertyTypeString**.|
| _Value_|Optional|**Variant**|The data value of the [Value](documentproperty-value-property-office.md) property, if it's not linked to the contents of the container document. The value is converted to match the data type specified by the _Type_ argument, and if it can't be converted, an error occurs. If _LinkToContent_ is **True**, the argument is ignored and the new document property is assigned a default value until the linked property values are updated by the container application (usually when the document is saved).|
| _LinkSource_|Optional|**Variant**|Ignored if  _LinkToContent_ is **False**. The source of the [LinkSource](http://msdn.microsoft.com/library/3e3a6ebc-615a-298e-c40f-cbb6d5cf63e3.md%28Office.15%29.aspx) property. The container application determines what types of source linking you can use. For example, DDE links use the "Server\|Document!Item" syntax.|

<br/>

## Remarks

If you add a custom document property to the **DocumentProperties** collection that's linked to a given value in an Office document, you must save the document to see the change to the **DocumentProperty** object.

## Example

This example, which is designed to run in Word, adds three custom document properties to the **DocumentProperties** collection.

```
With ActiveDocument.CustomDocumentProperties 
    .Add Name:="LastModifiedBy", _ 
        LinkToContent:=True, _ 
        Type:=msoPropertyTypeString, _ 
        LinkSource:=Author
    .Add Name:="CustomNumber", _ 
        LinkToContent:=False, _ 
        Type:=msoPropertyTypeNumber, _ 
        Value:=1000 
    .Add Name:="CustomString", _ 
        LinkToContent:=False, _ 
        Type:=msoPropertyTypeString, _ 
        Value:="This is a custom property." 
    .Add Name:="CustomDate", _ 
        LinkToContent:=False, _ 
        Type:=msoPropertyTypeDate, _ 
        Value:=Date 
End With
```

<br/>

## See also

#### Concepts

- [DocumentProperties Object](documentproperties-object-office.md)

#### Other resources

- [DocumentProperties Object Members](documentproperties-members-office.md)

