---
title: DoCmd.DeleteObject Method (Access)
keywords: vbaac10.chm4147
f1_keywords:
- vbaac10.chm4147
ms.prod: access
api_name:
- Access.DoCmd.DeleteObject
ms.assetid: 8e59c5a8-89bd-0d90-9fd1-a1178c73c1c1
ms.date: 06/08/2017
---


# DoCmd.DeleteObject Method (Access)

The  **DeleteObject** method carries out the DeleteObject action in Visual Basic.


## Syntax

 _expression_. **DeleteObject**( ** _ObjectType_**, ** _ObjectName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Optional|**AcObjectType**|A  **[AcObjectType](acobjecttype-enumeration-access.md)** constant that represents the type of object to delete.|
| _ObjectName_|Optional|**Variant**| string expression that's the valid name of an object of the type selected by the _objecttype_ argument. If you run Visual Basic code containing the **DeleteObject** method in a library database, Microsoft Access looks for the object with this name first in the library database, then in the current database.|

## Remarks

You can use the  **DeleteObject** method to delete a specified database object.

If you leave the  _objecttype_ and _objectname_ arguments blank (the default constant, **acDefault**, is assumed for _objecttype_), Microsoft Access deletes the object selected in the Database window. To select an object in the Database window, you can use the SelectObject action or  **SelectObject** method with the In Database Window argument set to Yes ( **True** ).


## Example

The following example deletes the specified table:


```vb
DoCmd.DeleteObject acTable, "Former Employees Table"
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

