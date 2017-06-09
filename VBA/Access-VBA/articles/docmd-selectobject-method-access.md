---
title: DoCmd.SelectObject Method (Access)
keywords: vbaac10.chm4178
f1_keywords:
- vbaac10.chm4178
ms.prod: access
api_name:
- Access.DoCmd.SelectObject
ms.assetid: def1bac5-57b1-0b2c-d39a-f0c10962880c
ms.date: 06/08/2017
---


# DoCmd.SelectObject Method (Access)

The  **SelectObject** method carries out the SelectObject action in Visual Basic.


## Syntax

 _expression_. **SelectObject**( ** _ObjectType_**, ** _ObjectName_**, ** _InNavigationPane_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcObjectType**|A  **[AcObjectType](acobjecttype-enumeration-access.md)** constant that specifies the type of object you want to select.|
| _ObjectName_|Optional|**Variant**|A string expression that's the valid name of an object of the type selected by the  _objecttype_ argument. This is a required argument, unless you specify **True** (?1) for the _InNavigationPane_ argument.|
| _InNavigationPane_|Optional|**Variant**|Use  **True** to select the object in the Database window. Use **False** (0) to select an object that's already open. If you leave this argument blank, the default ( **False** ) is assumed.|

## Remarks

You can use the  **SelectObject** method to select a specified database object.

The SelectObject action works with any Access object that can receive the focus. This action gives the specified object the focus and shows the object if it's hidden.


## Example

The following example selects the form Customers in the Database window:


```vb
DoCmd.SelectObject acForm, "Customers", True
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

