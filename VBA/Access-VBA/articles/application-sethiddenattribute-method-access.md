---
title: Application.SetHiddenAttribute Method (Access)
keywords: vbaac10.chm12571
f1_keywords:
- vbaac10.chm12571
ms.prod: access
api_name:
- Access.Application.SetHiddenAttribute
ms.assetid: b92a1edc-033a-095c-980f-852b8f7e0785
ms.date: 06/08/2017
---


# Application.SetHiddenAttribute Method (Access)

The  **SetHiddenAttribute** method sets the hidden attribute of an Access object.


## Syntax

 _expression_. **SetHiddenAttribute**( ** _ObjectType_**, ** _ObjectName_**, ** _fHidden_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcObjectType**|A  **[AcObjectType](acobjecttype-enumeration-access.md)** constant that specifies the type of Access object.|
| _ObjectName_|Required|**String**|The name of the Access object.|
| _fHidden_|Required|**Boolean**|**True** sets the hidden attribute and **False** clears the attribute.|

### Return Value

Nothing


## Remarks

Together with the  **GetHiddenAttribute** method, the **SetHiddenAttribute** method provides a means of changing an object's visibility from Visual Basic code. With these methods, you can set or read the Hidden property available in the object's **Properties** dialog box.

To set this option by using the  **SetHiddenAttribute** method, specify **True** or **False** for the setting, as in the following example.




```vb
Application.SetHiddenAttribute acTable,"Customers", True
```


## See also


#### Concepts


[Application Object](application-object-access.md)

