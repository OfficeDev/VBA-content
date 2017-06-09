---
title: Application.GetHiddenAttribute Method (Access)
keywords: vbaac10.chm12570
f1_keywords:
- vbaac10.chm12570
ms.prod: access
api_name:
- Access.Application.GetHiddenAttribute
ms.assetid: aee0e022-08d5-10f8-bfd0-588b5310fb43
ms.date: 06/08/2017
---


# Application.GetHiddenAttribute Method (Access)

The  **GetHiddenAttribute** method returns the value of hidden attribute of a Microsoft Access object in the object's **Properties** dialog box, available by selecting the object in the Database window and clicking **Properties** on the **View** menu.


## Syntax

 _expression_. **GetHiddenAttribute**( ** _ObjectType_**, ** _ObjectName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcObjectType**|A  **[AcObjectType](acobjecttype-enumeration-access.md)** constant that specifies the type of Access object.|
| _ObjectName_|Required|**String**|The name of the Access object.|

### Return Value

Boolean


## Remarks

The  **GetHiddenAttribute** method (along with the **SetHiddenAttribute** method) provide a means of changing an object's hidden attribute from Visual Basic code. With these methods, you can set or read the hidden option available in the object's **Properties** dialog box.

Since the hidden attributes that the user can set by selecting or clearing a check box, the  **GetHiddenAttribute** method returns **True** if the option setting is Yes (the check box is selected) or **False** if the option setting is No (the check box is cleared). For example, to set an option of this kind by using the **SetHiddenAttribute** method, specify **True** or **False** for the setting argument, as in the following:




```vb
Application.SetHiddenAttribute acTable,"Customers", True
```


## See also


#### Concepts


[Application Object](application-object-access.md)

