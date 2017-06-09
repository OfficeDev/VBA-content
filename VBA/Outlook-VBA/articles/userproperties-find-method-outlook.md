---
title: UserProperties.Find Method (Outlook)
keywords: vbaol11.chm210
f1_keywords:
- vbaol11.chm210
ms.prod: outlook
api_name:
- Outlook.UserProperties.Find
ms.assetid: 3b71ce5a-4bb0-fdab-a24e-02c631816b80
ms.date: 06/08/2017
---


# UserProperties.Find Method (Outlook)

Locates and returns a  **[UserProperty](userproperty-object-outlook.md)** object for the requested property name, if it exists.


## Syntax

 _expression_ . **Find**( **_Name_** , **_Custom_** )

 _expression_ A variable that represents an **UserProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the requested property.|
| _Custom_|Optional| **Variant**| **True** if custom properties on the item should be searched, **False** if built-in properties should be searched.|

### Return Value

If you use  **UserProperties.Find** to look for a custom property and the call succeeds, it will return a **UserProperty** object. If it fails, it will return **Null** ( **Nothing** in Visual Basic). If you use **UserProperties.Find** to look for a built-in property, specify **False** for the _Custom_ parameter. If the call succeeds, it will return the property as a **UserProperty** object. If the call fails, it will return **Null** ( **Nothing** in Visual Basic). If you specify **True** for _Custom_ , the call will not find the built-in property and will return **Null** ( **Nothing** in Visual Basic).


## Remarks

If  _Custom_ parameter is **True** , only custom user properties will be searched. The default value is **True** . To find a non custom property such as **Subject** , specify _Custom_ parameter as **False** , otherwise will return **Nothing** .


## See also


#### Concepts


[UserProperties Object](userproperties-object-outlook.md)

