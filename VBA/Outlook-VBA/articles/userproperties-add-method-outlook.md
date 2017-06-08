---
title: UserProperties.Add Method (Outlook)
keywords: vbaol11.chm209
f1_keywords:
- vbaol11.chm209
ms.prod: outlook
api_name:
- Outlook.UserProperties.Add
ms.assetid: 88b86622-2234-77be-41e7-b76b0b3a75ad
ms.date: 06/08/2017
---


# UserProperties.Add Method (Outlook)

Creates a new user property in the  **[UserProperties](userproperties-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Type_** , **_AddToFolderFields_** , **_DisplayFormat_** )

 _expression_ A variable that represents an **UserProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property. The maximum length is 64 characters. The characters, '[', ']', '_' and '#', are not permitted in the name.|
| _Type_|Required| **[OlUserPropertyType](oluserpropertytype-enumeration-outlook.md)**|The type of the new property.|
| _AddToFolderFields_|Optional| **Boolean**| **True** if the property will be added as a custom field to the folder that the item is in. This field can be displayed in the folder's view. **False** if the property will be added as a custom field to the item but not to the folder. The default value is **True** .|
| _DisplayFormat_|Optional| **Long**|Specifies how the property will be displayed in the Outlook user interface. This parameter can be set to a value from one of several different enumerations, determined by the  **OlUserPropertyType** constant specified in the _Type_ parameter. For more information on how _Type_ and _DisplayFormat_ interact, see[DisplayFormat Property](userdefinedproperty-displayformat-property-outlook.md).|

### Return Value

A  **[UserProperty](userproperty-object-outlook.md)** object that represents the new property.


## Remarks

You can define custom properties by calling either the  **UserProperties.Add** method for an Outlook item or folder, or the **[UserDefinedProperties.Add](userdefinedproperties-add-method-outlook.md)** method for a folder.

You can create a property of a type that is defined by the  **OlUserPropertyType** enumeration, except for the following types: **olEnumeration**,  **olOutlookInternal**, and  **olSmartFrom**.

To set for the first time a property created by the  **UserProperties.Add** method, use the **[UserProperty.Value](userproperty-value-property-outlook.md)** property instead of the **[SetProperties](propertyaccessor-setproperties-method-outlook.md)** and **[SetProperty](propertyaccessor-setproperty-method-outlook.md)** methods of the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object.

If you want to view a custom property on an item, you must use the  **UserProperties.Add** method to create that property. Custom properties created by the **[PropertyAccessor](propertyaccessor-object-outlook.md)** are not supported in a custom view.

You cannot add custom properties to Office document items such as Word, Excel, or PowerPoint files. You will receive an error when you try to programmatically add a user-defined field to a  **[DocumentItem](documentitem-object-outlook.md)** object.


## Example

This VBA example creates a new  **[ContactItem](contactitem-object-outlook.md)** object and adds "LastDateSpokenWith" as a custom property.


```vb
Sub AddUserProperty() 
 Dim myItem As Outlook.ContactItem 
 Dim myUserProperty As Outlook.UserProperty 
 
 Set myItem = Application.CreateItem(olContactItem) 
 Set myUserProperty = myItem.UserProperties _ 
 .Add("LastDateSpokenWith", olDateTime) 
 myItem.Display 
End Sub
```

This VBA example creates a new  **ContactItem** object and adds "Details" as a user property. The value is set by changing the **[Value](userproperty-value-property-outlook.md)** property of the **UserProperty** object.




```vb
Sub AddUserProperty() 
 Dim myItem As Outlook.ContactItem 
 Dim myUserProperty As Outlook.UserProperty 
 
 Set myItem = Application.CreateItem(olContactItem) 
 Set myUserProperty = myItem.UserProperties _ 
 .Add("Details", olText) 
 myUserProperty.Value = "Neighbor" 
 myItem.Display 
End Sub
```


## See also


#### Concepts


[UserProperties Object](userproperties-object-outlook.md)

