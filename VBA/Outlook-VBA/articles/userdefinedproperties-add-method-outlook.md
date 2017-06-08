---
title: UserDefinedProperties.Add Method (Outlook)
keywords: vbaol11.chm588
f1_keywords:
- vbaol11.chm588
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperties.Add
ms.assetid: e033b27e-101d-4ef8-ed84-790fd9e6107a
ms.date: 06/08/2017
---


# UserDefinedProperties.Add Method (Outlook)

Creates a new  **[UserDefinedProperty](userdefinedproperty-object-outlook.md)** object and appends it to the collection.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Type_** , **_DisplayFormat_** , **_Formula_** )

 _expression_ A variable that represents a **UserDefinedProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new user-defined property.|
| _Type_|Required| **[OlUserPropertyType](oluserpropertytype-enumeration-outlook.md)**|The type of the new user-defined property.|
| _DisplayFormat_|Optional| **Variant**|The display format of the new user-defined property. This parameter can be set to a value from one of several different enumerations, determined by the  **OlUserPropertyType** constant specified in the _Type_ parameter. For more information on how _Type_ and _DisplayFormat_ interact, see[DisplayFormat Property](userdefinedproperty-displayformat-property-outlook.md).|
| _Formula_|Optional| **Variant**|The formula used to calculate values for the new user-defined property. This parameter is ignored if the  _Type_ parameter is set to any value other than **olCombination** or **olFormula** .|

### Return Value

A  **UserDefinedProperty** object that represents the new user-defined property.


## Remarks

You can create a property of a type that is defined by the  **OlUserPropertyType** enumeration, except for the following types: **olEnumeration**,  **olOutlookInternal**, and  **olSmartFrom**.


## Example

The following Visual Basic for Applications (VBA) example uses the  **Add** method to create and add several **UserDefinedProperty** objects to the **Inbox** default folder.


```vb
Sub AddStatusProperties() 
 Dim objNamespace As NameSpace 
 Dim objFolder As Folder 
 Dim objProperty As UserDefinedProperty 
 
 ' Obtain a Folder object reference to the 
 ' Inbox default folder. 
 Set objNamespace = Application.GetNamespace("MAPI") 
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox) 
 
 ' Add five user-defined properties, used to identify and 
 ' track customer issues. 
 With objFolder.UserDefinedProperties 
 Set objProperty = .Add("Issue?", olYesNo, olFormatYesNoIcon) 
 Set objProperty = .Add("Issue Research Time", olDuration) 
 Set objProperty = .Add("Issue Resolution Time", olDuration) 
 Set objProperty = .Add("Customer Follow-Up", olYesNo, olFormatYesNoYesNo) 
 Set objProperty = .Add("Issue Closed", olYesNo, olFormatYesNoYesNo) 
 End With 
End Sub
```


## See also


#### Concepts


[UserDefinedProperties Object](userdefinedproperties-object-outlook.md)

