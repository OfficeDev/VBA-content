---
title: UserDefinedProperties Object (Outlook)
keywords: vbaol11.chm3152
f1_keywords:
- vbaol11.chm3152
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperties
ms.assetid: 196e5d4c-22be-02d3-95e0-3ea7594c2e4b
ms.date: 06/08/2017
---


# UserDefinedProperties Object (Outlook)

Contains a set of  **[UserDefinedProperty](userdefinedproperty-object-outlook.md)** objects representing the user-defined properties defined for a **[Folder](folder-object-outlook.md)** object.


## Remarks

The members of the  **UserDefinedProperties** collection correspond to the fields under **User-defined fields in folder** that you get in the **Show Fields** dialog.

Use the  **[UserDefinedProperties](folder-userdefinedproperties-property-outlook.md)** property to retrieve the **UserDefinedProperties** object from a **Folder** object.

Use the  **[Add](userdefinedproperties-add-method-outlook.md)** method to define and add a user-defined property to, and the **[Remove](userdefinedproperties-remove-method-outlook.md)** method to remove an existing user-defined property from, the **UserDefinedProperties** collection. Use the **[Item](userdefinedproperties-item-method-outlook.md)** method to retrieve by name or index, or the **[Find](userdefinedproperties-find-method-outlook.md)** method to locate and retrieve by name, a **UserDefinedProperty** object from the **UserDefinedProperties** collection. Use the **[Refresh](userdefinedproperties-refresh-method-outlook.md)** method to reload the **UserDefinedProperties** collection from the store.

The  **UserDefinedProperties** collection contains only the definitions of user-defined properties, which are applicable to all Outlook items contained by the folder. To retrieve or change user-defined property values for an Outlook item in that folder, use the **[UserProperties](mailitem-userproperties-property-outlook.md)** property of the Outlook item, such as a **[MailItem](mailitem-object-outlook.md)** object, to retrieve the **[UserProperties](userproperties-object-outlook.md)** collection for that item. You can then use the **[UserProperty](userproperty-object-outlook.md)** object for the appropriate user-defined property to retrieve or change the value of that user-defined property for the Outlook item.


## Example

The following Visual Basic for Applications (VBA) example uses the  **Add** method to create and add several **UserDefinedProperty** objects to the **Inbox** default folder.


```
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


## Methods



|**Name**|
|:-----|
|[Add](userdefinedproperties-add-method-outlook.md)|
|[Find](userdefinedproperties-find-method-outlook.md)|
|[Item](userdefinedproperties-item-method-outlook.md)|
|[Refresh](userdefinedproperties-refresh-method-outlook.md)|
|[Remove](userdefinedproperties-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](userdefinedproperties-application-property-outlook.md)|
|[Class](userdefinedproperties-class-property-outlook.md)|
|[Count](userdefinedproperties-count-property-outlook.md)|
|[Parent](userdefinedproperties-parent-property-outlook.md)|
|[Session](userdefinedproperties-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
