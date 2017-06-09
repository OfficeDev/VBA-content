---
title: Items.Sort Method (Outlook)
keywords: vbaol11.chm72
f1_keywords:
- vbaol11.chm72
ms.prod: outlook
api_name:
- Outlook.Items.Sort
ms.assetid: 7cb248a2-6885-8be5-df7b-fd5683081e01
ms.date: 06/08/2017
---


# Items.Sort Method (Outlook)

Sorts the collection of items by the specified property. The index for the collection is reset to 1 upon completion of this method.


## Syntax

 _expression_ . **Sort**( **_Property_** , **_Descending_** )

 _expression_ A variable that represents an **Items** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Property_|Required| **String**|The name of the property by which to sort, which may be enclosed in brackets, for example, "[CompanyName]". User-defined properties that contain spaces must be enclosed in brackets. May not be a user-defined property of type keywords, and may not be a multi-valued property, such as a category. For user-defined properties, the property must exist in the  **UserDefinedProperties** collection for **[Items.Parent](items-parent-property-outlook.md)** , which represents the **[Folder](folder-object-outlook.md)** object that contains the items.|
| _Descending_|Optional| **Variant**| **True** to sort in descending order. The default value is **False** (ascending).|

## Remarks

 **Sort** only affects the order of items in a collection. It does not affect the order of items in an explorer view.

 **Sort** cannot be used and will cause an error if the _Property_ paramater is one of the following properties:



| **Categories**| **[LastFirstSpaceOnly](contactitem-lastfirstspaceonly-property-outlook.md)**|
| **[Children](contactitem-children-property-outlook.md)**| **[LastFirstSpaceOnlyCompany](contactitem-lastfirstspaceonlycompany-property-outlook.md)**|
| **Class**| **[MemberCount](distlistitem-membercount-property-outlook.md)**|
| **[CompanyLastFirstNoSpace](contactitem-companylastfirstnospace-property-outlook.md)**| **[NetMeetingAlias](contactitem-netmeetingalias-property-outlook.md)**|
| **[CompanyLastFirstSpaceOnly](contactitem-companylastfirstspaceonly-property-outlook.md)**| **[RecurrenceState](appointmentitem-recurrencestate-property-outlook.md)**|
| **[DLName](distlistitem-dlname-property-outlook.md)**| **[ResponseState](taskitem-responsestate-property-outlook.md)**|
| **[LastFirstAndSuffix](contactitem-lastfirstandsuffix-property-outlook.md)**| **Saved**|
| **[LastFirstNoSpace](contactitem-lastfirstnospace-property-outlook.md)**| **Sent**|
| **[LastFirstNoSpaceCompany](contactitem-lastfirstnospacecompany-property-outlook.md)**||

## Example

The following Visual Basic for Applications (VBA) example uses the  **Sort** method to sort the **[Items](items-object-outlook.md)** collection for the default **Tasks** folder by the "DueDate" property and displays the due dates each in turn.


```vb
Sub SortByDueDate() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myFolder As Outlook.Folder 
 Dim myItem As Outlook.TaskItem 
 Dim myItems As Outlook.Items 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderTasks) 
 Set myItems = myFolder.Items 
 myItems.Sort "[DueDate]", False 
 For Each myItem In myItems 
 MsgBox myItem.Subject &; "-- " &; myItem.DueDate 
 Next myItem 
End Sub
```


## See also


#### Concepts


[Items Object](items-object-outlook.md)

