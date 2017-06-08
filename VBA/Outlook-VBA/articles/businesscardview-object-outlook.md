---
title: BusinessCardView Object (Outlook)
keywords: vbaol11.chm3212
f1_keywords:
- vbaol11.chm3212
ms.prod: outlook
api_name:
- Outlook.BusinessCardView
ms.assetid: 83706cf8-080c-fbf0-9381-5801a2dd4dfd
ms.date: 06/08/2017
---


# BusinessCardView Object (Outlook)

Represents a view that displays data as a series of Electronic Business Card (EBC) images.


## Remarks

The  **BusinessCardView** object, derived from the **[View](view-object-outlook.md)** object, allows you to create customizable views that allow you to better sort, group and ultimately view contact items in Outlook as a series of Electronic Business Cards, each of which displays the contact information for an Outlook contact item based on the EBC design associated with the contact item.

Use the  **[Add](views-add-method-outlook.md)** method of the **[Views](views-object-outlook.md)** collection to add a new **BusinessCardView** to a **[Folder](folder-object-outlook.md)** object.

Use the  **[Filter](businesscardview-filter-property-outlook.md)** property to determine which Outlook contact items to display in the view, the **[CardSize](businesscardview-cardsize-property-outlook.md)** property to specify the size of each Electronic Business Card in the view, and the **[HeadingsFont](businesscardview-headingsfont-property-outlook.md)** to retrieve the **[ViewFont](viewfont-object-outlook.md)** object for the view. Use the **[LockUserChanges](businesscardview-lockuserchanges-property-outlook.md)** property to allow or prevent changes to the user interface for the view.


## Example

The following Visual Basic for Applications (VBA) example creates, saves, and applies a new  **BusinessCardView** object.


```
Sub CreateBusinessCardView() 
 
 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As BusinessCardView 
 
 
 
 ' Get the Views collection of the Inbox default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderContacts).Views 
 
 
 
 ' Create the new view. 
 
 Set objView = objViews.Add( _ 
 
 "Card View", _ 
 
 olBusinessCardView, _ 
 
 olViewSaveOptionAllFoldersOfType) 
 
 
 
 ' Save and apply the new view. 
 
 objView.Save 
 
 objView.Apply 
 
 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Apply](businesscardview-apply-method-outlook.md)|
|[Copy](businesscardview-copy-method-outlook.md)|
|[Delete](businesscardview-delete-method-outlook.md)|
|[GoToDate](businesscardview-gotodate-method-outlook.md)|
|[Reset](businesscardview-reset-method-outlook.md)|
|[Save](businesscardview-save-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](businesscardview-application-property-outlook.md)|
|[CardSize](businesscardview-cardsize-property-outlook.md)|
|[Class](businesscardview-class-property-outlook.md)|
|[Filter](businesscardview-filter-property-outlook.md)|
|[HeadingsFont](businesscardview-headingsfont-property-outlook.md)|
|[Language](businesscardview-language-property-outlook.md)|
|[LockUserChanges](businesscardview-lockuserchanges-property-outlook.md)|
|[Name](businesscardview-name-property-outlook.md)|
|[Parent](businesscardview-parent-property-outlook.md)|
|[SaveOption](businesscardview-saveoption-property-outlook.md)|
|[Session](businesscardview-session-property-outlook.md)|
|[SortFields](businesscardview-sortfields-property-outlook.md)|
|[Standard](businesscardview-standard-property-outlook.md)|
|[ViewType](businesscardview-viewtype-property-outlook.md)|
|[XML](businesscardview-xml-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
