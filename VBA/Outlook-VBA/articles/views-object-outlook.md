---
title: Views Object (Outlook)
keywords: vbaol11.chm3013
f1_keywords:
- vbaol11.chm3013
ms.prod: outlook
api_name:
- Outlook.Views
ms.assetid: 5dd7edc2-12a2-f4c2-d158-8053d80e8dc9
ms.date: 06/08/2017
---


# Views Object (Outlook)

Contains a collection of all  **[View](view-object-outlook.md)** objects in the current folder.


## Remarks

Use the  **Views** property of the **[Folder](folder-object-outlook.md)** object to return the **Views** collection. Use **Views** ( _index_ ),where _index_ is the object's name or position within the collection, to return a single **View** object.

Use the  **[Add](views-add-method-outlook.md)** method of the views collection to add a new view to the collection.

Use the  **[Remove](views-remove-method-outlook.md)** method to remove a view from the collection.


## Example

The following example returns a  **View** object of type **olTableView** called Table View. Before running this example, make sure a view by the name 'Table View' exists.


```
Sub GetView() 
 
 'Returns a view called Table View 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 'Return a view called Table View 
 
 Set objView = objViews.Item("Table View") 
 
End Sub
```

The following example adds a new view of type  **olIconView** in the user's Notes folder.


 **Note**  The  **Add** method will fail if a view with the same name already exists.




```
Sub CreateView() 
 
 'Creates a new view 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderNotes).Views 
 
 Set objNewView = objViews.Add(Name:="New Icon View Type", _ 
 
 ViewType:=olIconView, SaveOption:=olViewSaveOptionThisFolderEveryone) 
 
 
 
End Sub
```

 The following example removes the above view, "New Icon View Type", from the collection.




```
Sub DeleteView() 
 
 'Deletes a view from the collection 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderNotes).Views 
 
 objViews.Remove ("New Icon View Type") 
 
End Sub
```


## Events



|**Name**|
|:-----|
|[ViewAdd](views-viewadd-event-outlook.md)|
|[ViewRemove](views-viewremove-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Add](views-add-method-outlook.md)|
|[Item](views-item-method-outlook.md)|
|[Remove](views-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](views-application-property-outlook.md)|
|[Class](views-class-property-outlook.md)|
|[Count](views-count-property-outlook.md)|
|[Parent](views-parent-property-outlook.md)|
|[Session](views-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
