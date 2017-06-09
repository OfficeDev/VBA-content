---
title: View Object (Outlook)
keywords: vbaol11.chm2479
f1_keywords:
- vbaol11.chm2479
ms.prod: outlook
api_name:
- Outlook.View
ms.assetid: 41c8d149-9912-1685-4c8b-3c849cc6f1ed
ms.date: 06/08/2017
---


# View Object (Outlook)

Represents a customizable view used to sort, group, and view data.


## Remarks

The  **View** object allows you to create customizable views that allow you to better sort, group and ultimately view data of all different types. There are a variety of different view types that provide the flexibility needed to create and maintain your important data.


- The table view type ( **olTableView** ) allows you to view data in a simple field-based table.
    
- The Calendar view type ( **olCalendarView** ) allows you to view data in a calendar format.
    
- The card view type ( **olCardView** ) allows you to view data in a series of cards. Each card displays the information contained by the item and can be sorted.
    
- The icon view type ( **olIconView** ) allows you to view data as icons, similar to a Windows folder or explorer.
    
- The timeline view type ( **olTimelineView** ) allows you to view data as it is received in a customizable linear time line.
    
Views are defined and customized using the  **View** object's **[XML](http://msdn.microsoft.com/library/a933daaa-370f-2ed3-0a59-86f766a1f2c8%28Office.15%29.aspx)** property. The **XML** property allows you to create and set a customized XML schema that defines the various features of a view.

Use  **Views** ( _index_ ), where _index_ is the name of the **View** object or its ordinal value, to return a single **View** object.

Use the  **[Add](http://msdn.microsoft.com/library/8005ca2e-8b28-1286-74d1-448f2a168c65%28Office.15%29.aspx)** method of the **Views** collection to create a new view.

Always use  **[Save](http://msdn.microsoft.com/library/effc4046-2e9c-3898-e37f-c4de817ddde7%28Office.15%29.aspx)** to save a view after you change any property of the view.


## Example

The following example returns a view called Table View and stores it in a variable of type  **View** called objView. Before running this example, make sure a view by the name 'Table View' exists.


```
Sub GetView() 
 
 'Creates a new view 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 'Return a view called Table View 
 
 Set objView = objViews.Item("Table View") 
 
End Sub
```

The following example creates a new view of type  **olTableView** called New Table.




```
Sub CreateView() 
 
 'Creates a new view 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 Set objNewView = objViews.Add(Name:="New Table", _ 
 
 ViewType:=olTableView, SaveOption:=olViewSaveOptionThisFolderEveryone) 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Apply](http://msdn.microsoft.com/library/b121d1ce-24b7-4ace-8369-42e5c7becd0a%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/dfa82ef6-94f1-5c7d-eea5-600f992992d3%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/6d332021-6e93-7665-2a5b-526c927621de%28Office.15%29.aspx)|
|[GoToDate](http://msdn.microsoft.com/library/5ad66fcc-fcdf-9a48-a8e1-669dd294967b%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/fb909688-309d-0a70-0b67-0f1793f6a27d%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/effc4046-2e9c-3898-e37f-c4de817ddde7%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/84fdf8a6-891f-133f-e587-f6d2ced35304%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/416a79d6-bca6-30ae-b119-cba355a1bb77%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/9a4b4b27-d543-df82-3058-e0a6ad2f51a1%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/caa2eb1b-26e3-e8da-c0d8-118d9ba654dc%28Office.15%29.aspx)|
|[LockUserChanges](http://msdn.microsoft.com/library/f4347b6f-b00d-6508-09e3-35cf98da26b1%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/77071683-8f06-7d4a-96ad-5888bea53104%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/96260360-b686-f60a-442e-38eeaaa1d429%28Office.15%29.aspx)|
|[SaveOption](http://msdn.microsoft.com/library/d7990708-5eb4-1b11-944e-127793bdb5b1%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/32c6c27e-2351-c10c-47cd-bcca06d25660%28Office.15%29.aspx)|
|[Standard](http://msdn.microsoft.com/library/99fc4067-29e6-8597-09e7-057d2533b022%28Office.15%29.aspx)|
|[ViewType](http://msdn.microsoft.com/library/db44b9ec-cb55-c9f4-d621-32d2f46598dd%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/a933daaa-370f-2ed3-0a59-86f766a1f2c8%28Office.15%29.aspx)|

## See also


#### Other resources


[View Object Members](http://msdn.microsoft.com/library/ed3196c6-e779-64f7-db1d-e2fd22bb4688%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
