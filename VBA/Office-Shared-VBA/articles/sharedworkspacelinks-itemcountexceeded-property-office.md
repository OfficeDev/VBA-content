---
title: SharedWorkspaceLinks.ItemCountExceeded Property (Office)
keywords: vbaof11.chm271005
f1_keywords:
- vbaof11.chm271005
ms.prod: office
api_name:
- Office.SharedWorkspaceLinks.ItemCountExceeded
ms.assetid: 53d5ab73-4d7a-7cf1-07d5-3dd5598fb1c5
ms.date: 06/08/2017
---


# SharedWorkspaceLinks.ItemCountExceeded Property (Office)

Gets a  **Boolean** value that indicates whether the number of **SharedWorkspaceLinks** items in the collection has exceeded the 99 that can be displayed in the Shared Workspace task pane. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **ItemCountExceeded**

 _expression_ A variable that represents a **SharedWorkspaceLinks** object.


### Return Value

Boolean


## Remarks

The  **Shared Workspace** task pane can only display 99 shared workspace files and folders, links, members, or tasks. If more than 99 items are added to any of these collections: the corresponding tab of the **Shared Workspace** task pane stops displaying the list of items and displays a link to the shared workspace site Web page instead; the collection is no longer populated locally and its **Count** property returns 0 (zero).

Furthermore, once the  **ItemCountExceeded** property returns **True** for one of the collections listed above, the developer can no longer remedy the situation programmatically by deleting items from the collection to reduce the count below 99, because the collection is no longer populated.


## Example

The following example checks the Count property of the  **SharedWorkspaceLinks** collection. If **Count** returns 0 (zero), it checks the **ItemCountExceeded** property to determine whether in fact the shared workspace has no saved links, or whether it has more than 99 and the links collection has been cleared.


```
ActiveWorkbook.SharedWorkspace.Refresh 
    If ActiveWorkbook.SharedWorkspace.Links.Count = 0 Then 
        If ActiveWorkbook.SharedWorkspace.Links.ItemCountExceeded Then 
            MsgBox "More than 99 links in shared workspace.", _ 
                vbInformation + vbOKOnly, "Item Count Exceeded" 
        Else 
            MsgBox "No links in shared workspace.", _ 
                vbInformation + vbOKOnly, "No Links" 
        End If 
    End If
```


## See also


#### Concepts


[SharedWorkspaceLinks Object](sharedworkspacelinks-object-office.md)
#### Other resources


[SharedWorkspaceLinks Object Members](sharedworkspacelinks-members-office.md)

