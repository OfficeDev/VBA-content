---
title: Explorer.AddToSelection Method (Outlook)
keywords: vbaol11.chm3309
f1_keywords:
- vbaol11.chm3309
ms.prod: outlook
api_name:
- Outlook.Explorer.AddToSelection
ms.assetid: b85ad121-9e26-0782-3c5e-7651499f8e66
ms.date: 06/08/2017
---


# Explorer.AddToSelection Method (Outlook)

Adds the specified Microsoft Outlook item to the selection in the active explorer.


## Syntax

 _expression_ . **AddToSelection**( **_Item_** )

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item to add to the selection in the active explorer.|

## Remarks

The selection in the active explore is represented by the  **[Selection](selection-object-outlook.md)** object that is returned by the **[Explorer.Selection](explorer-selection-property-outlook.md)** property.

To be selected, the item must be selectable in the current view of the active explorer. Use the  **[IsItemSelectableInView](explorer-isitemselectableinview-method-outlook.md)** method to determine whether the item can be selected in the view. The item does not have to be visible in the view.

Under the following conditions, Outlook returns an error when you call the  **AddToSelection** method:


- The specified item is not in the current view of the active explorer.
    
- The specified item is being edited in the current view of the active explorer.
    
- The current view has been filtered and the application of the filter removed the item from the view.
    
- The specified item has not been saved.
    
- The specified item represents a  **[StorageItem](storageitem-object-outlook.md)** .
    
- No current view exists for the active explorer.
    


If the item is not selected and is selectable in the current view, calling  **AddToSelection** causes the **SelectionChange** event to fire.

Calling  **AddToSelection** does not scroll the view to make the selected item visible in the view and does not expand or collapse groups in the view.

The following table illustrates the results of calling  **AddToSelection** , taking into consideration any current selection (the **[Selection.Count](selection-count-property-outlook.md)** property) and whether the Reading Pane is displayed.



| **Existing** **Selection.Count**| **Reading Pane Displayed**| **Results**|
|0|Yes|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The item is added to the selection.</p></li><li><p><b>SelectionChange</b>  fires.</p></li><li><p>Reading Pane displays the item.</p></li></ul>|
|0|No|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The item is added to the selection.</p></li><li><p><b>SelectionChange</b>  fires.</p></li></ul>|
|>=1|Yes|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The item is added to the selection.</p></li><li><p><b>SelectionChange</b>  fires.</p></li><li><p>Reading Pane does not change the item it displays unless the view is a Calendar view, in which case, calling <b>AddToSelection</b>  can cause the Reading Pane to display a different item.</p></li></ul>|
|>=1|No|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The item is added to the selection.</p></li><li><p><b>SelectionChange</b>  fires.</p></li></ul>|
If the specified item is already selected in the active explorer, calling  **AddToSelection** does not result in any change to the selection and the **SelectionChange** event does not fire.

When you specify an item in a recurring appointment or task as an argument to the  **AddToSelection** method, make sure that before you pass the argument, you obtain an instance of the occurrence by first expanding the recurrences, using the **[IncludeRecurrences](items-includerecurrences-property-outlook.md)** property and the **[Items](items-object-outlook.md)** collection. If you do not expand the recurrences and obtain an occurrence in the series, you pass an instance variable that represents the appointment or task series, and the **AddToSelection** method operates on the series instead of the occurrence.

Note that you can use  **AddToSelection** to add items to a selection, but you cannot add conversation headers to a selection.


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

