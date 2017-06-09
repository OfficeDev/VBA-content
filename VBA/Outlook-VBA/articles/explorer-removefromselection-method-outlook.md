---
title: Explorer.RemoveFromSelection Method (Outlook)
keywords: vbaol11.chm3310
f1_keywords:
- vbaol11.chm3310
ms.prod: outlook
api_name:
- Outlook.Explorer.RemoveFromSelection
ms.assetid: f31bc78f-500e-2f73-ea14-8d5f19cd44e9
ms.date: 06/08/2017
---


# Explorer.RemoveFromSelection Method (Outlook)

Cancels the selection of the specified Microsoft Outlook item in the active explorer.


## Syntax

 _expression_ . **RemoveFromSelection**( **_Item_** )

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item to be removed from the selection.|

## Remarks

The selection in the active explore is represented by the  **[Selection](selection-object-outlook.md)** object returned by the **[Explorer.Selection](explorer-selection-property-outlook.md)** property.

To be removed from a selection, an item must be selectable in the current view of the active explorer. However, the item does not have to be visible in the view.

Outlook will return an error when you call the  **RemoveFromSelection** method under the following conditions:


- The specified item is not in the current view of the active explorer.
    
- The specified item is being edited in the current view of the active explorer.
    
- The current view has been filtered and the application of the filter removed the item from the view.
    
- The specified item has not been saved.
    
- The specified item represents a  **[StorageItem](storageitem-object-outlook.md)** .
    
- The current view is a conversation view.
    
- No current view exists for the active explorer.
    


If the specified item is selected, calling  **RemoveFromSelection** will cause the **[SelectionChange](explorer-selectionchange-event-outlook.md)** event to fire. If the item is not selected, calling **RemoveFromSelection** will not cause the **SelectionChange** event to fire.

Calling  **RemoveFromSelection** does not scroll the view to make the specified item visible in the view and does not expand or collapse groups in the view.

The following table illustrates the results of calling  **RemoveFromSelection** , taking into consideration any current selection (the **[Selection.Count](selection-count-property-outlook.md)** property), whether the Reading Pane is displayed, and whether the specified item is displayed in the Reading Pane.



| **Existing** **Selection.Count**| **Reading Pane Displayed**| **Specified Item Displayed in Reading Pane**| **Results**|
|1|Yes|Yes|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The selection is cleared.</p></li><li><p><b>SelectionChange</b>  fires.</p></li><li><p>Reading Pane is empty.</p></li></ul>|
|>1|Yes|No|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The item is removed from the selection.</p></li><li><p><b>SelectionChange</b>  fires.</p></li><li><p>Reading Pane does not change.</p></li></ul>|
|>1|Yes|Yes|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The item is removed from the selection.</p></li><li><p><b>SelectionChange</b>  fires.</p></li><li><p>Reading Pane displays the next item or adjacent item in the selection.</p></li></ul>|
|>=1|No|N/A|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The item is removed from the selection.</p></li><li><p><b>SelectionChange</b>  fires.</p></li></ul>|
If the specified item exists in the current view but is not selected in that view, calling  **RemoveFromSelection** does not result in any change to the selection and does not fire the **SelectionChange** event.

When you specify an item in a recurring appointment or task as an argument to the  **RemoveFromSelection** method, make sure that before you pass the argument, you obtain an instance of the occurrence by first expanding the recurrences, using the **[IncludeRecurrences](items-includerecurrences-property-outlook.md)** property and the **[Items](items-object-outlook.md)** collection. If you do not expand the recurrences and obtain an occurrence in the series, you would be passing an instance variable that represents the appointment or task series, and the **RemoveFromSelection** method would be operating on the series instead of the occurrence.


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

