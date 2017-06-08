---
title: Lists Object (Word)
ms.prod: word
ms.assetid: 1fd927c5-6186-5ca0-80ae-c2ab225d092c
ms.date: 06/08/2017
---


# Lists Object (Word)

A collection of  **List** objects that represent all the lists in the specified document.


## Remarks

Use the  **Lists** property to return the **Lists** collection. The following example displays the number of items in each list in the active document.


```
For Each li In ActiveDocument.Lists 
 MsgBox li.CountNumberedItems 
Next li
```

Use  **Lists** (Index), where Index is the index number, to return a single **[List](list-object-word.md)** object. The following example applies the first list format (excluding **None**) on the  **Numbered** tab in the **Bullets and Numbering** dialog box to the second list in the active document.




```
Set temp1 = ListGalleries(wdNumberGallery).ListTemplates(1) 
ActiveDocument.Lists(2).ApplyListTemplate ListTemplate:=temp1
```

When you use a  **For Each** loop to enumerate the **Lists** collection, the lists in a document are returned in reverse order. The following example counts the items for each list in the active document, from the bottom of the document upward.




```
For Each li In ActiveDocument.Lists 
 MsgBox li.CountNumberedItems 
Next li
```

To add a new list to a document, use the  **ApplyListTemplate** method with the **[ListFormat](listformat-object-word.md)** object for a specified range.

You can manipulate the individual  **[List](list-object-word.md)** objects within a document, but for more precise control you should work with the **ListFormat** object.


 **Note**  Picture-bulleted lists are not included in the  **Lists** collection.


## Methods



|**Name**|
|:-----|
|[Item](lists-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](lists-application-property-word.md)|
|[Count](lists-count-property-word.md)|
|[Creator](lists-creator-property-word.md)|
|[Parent](lists-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
