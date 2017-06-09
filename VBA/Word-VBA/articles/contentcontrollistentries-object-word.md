---
title: ContentControlListEntries Object (Word)
keywords: vbawd10.chm3524
f1_keywords:
- vbawd10.chm3524
ms.prod: word
api_name:
- Word.ContentControlListEntries
ms.assetid: 74b90054-e0a3-37c5-40d2-dc6dd6389cc5
ms.date: 06/08/2017
---


# ContentControlListEntries Object (Word)

The  **ContentControlListEntries** collection contains **ContentControlListEntry** objects that represent the items in a drop-down list or combo box content control.


## Remarks

Use the  **[Add](contentcontrollistentries-add-method-word.md)** method to add an item to a drop-down list or combo box. The following code example uses the **Add** method to add several list items to a new drop-down list content control.


```vb
Dim objCC As ContentControl Dim objLE As ContentControlListEntry Dim objMap As XMLMapping  Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)  'List items objCC.DropdownListEntries.Add "Cat" objCC.DropdownListEntries.Add "Dog" objCC.DropdownListEntries.Add "Equine" objCC.DropdownListEntries.Add "Monkey" objCC.DropdownListEntries.Add "Snake" objCC.DropdownListEntries.Add "Other"
```

Use the  **[Item](contentcontrollistentries-item-method-word.md)** method or the **[DropdownListEntries](contentcontrol-dropdownlistentries-property-word.md)** property of a **ContentControl** object to access an individual list item within a collection. The following code example uses the **Item** method to access the third item in a list and change the display text.


 **Note**  The following code example assumes that the first  **ContentControl** object in the active document is a drop-down list or combo box.




```vb
Dim objCC As ContentControl Dim objLE As ContentControlListEntry Dim objMap As XMLMapping  Set objCC = ActiveDocument.ContentControls(1) Set objLE = objCC.DropdownListEntries.Item(3) objLE.Text = "Horse"
```

Use the  **Clear** method to remove all items from a drop-down list or combo box. The following code example clears all items from the first content control in the active document.


 **Note**  The following code example assumes that the first content control in the active document is a drop-down list or combo box.




```vb
Dim objCC As ContentControl  Set objCC = ActiveDocument.ContentControls(1)  objCC.DropdownListEntries.Clear
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


