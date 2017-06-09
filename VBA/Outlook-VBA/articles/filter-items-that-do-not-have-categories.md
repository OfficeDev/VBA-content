---
title: Filter Items that Do Not Have Categories
ms.prod: outlook
ms.assetid: d351052d-6cc5-85ac-9791-c7b8ccfc5282
ms.date: 06/08/2017
---


# Filter Items that Do Not Have Categories

This topic shows a code sample that uses a DAV Searching and Locating (DASL) query to filter items in the current folder that do not have any category assigned to them. Note that filtering items with an empty string in their categories requires a DASL query; the Microsoft Jet syntax does not support such filters.

When filtering an empty string with a DASL query, you can use the  **Is Null** keyword. **Is Null** operations are useful to determine if a string property is empty or if a date property has been set. For more information, see [Filtering Items Using Query Keywords](filtering-items-using-query-keywords.md).

The code sample sets up a DASL filter on the  **Categories** property, which in the DASL query is expressed in the Office namespace as **urn:schemas-microsoft-com:office:office#Keywords**. The filter compares the value of the  **Categories** property with an emptry string using the **Is Null** keyword. The code sample then applies the filer to items in the current folder. It then prints the number of items in the current folder that have been found to have no categories.




```vb
Sub NullCategoryRestriction() 
 Dim oFolder As Outlook.Folder 
 Dim oItems As Outlook.Items 
 Dim Filter As String 
 
 'DASL Filter can test for null property. 
 'This will return all items that have no category. 
 Filter = "@SQL=" &; Chr(34) &; _ 
 "urn:schemas-microsoft-com:office:office#Keywords" &; _ 
 Chr(34) &; " is null" 
 Set oFolder = Application.ActiveExplorer.CurrentFolder 
 Set oItems = oFolder.Items.Restrict(Filter) 
 Debug.Print oItems.Count 
End Sub
```


