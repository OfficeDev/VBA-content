---
title: TablesOfAuthoritiesCategories Object (Word)
keywords: vbawd10.chm2422
f1_keywords:
- vbawd10.chm2422
ms.prod: word
ms.assetid: 344b9c42-01d1-805c-6af6-c8301e24b97e
ms.date: 06/08/2017
---


# TablesOfAuthoritiesCategories Object (Word)

A collection of  **[TableOfAuthoritiesCategory](tableofauthoritiescategory-object-word.md)** objects that represent the table of authorities categories, such as Cases and Statutes. The **TablesOfAuthoritiesCategories** collection includes all 16 categories listed in the **Category** box on the **Table of Authorities** tab in the **Index and Tables** dialog box.


## Remarks

Use the  **TablesOfAuthoritiesCategories** property to return the **TablesOfAuthoritiesCategories** collection. The following example displays the names of the categories in the **TablesOfAuthoritiesCategories** collection.


```vb
For Each aCat In ActiveDocument.TablesOfAuthoritiesCategories 
 response = MsgBox(Prompt:=aCat, Buttons:=vbOKCancel) 
 If response = vbCancel Then Exit For 
Next aCat
```

The  **Add** method isn't available for the **TablesOfAuthoritiesCategories** collection. The collection is limited to 16 items; however, you can use the **Name** property to rename an existing category.

Use  **TablesOfAuthoritiesCategories** (Index), where Index is the category name or index number, to return a single **TableOfAuthoritiesCategory** object. The following example renames the Rules category as Other Provisions.




```vb
ActiveDocument.TablesOfAuthoritiesCategories("Rules").Name = _ 
 "Other Provisions"
```

The index number represents the position of the category in the  **Index and Tables** dialog box. The following example displays the name of the first category in the **TablesOfAuthoritiesCategories** collection.




```vb
MsgBox ActiveDocument.TablesOfAuthoritiesCategories(1).Name
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

