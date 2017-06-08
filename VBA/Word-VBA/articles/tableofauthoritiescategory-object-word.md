---
title: TableOfAuthoritiesCategory Object (Word)
keywords: vbawd10.chm2423
f1_keywords:
- vbawd10.chm2423
ms.prod: word
api_name:
- Word.TableOfAuthoritiesCategory
ms.assetid: ce481ec8-5d5f-fcb8-1d04-5b796accdd3b
ms.date: 06/08/2017
---


# TableOfAuthoritiesCategory Object (Word)

Represents a single table of authorities category. The  **TableOfAuthoritiesCategories** object is a member of the **[TablesOfAuthoritiesCategories](tablesofauthoritiescategories-object-word.md)** collection.


## Remarks

The  **TablesOfAuthoritiesCategories** collection includes all 16 categories listed in the **Category** box on the **Table of Authorities** tab in the **Index and Tables** dialog box.

Use  **TablesOfAuthoritiesCategories** (Index), where Index is the category name or index number, to return a single **TableOfAuthoritiesCategory** object. The following example renames the Rules category as Other Provisions.




```vb
ActiveDocument.TablesOfAuthoritiesCategories("Rules").Name = _ 
 "Other Provisions"
```

The index number represents the position of the category in the  **Index and Tables** dialog box ( **Insert** menu). The following example displays the name of the first category in the **TablesOfAuthoritiesCategories** collection.




```vb
MsgBox ActiveDocument.TablesOfAuthoritiesCategories(1).Name
```

The  **Add** method isn't available for the **TablesOfAuthoritiesCategories** collection. The collection is limited to 16 items; however, you can use the **Name** property to rename an existing category.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

