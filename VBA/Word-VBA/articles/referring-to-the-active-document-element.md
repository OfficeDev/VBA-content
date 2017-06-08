---
title: Referring to the Active Document Element
keywords: vbawd10.chm5212844
f1_keywords:
- vbawd10.chm5212844
ms.prod: word
ms.assetid: e7eabc36-c1d9-af61-d13a-3d4ac7a01658
ms.date: 06/08/2017
---


# Referring to the Active Document Element

To refer to the active paragraph, table, field, or other document element, use the  **[Selection](application-selection-property-word.md)** property to return a  **[Selection](selection-object-word.md)** object. From the  **Selection** object, you can access all paragraphs in the selection or the first paragraph in the selection. The following example applies a border around the first paragraph in the selection.


```vb
Sub BorderAroundFirstParagraph() 
 Selection.Paragraphs(1).Borders.Enable = True 
End Sub
```


The following example applies a border around each paragraph in the selection.




```vb
Sub BorderAroundSelection() 
 Selection.Paragraphs.Borders.Enable = True 
End Sub
```

The following example applies shading to the first row of the first table in the selection.



```vb
Sub ShadeTableRow() 
 Selection.Tables(1).Rows(1).Shading.Texture = wdTexture10Percent 
End Sub
```

An error occurs if the selection doesn't include a table. Use the  **[Count](tables-count-property-word.md)** property to determine if the selection includes a table. The following example applies shading to the first row of the first table in the selection.



```vb
Sub ShadeTableRow() 
 If Selection.Tables.Count >= 1 Then 
 Selection.Tables(1).Rows(1).Shading.Texture = wdTexture25Percent 
 Else 
 MsgBox "Selection doesn't include a table" 
 End If 
End Sub
```

The following example applies shading to the first row of every table in the selection. The  **For Each...Next** loop is used to step through the individual tables in the selection.



```vb
Sub ShadeAllFirstRowsInTables() 
 Dim tblTable As Table 
 If Selection.Tables.Count >= 1 Then 
 For Each tblTable In Selection.Tables 
 tblTable.Rows(1).Shading.Texture = wdTexture30Percent 
 Next tblTable 
 End If 
End Sub
```


