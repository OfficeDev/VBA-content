---
title: Columns Object (PowerPoint)
keywords: vbapp10.chm623000
f1_keywords:
- vbapp10.chm623000
ms.prod: powerpoint
api_name:
- PowerPoint.Columns
ms.assetid: ba2fb830-bb60-b259-3a3f-1281f77d6368
ms.date: 06/08/2017
---


# Columns Object (PowerPoint)

A collection of  **[Column](column-object-powerpoint.md)** objects that represent the columns in a table.


## Example

Use the  **Columns** property to return the **Columns** collection. This example finds the first table in the active presentation, counts the number of **Column** objects in the **Columns** collection, and displays information to the user.


```vb
Dim ColCount, sl, sh As Integer

With ActivePresentation
    For sl = 1 To .Slides.Count
        For sh = 1 To .Slides(sl).Shapes.Count
            If .Slides(sl).Shapes(sh).HasTable Then
                ColCount = .Slides(sl).Shapes(sh) _
                    .Table.Columns.Count
                MsgBox "Shape " &; sh &; " on slide " &; sl &; _
                    " contains the first table and has " &; _
                    ColCount &; " columns."
                Exit Sub
            End If
        Next
    Next
End With
```

Use the [Add](columns-add-method-powerpoint.md)method to add a column to a table. This example creates a column in an existing table and sets the width of the new column to 72 points (one inch).




```vb
With ActivePresentation.Slides(2).Shapes(5).Table

    .Columns.Add.Width = 72

End With
```

Use  **Columns** (index) to return a single **Column** object. Index represents the position of the column in the **Columns** collection (usually counting from left to right; although the[TableDirection](table-tabledirection-property-powerpoint.md)property can reverse this). This example selects the first column of the table in shape five on the second slide.




```vb
ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Select
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

