---
title: Work with Tables
keywords: vbapp10.chm5278659
f1_keywords:
- vbapp10.chm5278659
ms.prod: powerpoint
ms.assetid: 1c962e26-8a16-0d88-92bc-58462b31fca9
ms.date: 06/08/2017
---


# Work with Tables

In PowerPoint, you can create native tables without having to import them from Word. Tables are members of the  **Shapes** collection. Each cell, column, and row in a table is a separate programmable object.


## Creating a Table

To create a table on a slide, use the  **AddTable** method. This method adds a table to the **Shapes** collection with the number of rows and columns designated by the **NumRows** and **NumColumns** arguments. This example adds a table with three rows and four columns to slide two.


```vb
ActivePresentation.Slides(2).Shapes _
    .AddTable NumRows:=3, NumColumns:=4, Left:=10, _
    Top:=10, Width:=288, Height:=288
```


## Testing to See Whether a Shape Is a Table

Before you can work with the contents or objects in a table, you must first know if the shape you are working with is a table. To see whether a shape is a table, use the  **HasTable** property. For example, assume that slide one has numerous shapes and you know one of them is a table. You want to resize this table so that it is the proper size to accept the data you are going to import from another source. This code walks through the **Shapes** collection on slide two to find the table and then it resizes the width of the columns.


```vb
With ActivePresentation.Slides(2)
    For sh = 1 To .Shapes.Count
        If .Shapes(sh).HasTable Then
            For Each col In .Shapes(sh).Table.Columns
                col.Width = 110
            Next col
        End If
    Next
End With
```


## Working with Cells, Columns, and Rows

To return the contents and properties of an individual column or row, use a specific member of either the  **Columns** or the **Rows** collection. The **Cell** method returns a single **Cell** object within a **Table**. This example changes various attributes of the table represented by shape five on slide two. It changes the color of row two, the width of column one, and the text contained in the row two, column one cell.


```vb
With ActivePresentation.Slides(2).Shapes(4).Table
    For Each cl In .Rows(2).Cells
        cl.Shape.Fill.ForeColor.RGB = RGB(50, 125, 0)
    Next cl
    .Columns(1).Width = 110
    .Cell(2, 1).Shape.TextFrame.TextRange.Text = "Mallard"
End With

```


