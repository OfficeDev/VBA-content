---
title: Returning an Object from a Collection (Excel)
keywords: vbaxl10.chm5204603
f1_keywords:
- vbaxl10.chm5204603
ms.prod: excel
ms.assetid: f8a36459-f9dd-9f4c-ef7a-b188173434d5
ms.date: 06/08/2017
---


# Returning an Object from a Collection (Excel)

The  **Item** property of a collection returns a single object from that collection. The following example sets the `firstBook` variable to a **[Workbook](workbook-object-excel.md)** object that represents the first workbook in the  **[Workbooks](workbooks-object-excel.md)** collection.


```vb
Set FirstBook = Workbooks.Item(1)
```


The  **Item** property is the default property for most collections, so you can write the same statement more concisely by omitting the **Item** keyword.




```vb
Set FirstBook = Workbooks(1)
```

For more information about a specific collection, see the Help topic for that collection or the  **Item** property for the collection.

## Named Objects

Although you can usually specify an integer value with the  **Item** property, it may be more convenient to return an object by name. Before you can use a name with the **Item** property, you must name the object. Most often, this is done by setting the object's **Name** property. The following example creates a named worksheet in the active workbook and then refers to the worksheet by name.


```vb
ActiveWorkbook.Worksheets.Add.Name = "A New Sheet" 
With Worksheets("A New Sheet") 
 .Range("A5:A10").Formula = "=RAND()" 
End With
```


## Predefined Index Values

Some collections have predefined index values you can use to return single objects. Each predefined index value is represented by a constant. For example, you specify an  **XlBordersIndex** constant with the **Item** property of the **Borders** collection to return a single border.

The following example sets the bottom border of cells A1:G1 on Sheet1 to a double line.




```vb
Worksheets("Sheet1").Range("A1:A1"). _ 
 Borders.Item(xlEdgeBottom).LineStyle = xlDouble
```


