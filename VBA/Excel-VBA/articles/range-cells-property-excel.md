---
title: Range.Cells Property (Excel)
keywords: vbaxl10.chm144091
f1_keywords:
- vbaxl10.chm144091
ms.prod: excel
api_name:
- Excel.Range.Cells
ms.assetid: 32a6ecc7-2366-2cec-1feb-0966241a435d
ms.date: 06/08/2017
---


# Range.Cells Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the cells in the specified range.


## Syntax

 _expression_ . **Cells**

 _expression_ A variable that represents a **Range** object.


## Remarks

Because the  **[Item](range-item-property-excel.md)** property is the default property for the **Range** object, you can specify the row and column index immediately after the **Cells** keyword. For more information, see the **Item** property and the examples for this topic.

Using this property without an object qualifier returns a  **Range** object that represents all the cells on the active worksheet.


## Example

This example sets the font style for cells A1:C5 on Sheet1 to italic.


```vb
Worksheets("Sheet1").Activate 
Range(Cells(1, 1), Cells(5, 3)).Font.Italic = True
```

This example scans a column of data named "myRange." If a cell has the same value as the cell immediately above it, the example displays the address of the cell that contains the duplicate data.




```vb
Set r = Range("myRange") 
For n = 2 To r.Rows.Count 
    If r.Cells(n-1, 1) = r.Cells(n, 1) Then 
        MsgBox "Duplicate data in " &; r.Cells(n, 1).Address 
    End If 
Next n
```

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1)

This example looks through column C, and for every cell that has a comment, it puts the comment text into column D and deletes the comment from column C.




```vb
Sub SplitComments()
   'Set up your variables
   Dim cmt As Comment
   Dim iRow As Integer
   
   'Go through all the cells in Column C, and check to see if the cell has a comment.
   For iRow = 1 To WorksheetFunction.CountA(Columns(3))
      Set cmt = Cells(iRow, 3).Comment
      If Not cmt Is Nothing Then
      
         'If there is a comment, paste the comment text into column D and delete the original comment.
         Cells(iRow, 4) = Cells(iRow, 3).Comment.Text
         Cells(iRow, 3).Comment.Delete
      End If
   Next iRow
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Range Object](range-object-excel.md)

