
# Cell.SetWidth Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets the width of columns or cells in a table.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SetWidth**( **_ColumnWidth_**,  **_RulerStyle_**)

 _expression_Required. A variable that represents a  ** [Cell](cbe6ae71-b2da-63a9-1446-0a2f81ab8b14.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ColumnWidth|Required| **Single**|The width of the specified column or columns, in points.|
|RulerStyle|Required| **WdRulerStyle**|Controls the way Word adjusts cell widths.|

## Remarks
<a name="sectionSection1"> </a>

The  **WdRulerStyle** behavior described above applies to left-aligned tables. The **WdRulerStyle** behavior for center- and right-aligned tables can be unexpected; in these cases, the **SetWidth** method should be used with care.


## Example
<a name="sectionSection2"> </a>

This example creates a table in a new document and sets the width of the first cell in the second row to 1.5 inches. The example preserves the widths of the other cells in the table.


```
Set newDoc = Documents.Add 
Set myTable = _ 
 newDoc.Tables.Add(Range:=Selection.Range, NumRows:=3, _ 
 NumColumns:=3) 
myTable.Cell(2,1).SetWidth _ 
 ColumnWidth:=InchesToPoints(1.5), _ 
 RulerStyle:=wdAdjustNone
```

This example sets the width of the cell that contains the insertion point to 36 points. The example also narrows the first column to preserve the position of the right edge of the table.




```
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).SetWidth ColumnWidth:=36, _ 
 RulerStyle:=wdAdjustFirstColumn 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Cell Object](cbe6ae71-b2da-63a9-1446-0a2f81ab8b14.md)
#### Other resources


 [Cell Object Members](f718bcaa-af8a-682b-f403-6db1aeb9bb73.md)
