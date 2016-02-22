
# PivotTable.CalculatedMembers Property (Excel)

Returns a  **[CalculatedMembers](3c664ac6-e2f8-f631-006d-6a16c380641e.md)** collection representing all the calculated members and calculated measures for an OLAP PivotTable.


## Syntax

 _expression_ . **CalculatedMembers**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

This property is used for Online Analytical Processing (OLAP) sources; a non-OLAP source will return a run-time error.


## Example

This example adds a set to the PivotTable. It assumes a PivotTable exists on the active worksheet that is connected to an OLAP data source which contains a field titled "[Product].[All Products]".


```vb
Sub UseCalculatedMember() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Add the calculated member. 
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _ 
 Formula:="'{[Product].[All Products].Children}'", _ 
 Type:=xlCalculatedSet 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](a9c1d4a0-78a9-f9a6-6daf-91cb63e45842.md)
#### Other resources


[PivotTable Object Members](8e8d1692-cf32-63c6-a1f6-54ddcc2a4964.md)
