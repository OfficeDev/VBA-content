---
title: TableObject Object (Excel)
keywords: vbaxl10.chm915072
f1_keywords:
- vbaxl10.chm915072
ms.prod: excel
ms.assetid: afc981f4-155b-085a-3c17-c8d46c4d7037
ms.date: 06/08/2017
---


# TableObject Object (Excel)

Represents a worksheet table built from data returned from a PowerPivot model.


## Example

The following sample code creates a PowerPivot query table by connecting to a data source.


```
Sub CreateTable()
Dim objWBConnection As WorkbookConnection
Dim objWorksheet As Worksheet
Dim objTable As TableObject   'This is the new Table object

Set objWorksheet = ActiveWorkbook.Worksheets("Sheet1")

'Create a WorkbookConnection to the external data source first.
Set objWBConnection = ActiveWorkbook.Connections.Add2( _
        "Cubes3 AdventureWorksDW DimEmployee1", "", Array( _
        "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=AdventureWorksDW;Data Source=MyServer;Use " _
        , _
        "Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MYWORKSTATION;Use Encryption for Data=False;Tag with co" _
        , "lumn collation when possible=False"), Array( _
        """AdventureWorksDW"".""dbo"".""DimEmployee"""), 3, True)

'Create a new table connected to the model.
Set objTable = objWorksheet.ListObjects.Add(SourceType:=xlSrcModel, Source:=objWBConnection, Destination:=Range("$A$1")).TableObject

objTable.Refresh

End Sub

```


## Methods



|**Name**|
|:-----|
|[Delete](tableobject-delete-method-excel.md)|
|[Refresh](tableobject-refresh-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AdjustColumnWidth](tableobject-adjustcolumnwidth-property-excel.md)|
|[Application](tableobject-application-property-excel.md)|
|[Creator](tableobject-creator-property-excel.md)|
|[Destination](tableobject-destination-property-excel.md)|
|[EnableEditing](tableobject-enableediting-property-excel.md)|
|[EnableRefresh](tableobject-enablerefresh-property-excel.md)|
|[FetchedRowOverflow](tableobject-fetchedrowoverflow-property-excel.md)|
|[ListObject](tableobject-listobject-property-excel.md)|
|[Parent](tableobject-parent-property-excel.md)|
|[PreserveColumnInfo](tableobject-preservecolumninfo-property-excel.md)|
|[PreserveFormatting](tableobject-preserveformatting-property-excel.md)|
|[RefreshStyle](tableobject-refreshstyle-property-excel.md)|
|[ResultRange](tableobject-resultrange-property-excel.md)|
|[RowNumbers](tableobject-rownumbers-property-excel.md)|
|[WorkbookConnection](tableobject-workbookconnection-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
