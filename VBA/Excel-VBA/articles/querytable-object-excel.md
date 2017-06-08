---
title: QueryTable Object (Excel)
keywords: vbaxl10.chm517072
f1_keywords:
- vbaxl10.chm517072
ms.prod: excel
api_name:
- Excel.QueryTable
ms.assetid: 505b84ea-64b3-b4fe-741a-de6884eb69eb
ms.date: 06/08/2017
---


# QueryTable Object (Excel)

Represents a worksheet table built from data returned from an external data source, such as an SQL server or a Microsoft Access database.


## Remarks

 The **QueryTable** object is a member of the **[QueryTables](querytables-object-excel.md)** collection.


## Example

Use  **[QueryTables](worksheet-querytables-property-excel.md)** ( _index_ ), where _index_ is the index number of the query table, to return a single **QueryTable** object. The following example sets query table one so that formulas to the right of it are automatically updated whenever it's refreshed.


```
Sheets("sheet1").QueryTables(1).FillAdjacentFormulas = True
```


## Events



|**Name**|
|:-----|
|[AfterRefresh](querytable-afterrefresh-event-excel.md)|
|[BeforeRefresh](querytable-beforerefresh-event-excel.md)|

## Methods



|**Name**|
|:-----|
|[CancelRefresh](querytable-cancelrefresh-method-excel.md)|
|[Delete](querytable-delete-method-excel.md)|
|[Refresh](querytable-refresh-method-excel.md)|
|[ResetTimer](querytable-resettimer-method-excel.md)|
|[SaveAsODC](querytable-saveasodc-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AdjustColumnWidth](querytable-adjustcolumnwidth-property-excel.md)|
|[Application](querytable-application-property-excel.md)|
|[BackgroundQuery](querytable-backgroundquery-property-excel.md)|
|[CommandText](querytable-commandtext-property-excel.md)|
|[CommandType](querytable-commandtype-property-excel.md)|
|[Connection](querytable-connection-property-excel.md)|
|[Creator](querytable-creator-property-excel.md)|
|[Destination](querytable-destination-property-excel.md)|
|[EditWebPage](querytable-editwebpage-property-excel.md)|
|[EnableEditing](querytable-enableediting-property-excel.md)|
|[EnableRefresh](querytable-enablerefresh-property-excel.md)|
|[FetchedRowOverflow](querytable-fetchedrowoverflow-property-excel.md)|
|[FieldNames](querytable-fieldnames-property-excel.md)|
|[FillAdjacentFormulas](querytable-filladjacentformulas-property-excel.md)|
|[ListObject](querytable-listobject-property-excel.md)|
|[MaintainConnection](querytable-maintainconnection-property-excel.md)|
|[Name](querytable-name-property-excel.md)|
|[Parameters](querytable-parameters-property-excel.md)|
|[Parent](querytable-parent-property-excel.md)|
|[PostText](querytable-posttext-property-excel.md)|
|[PreserveColumnInfo](querytable-preservecolumninfo-property-excel.md)|
|[PreserveFormatting](querytable-preserveformatting-property-excel.md)|
|[QueryType](querytable-querytype-property-excel.md)|
|[Recordset](querytable-recordset-property-excel.md)|
|[Refreshing](querytable-refreshing-property-excel.md)|
|[RefreshOnFileOpen](querytable-refreshonfileopen-property-excel.md)|
|[RefreshPeriod](querytable-refreshperiod-property-excel.md)|
|[RefreshStyle](querytable-refreshstyle-property-excel.md)|
|[ResultRange](querytable-resultrange-property-excel.md)|
|[RobustConnect](querytable-robustconnect-property-excel.md)|
|[RowNumbers](querytable-rownumbers-property-excel.md)|
|[SaveData](querytable-savedata-property-excel.md)|
|[SavePassword](querytable-savepassword-property-excel.md)|
|[Sort](querytable-sort-property-excel.md)|
|[SourceConnectionFile](querytable-sourceconnectionfile-property-excel.md)|
|[SourceDataFile](querytable-sourcedatafile-property-excel.md)|
|[TextFileColumnDataTypes](querytable-textfilecolumndatatypes-property-excel.md)|
|[TextFileCommaDelimiter](querytable-textfilecommadelimiter-property-excel.md)|
|[TextFileConsecutiveDelimiter](querytable-textfileconsecutivedelimiter-property-excel.md)|
|[TextFileDecimalSeparator](querytable-textfiledecimalseparator-property-excel.md)|
|[TextFileFixedColumnWidths](querytable-textfilefixedcolumnwidths-property-excel.md)|
|[TextFileOtherDelimiter](querytable-textfileotherdelimiter-property-excel.md)|
|[TextFileParseType](querytable-textfileparsetype-property-excel.md)|
|[TextFilePlatform](querytable-textfileplatform-property-excel.md)|
|[TextFilePromptOnRefresh](querytable-textfilepromptonrefresh-property-excel.md)|
|[TextFileSemicolonDelimiter](querytable-textfilesemicolondelimiter-property-excel.md)|
|[TextFileSpaceDelimiter](querytable-textfilespacedelimiter-property-excel.md)|
|[TextFileStartRow](querytable-textfilestartrow-property-excel.md)|
|[TextFileTabDelimiter](querytable-textfiletabdelimiter-property-excel.md)|
|[TextFileTextQualifier](querytable-textfiletextqualifier-property-excel.md)|
|[TextFileThousandsSeparator](querytable-textfilethousandsseparator-property-excel.md)|
|[TextFileTrailingMinusNumbers](querytable-textfiletrailingminusnumbers-property-excel.md)|
|[TextFileVisualLayout](querytable-textfilevisuallayout-property-excel.md)|
|[WebConsecutiveDelimitersAsOne](querytable-webconsecutivedelimitersasone-property-excel.md)|
|[WebDisableDateRecognition](querytable-webdisabledaterecognition-property-excel.md)|
|[WebDisableRedirections](querytable-webdisableredirections-property-excel.md)|
|[WebFormatting](querytable-webformatting-property-excel.md)|
|[WebPreFormattedTextToColumns](querytable-webpreformattedtexttocolumns-property-excel.md)|
|[WebSelectionType](querytable-webselectiontype-property-excel.md)|
|[WebSingleBlockTextImport](querytable-websingleblocktextimport-property-excel.md)|
|[WebTables](querytable-webtables-property-excel.md)|
|[WorkbookConnection](querytable-workbookconnection-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
