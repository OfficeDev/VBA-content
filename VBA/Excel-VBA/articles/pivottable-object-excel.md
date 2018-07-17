---
title: PivotTable Object (Excel)
keywords: vbaxl10.chm234072
f1_keywords:
- vbaxl10.chm234072
ms.prod: excel
api_name:
- Excel.PivotTable
ms.assetid: a9c1d4a0-78a9-f9a6-6daf-91cb63e45842
ms.date: 06/08/2017
---


# PivotTable Object (Excel)

Represents a PivotTable report on a worksheet.


## Remarks

 The **PivotTable** object is a member of the **[PivotTables](pivottables-object-excel.md)** collection. The **PivotTables** collection contains all the **PivotTable** objects on a single worksheet.

Because PivotTable report programming can be complex, it's generally easiest to record PivotTable report actions and then revise the recorded code.


## Example

Use  **[PivotTables](worksheet-pivottables-method-excel.md)** ( _index_ ), where _index_ is the PivotTable index number or name, to return a single **PivotTable** object. The following example makes the field named year a row field in the first PivotTable report on Sheet3.


```
Worksheets("Sheet3").PivotTables(1) _ 
 .PivotFields("Year").Orientation = xlRowField
```


## Methods



|**Name**|
|:-----|
|[AddDataField](pivottable-adddatafield-method-excel.md)|
|[AddFields](pivottable-addfields-method-excel.md)|
|[AllocateChanges](pivottable-allocatechanges-method-excel.md)|
|[CalculatedFields](pivottable-calculatedfields-method-excel.md)|
|[ChangeConnection](pivottable-changeconnection-method-excel.md)|
|[ChangePivotCache](pivottable-changepivotcache-method-excel.md)|
|[ClearAllFilters](pivottable-clearallfilters-method-excel.md)|
|[ClearTable](pivottable-cleartable-method-excel.md)|
|[CommitChanges](pivottable-commitchanges-method-excel.md)|
|[ConvertToFormulas](pivottable-converttoformulas-method-excel.md)|
|[CreateCubeFile](pivottable-createcubefile-method-excel.md)|
|[DiscardChanges](pivottable-discardchanges-method-excel.md)|
|[DrillDown](pivottable-drilldown-method-excel.md)|
|[DrillTo](pivottable-drillto-method-excel.md)|
|[DrillUp](pivottable-drillup-method-excel.md)|
|[GetData](pivottable-getdata-method-excel.md)|
|[GetPivotData](pivottable-getpivotdata-method-excel.md)|
|[ListFormulas](pivottable-listformulas-method-excel.md)|
|[PivotCache](pivottable-pivotcache-method-excel.md)|
|[PivotFields](pivottable-pivotfields-method-excel.md)|
|[PivotSelect](pivottable-pivotselect-method-excel.md)|
|[PivotTableWizard](pivottable-pivottablewizard-method-excel.md)|
|[PivotValueCell](pivottable-pivotvaluecell-method-excel.md)|
|[RefreshDataSourceValues](pivottable-refreshdatasourcevalues-method-excel.md)|
|[RefreshTable](pivottable-refreshtable-method-excel.md)|
|[RepeatAllLabels](pivottable-repeatalllabels-method-excel.md)|
|[RowAxisLayout](pivottable-rowaxislayout-method-excel.md)|
|[ShowPages](pivottable-showpages-method-excel.md)|
|[SubtotalLocation](pivottable-subtotallocation-method-excel.md)|
|[Update](pivottable-update-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[ActiveFilters](pivottable-activefilters-property-excel.md)|
|[Allocation](pivottable-allocation-property-excel.md)|
|[AllocationMethod](pivottable-allocationmethod-property-excel.md)|
|[AllocationValue](pivottable-allocationvalue-property-excel.md)|
|[AllocationWeightExpression](pivottable-allocationweightexpression-property-excel.md)|
|[AllowMultipleFilters](pivottable-allowmultiplefilters-property-excel.md)|
|[AlternativeText](pivottable-alternativetext-property-excel.md)|
|[Application](pivottable-application-property-excel.md)|
|[CacheIndex](pivottable-cacheindex-property-excel.md)|
|[CalculatedMembers](pivottable-calculatedmembers-property-excel.md)|
|[CalculatedMembersInFilters](pivottable-calculatedmembersinfilters-property-excel.md)|
|[ChangeList](pivottable-changelist-property-excel.md)|
|[ColumnFields](pivottable-columnfields-property-excel.md)|
|[ColumnGrand](pivottable-columngrand-property-excel.md)|
|[ColumnRange](pivottable-columnrange-property-excel.md)|
|[CompactLayoutColumnHeader](pivottable-compactlayoutcolumnheader-property-excel.md)|
|[CompactLayoutRowHeader](pivottable-compactlayoutrowheader-property-excel.md)|
|[CompactRowIndent](pivottable-compactrowindent-property-excel.md)|
|[Creator](pivottable-creator-property-excel.md)|
|[CubeFields](pivottable-cubefields-property-excel.md)|
|[DataBodyRange](pivottable-databodyrange-property-excel.md)|
|[DataFields](pivottable-datafields-property-excel.md)|
|[DataLabelRange](pivottable-datalabelrange-property-excel.md)|
|[DataPivotField](pivottable-datapivotfield-property-excel.md)|
|[DisplayContextTooltips](pivottable-displaycontexttooltips-property-excel.md)|
|[DisplayEmptyColumn](pivottable-displayemptycolumn-property-excel.md)|
|[DisplayEmptyRow](pivottable-displayemptyrow-property-excel.md)|
|[DisplayErrorString](pivottable-displayerrorstring-property-excel.md)|
|[DisplayFieldCaptions](pivottable-displayfieldcaptions-property-excel.md)|
|[DisplayImmediateItems](pivottable-displayimmediateitems-property-excel.md)|
|[DisplayMemberPropertyTooltips](pivottable-displaymemberpropertytooltips-property-excel.md)|
|[DisplayNullString](pivottable-displaynullstring-property-excel.md)|
|[EnableDataValueEditing](pivottable-enabledatavalueediting-property-excel.md)|
|[EnableDrilldown](pivottable-enabledrilldown-property-excel.md)|
|[EnableFieldDialog](pivottable-enablefielddialog-property-excel.md)|
|[EnableFieldList](pivottable-enablefieldlist-property-excel.md)|
|[EnableWizard](pivottable-enablewizard-property-excel.md)|
|[EnableWriteback](pivottable-enablewriteback-property-excel.md)|
|[ErrorString](pivottable-errorstring-property-excel.md)|
|[FieldListSortAscending](pivottable-fieldlistsortascending-property-excel.md)|
|[GrandTotalName](pivottable-grandtotalname-property-excel.md)|
|[HasAutoFormat](pivottable-hasautoformat-property-excel.md)|
|[Hidden](pivottable-hidden-property-excel.md)|
|[HiddenFields](pivottable-hiddenfields-property-excel.md)|
|[InGridDropZones](pivottable-ingriddropzones-property-excel.md)|
|[InnerDetail](pivottable-innerdetail-property-excel.md)|
|[LayoutRowDefault](pivottable-layoutrowdefault-property-excel.md)|
|[Location](pivottable-location-property-excel.md)|
|[ManualUpdate](pivottable-manualupdate-property-excel.md)|
|[MDX](pivottable-mdx-property-excel.md)|
|[MergeLabels](pivottable-mergelabels-property-excel.md)|
|[Name](pivottable-name-property-excel.md)|
|[NullString](pivottable-nullstring-property-excel.md)|
|[PageFieldOrder](pivottable-pagefieldorder-property-excel.md)|
|[PageFields](pivottable-pagefields-property-excel.md)|
|[PageFieldStyle](pivottable-pagefieldstyle-property-excel.md)|
|[PageFieldWrapCount](pivottable-pagefieldwrapcount-property-excel.md)|
|[PageRange](pivottable-pagerange-property-excel.md)|
|[PageRangeCells](pivottable-pagerangecells-property-excel.md)|
|[Parent](pivottable-parent-property-excel.md)|
|[PivotChart](pivottable-pivotchart-property-excel.md)|
|[PivotColumnAxis](pivottable-pivotcolumnaxis-property-excel.md)|
|[PivotFormulas](pivottable-pivotformulas-property-excel.md)|
|[PivotRowAxis](pivottable-pivotrowaxis-property-excel.md)|
|[PivotSelection](pivottable-pivotselection-property-excel.md)|
|[PivotSelectionStandard](pivottable-pivotselectionstandard-property-excel.md)|
|[PreserveFormatting](pivottable-preserveformatting-property-excel.md)|
|[PrintDrillIndicators](pivottable-printdrillindicators-property-excel.md)|
|[PrintTitles](pivottable-printtitles-property-excel.md)|
|[RefreshDate](pivottable-refreshdate-property-excel.md)|
|[RefreshName](pivottable-refreshname-property-excel.md)|
|[RepeatItemsOnEachPrintedPage](pivottable-repeatitemsoneachprintedpage-property-excel.md)|
|[RowFields](pivottable-rowfields-property-excel.md)|
|[RowGrand](pivottable-rowgrand-property-excel.md)|
|[RowRange](pivottable-rowrange-property-excel.md)|
|[SaveData](pivottable-savedata-property-excel.md)|
|[SelectionMode](pivottable-selectionmode-property-excel.md)|
|[ShowDrillIndicators](pivottable-showdrillindicators-property-excel.md)|
|[ShowPageMultipleItemLabel](pivottable-showpagemultipleitemlabel-property-excel.md)|
|[ShowTableStyleColumnHeaders](pivottable-showtablestylecolumnheaders-property-excel.md)|
|[ShowTableStyleColumnStripes](pivottable-showtablestylecolumnstripes-property-excel.md)|
|[ShowTableStyleLastColumn](pivottable-showtablestylelastcolumn-property-excel.md)|
|[ShowTableStyleRowHeaders](pivottable-showtablestylerowheaders-property-excel.md)|
|[ShowTableStyleRowStripes](pivottable-showtablestylerowstripes-property-excel.md)|
|[ShowValuesRow](pivottable-showvaluesrow-property-excel.md)|
|[Slicers](pivottable-slicers-property-excel.md)|
|[SmallGrid](pivottable-smallgrid-property-excel.md)|
|[SortUsingCustomLists](pivottable-sortusingcustomlists-property-excel.md)|
|[SourceData](pivottable-sourcedata-property-excel.md)|
|[SubtotalHiddenPageItems](pivottable-subtotalhiddenpageitems-property-excel.md)|
|[Summary](pivottable-summary-property-excel.md)|
|[TableRange1](pivottable-tablerange1-property-excel.md)|
|[TableRange2](pivottable-tablerange2-property-excel.md)|
|[TableStyle2](pivottable-tablestyle2-property-excel.md)|
|[Tag](pivottable-tag-property-excel.md)|
|[TotalsAnnotation](pivottable-totalsannotation-property-excel.md)|
|[VacatedStyle](pivottable-vacatedstyle-property-excel.md)|
|[Value](pivottable-value-property-excel.md)|
|[Version](pivottable-version-property-excel.md)|
|[ViewCalculatedMembers](pivottable-viewcalculatedmembers-property-excel.md)|
|[VisibleFields](pivottable-visiblefields-property-excel.md)|
|[VisualTotals](pivottable-visualtotals-property-excel.md)|
|[VisualTotalsForSets](pivottable-visualtotalsforsets-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
