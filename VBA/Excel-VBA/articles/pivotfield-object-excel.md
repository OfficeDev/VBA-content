---
title: PivotField Object (Excel)
keywords: vbaxl10.chm239072
f1_keywords:
- vbaxl10.chm239072
ms.prod: excel
api_name:
- Excel.PivotField
ms.assetid: 52784960-e2da-b43a-1e37-2d4dae61c6d8
ms.date: 06/08/2017
---


# PivotField Object (Excel)

Represents a field in a PivotTable report.


## Remarks

 The **PivotField** object is a member of the **[PivotFields](pivotfields-object-excel.md)** collection. The **PivotFields** collection contains all the fields in a PivotTable report, including hidden fields.

In some cases, it may be easier to use one of the properties that returns a subset of the PivotTable fields. The following properties are available:


-  **[ColumnFields](pivottable-columnfields-property-excel.md)** property
    
-  **[DataFields](pivottable-datafields-property-excel.md)** property
    
-  **[HiddenFields](pivottable-hiddenfields-property-excel.md)** property
    
-  **[PageFields](pivottable-pagefields-property-excel.md)** property
    
-  **[RowFields](pivottable-rowfields-property-excel.md)** property
    
-  **[VisibleFields](pivottable-visiblefields-property-excel.md)** property
    

## Example

Use  **[PivotFields](pivottable-pivotfields-method-excel.md)** ( _index_ ), where _index_ is the field name or index number, to return a single **PivotField** object. The following example makes the Year field a row field in the first PivotTable report on Sheet3.


```
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").Orientation = xlRowField
```


## Methods



|**Name**|
|:-----|
|[AddPageItem](pivotfield-addpageitem-method-excel.md)|
|[AutoShow](pivotfield-autoshow-method-excel.md)|
|[AutoSort](pivotfield-autosort-method-excel.md)|
|[CalculatedItems](pivotfield-calculateditems-method-excel.md)|
|[ClearAllFilters](pivotfield-clearallfilters-method-excel.md)|
|[ClearLabelFilters](pivotfield-clearlabelfilters-method-excel.md)|
|[ClearManualFilter](pivotfield-clearmanualfilter-method-excel.md)|
|[ClearValueFilters](pivotfield-clearvaluefilters-method-excel.md)|
|[Delete](pivotfield-delete-method-excel.md)|
|[DrillTo](pivotfield-drillto-method-excel.md)|
|[PivotItems](pivotfield-pivotitems-method-excel.md)|
|[AutoGroup](pivotfield-autogroup-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AllItemsVisible](pivotfield-allitemsvisible-property-excel.md)|
|[Application](pivotfield-application-property-excel.md)|
|[AutoShowCount](pivotfield-autoshowcount-property-excel.md)|
|[AutoShowField](pivotfield-autoshowfield-property-excel.md)|
|[AutoShowRange](pivotfield-autoshowrange-property-excel.md)|
|[AutoShowType](pivotfield-autoshowtype-property-excel.md)|
|[AutoSortCustomSubtotal](pivotfield-autosortcustomsubtotal-property-excel.md)|
|[AutoSortField](pivotfield-autosortfield-property-excel.md)|
|[AutoSortOrder](pivotfield-autosortorder-property-excel.md)|
|[AutoSortPivotLine](pivotfield-autosortpivotline-property-excel.md)|
|[BaseField](pivotfield-basefield-property-excel.md)|
|[BaseItem](pivotfield-baseitem-property-excel.md)|
|[Calculation](pivotfield-calculation-property-excel.md)|
|[Caption](pivotfield-caption-property-excel.md)|
|[ChildField](pivotfield-childfield-property-excel.md)|
|[ChildItems](pivotfield-childitems-property-excel.md)|
|[Creator](pivotfield-creator-property-excel.md)|
|[CubeField](pivotfield-cubefield-property-excel.md)|
|[CurrentPage](pivotfield-currentpage-property-excel.md)|
|[CurrentPageList](pivotfield-currentpagelist-property-excel.md)|
|[CurrentPageName](pivotfield-currentpagename-property-excel.md)|
|[DatabaseSort](pivotfield-databasesort-property-excel.md)|
|[DataRange](pivotfield-datarange-property-excel.md)|
|[DataType](pivotfield-datatype-property-excel.md)|
|[DisplayAsCaption](pivotfield-displayascaption-property-excel.md)|
|[DisplayAsTooltip](pivotfield-displayastooltip-property-excel.md)|
|[DisplayInReport](pivotfield-displayinreport-property-excel.md)|
|[DragToColumn](pivotfield-dragtocolumn-property-excel.md)|
|[DragToData](pivotfield-dragtodata-property-excel.md)|
|[DragToHide](pivotfield-dragtohide-property-excel.md)|
|[DragToPage](pivotfield-dragtopage-property-excel.md)|
|[DragToRow](pivotfield-dragtorow-property-excel.md)|
|[DrilledDown](pivotfield-drilleddown-property-excel.md)|
|[EnableItemSelection](pivotfield-enableitemselection-property-excel.md)|
|[EnableMultiplePageItems](pivotfield-enablemultiplepageitems-property-excel.md)|
|[Formula](pivotfield-formula-property-excel.md)|
|[Function](pivotfield-function-property-excel.md)|
|[GroupLevel](pivotfield-grouplevel-property-excel.md)|
|[Hidden](pivotfield-hidden-property-excel.md)|
|[HiddenItems](pivotfield-hiddenitems-property-excel.md)|
|[HiddenItemsList](pivotfield-hiddenitemslist-property-excel.md)|
|[IncludeNewItemsInFilter](pivotfield-includenewitemsinfilter-property-excel.md)|
|[IsCalculated](pivotfield-iscalculated-property-excel.md)|
|[IsMemberProperty](pivotfield-ismemberproperty-property-excel.md)|
|[LabelRange](pivotfield-labelrange-property-excel.md)|
|[LayoutBlankLine](pivotfield-layoutblankline-property-excel.md)|
|[LayoutCompactRow](pivotfield-layoutcompactrow-property-excel.md)|
|[LayoutForm](pivotfield-layoutform-property-excel.md)|
|[LayoutPageBreak](pivotfield-layoutpagebreak-property-excel.md)|
|[LayoutSubtotalLocation](pivotfield-layoutsubtotallocation-property-excel.md)|
|[MemberPropertyCaption](pivotfield-memberpropertycaption-property-excel.md)|
|[MemoryUsed](pivotfield-memoryused-property-excel.md)|
|[Name](pivotfield-name-property-excel.md)|
|[NumberFormat](pivotfield-numberformat-property-excel.md)|
|[Orientation](pivotfield-orientation-property-excel.md)|
|[Parent](pivotfield-parent-property-excel.md)|
|[ParentField](pivotfield-parentfield-property-excel.md)|
|[ParentItems](pivotfield-parentitems-property-excel.md)|
|[PivotFilters](pivotfield-pivotfilters-property-excel.md)|
|[Position](pivotfield-position-property-excel.md)|
|[PropertyOrder](pivotfield-propertyorder-property-excel.md)|
|[PropertyParentField](pivotfield-propertyparentfield-property-excel.md)|
|[RepeatLabels](pivotfield-repeatlabels-property-excel.md)|
|[ServerBased](pivotfield-serverbased-property-excel.md)|
|[ShowAllItems](pivotfield-showallitems-property-excel.md)|
|[ShowDetail](pivotfield-showdetail-property-excel.md)|
|[ShowingInAxis](pivotfield-showinginaxis-property-excel.md)|
|[SourceCaption](pivotfield-sourcecaption-property-excel.md)|
|[SourceName](pivotfield-sourcename-property-excel.md)|
|[StandardFormula](pivotfield-standardformula-property-excel.md)|
|[SubtotalName](pivotfield-subtotalname-property-excel.md)|
|[Subtotals](pivotfield-subtotals-property-excel.md)|
|[TotalLevels](pivotfield-totallevels-property-excel.md)|
|[UseMemberPropertyAsCaption](pivotfield-usememberpropertyascaption-property-excel.md)|
|[Value](pivotfield-value-property-excel.md)|
|[VisibleItems](pivotfield-visibleitems-property-excel.md)|
|[VisibleItemsList](pivotfield-visibleitemslist-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
