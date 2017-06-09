---
title: CubeField Object (Excel)
keywords: vbaxl10.chm667072
f1_keywords:
- vbaxl10.chm667072
ms.prod: excel
api_name:
- Excel.CubeField
ms.assetid: 6db16910-6c27-651a-c388-e54e27fe4519
ms.date: 06/08/2017
---


# CubeField Object (Excel)

Represents a hierarchy or measure field from an OLAP cube. In a PivotTable report, the  **CubeField** object is a member of the **[CubeFields](cubefields-object-excel.md)** collection.


## Example

Use the  **[CubeField](pivotfield-cubefield-property-excel.md)** property to return the **CubeField** object. This example creates a list of the cube field names for all the hierarchy fields in the first OLAP-based PivotTable report on Sheet1.


```
Set objNewSheet = Worksheets.Add 
objNewSheet.Activate 
intRow = 1 
For Each objPF in _ 
 Worksheets("Sheet1").PivotTables(1).PivotFields 
 If objPF.CubeField.CubeFieldType = xlHierarchy Then 
 objNewSheet.Cells(intRow, 1).Value = objPF.Name 
 intRow = intRow + 1 
 End If 
Next objPF
```

Use  **CubeFields** ( _index_ ), where _index_ is the cube field's index number, to return a single **CubeField** object. The following example determines the name of the second cube field in the first PivotTable report on the active worksheet.




```
strAlphaName = _ 
 ActiveSheet.PivotTables(1).CubeFields(2).Name
```


## Methods



|**Name**|
|:-----|
|[AddMemberPropertyField](cubefield-addmemberpropertyfield-method-excel.md)|
|[AutoGroup](cubefield-autogroup-method-excel.md)|
|[ClearManualFilter](cubefield-clearmanualfilter-method-excel.md)|
|[CreatePivotFields](cubefield-createpivotfields-method-excel.md)|
|[Delete](cubefield-delete-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AllItemsVisible](cubefield-allitemsvisible-property-excel.md)|
|[Application](cubefield-application-property-excel.md)|
|[Caption](cubefield-caption-property-excel.md)|
|[Creator](cubefield-creator-property-excel.md)|
|[CubeFieldSubType](cubefield-cubefieldsubtype-property-excel.md)|
|[CubeFieldType](cubefield-cubefieldtype-property-excel.md)|
|[CurrentPageName](cubefield-currentpagename-property-excel.md)|
|[DragToColumn](cubefield-dragtocolumn-property-excel.md)|
|[DragToData](cubefield-dragtodata-property-excel.md)|
|[DragToHide](cubefield-dragtohide-property-excel.md)|
|[DragToPage](cubefield-dragtopage-property-excel.md)|
|[DragToRow](cubefield-dragtorow-property-excel.md)|
|[EnableMultiplePageItems](cubefield-enablemultiplepageitems-property-excel.md)|
|[FlattenHierarchies](cubefield-flattenhierarchies-property-excel.md)|
|[HasMemberProperties](cubefield-hasmemberproperties-property-excel.md)|
|[HierarchizeDistinct](cubefield-hierarchizedistinct-property-excel.md)|
|[IncludeNewItemsInFilter](cubefield-includenewitemsinfilter-property-excel.md)|
|[IsDate](cubefield-isdate-property-excel.md)|
|[LayoutForm](cubefield-layoutform-property-excel.md)|
|[LayoutSubtotalLocation](cubefield-layoutsubtotallocation-property-excel.md)|
|[Name](cubefield-name-property-excel.md)|
|[Orientation](cubefield-orientation-property-excel.md)|
|[Parent](cubefield-parent-property-excel.md)|
|[PivotFields](cubefield-pivotfields-property-excel.md)|
|[Position](cubefield-position-property-excel.md)|
|[ShowInFieldList](cubefield-showinfieldlist-property-excel.md)|
|[TreeviewControl](cubefield-treeviewcontrol-property-excel.md)|
|[Value](cubefield-value-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
