---
title: TreeviewControl.Drilled Property (Excel)
keywords: vbaxl10.chm666074
f1_keywords:
- vbaxl10.chm666074
ms.prod: excel
api_name:
- Excel.TreeviewControl.Drilled
ms.assetid: 5e4f1b52-a02f-655b-f3c8-b5e7aa54d928
ms.date: 06/08/2017
---


# TreeviewControl.Drilled Property (Excel)

Sets the "drilled" (expanded, or visible) status of the cube field members in the hierarchical member-selection control of a cube field. This property is used primarily for macro recording and isn't intended for any other use. Read/write.


## Syntax

 _expression_ . **Drilled**

 _expression_ A variable that represents a **TreeviewControl** object.


## Remarks

The  **Drilled** property accepts an array. Each element of the array corresponds to a level of the cube field that has been expanded. The maximum number of elements is the number of levels in the cube field. Each element of the array is an array of type **String** , containing unique member names that are visible (expanded) at the corresponding level of the control. See the **[TreeviewControl](treeviewcontrol-object-excel.md)** object's **[Hidden](treeviewcontrol-hidden-property-excel.md)** property to determine when members are explicitly hidden in an expanded view.




 **Note**  This property does not return an array that represents the drilled status of the cube field members in the hierarchical member-selection control of a cube field.


## Example

This example expands the second-level members of the first cube field in the first PivotTable report on the active worksheet.


```vb
ActiveSheet.PivotTables("PivotTable1").CubeFields(1) _ 
 .TreeviewControl.Drilled = _ 
 Array(Array("", "", "", "", "", "", "", "", _ 
 "", "", "", ""), _ 
 Array("[state].[states].[AB]", _ 
 "[state].[states].[CA]", _ 
 "[state].[states].[IN]", _ 
 "[state].[states].[KS]", _ 
 "[state].[states].[KY]", _ 
 "[state].[states].[MD]", _ 
 "[state].[states].[MI]", _ 
 "[state].[states].[OH]", _ 
 "[state].[states].[OR]", _ 
 "[state].[states].[TN]", _ 
 "[state].[states].[UT]", _ 
 "[state].[states].[WA]"))
```


## See also


#### Concepts


[TreeviewControl Object](treeviewcontrol-object-excel.md)

