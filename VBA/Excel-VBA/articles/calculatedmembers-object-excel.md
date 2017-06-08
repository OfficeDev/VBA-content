---
title: CalculatedMembers Object (Excel)
keywords: vbaxl10.chm683072
f1_keywords:
- vbaxl10.chm683072
ms.prod: excel
api_name:
- Excel.CalculatedMembers
ms.assetid: 3c664ac6-e2f8-f631-006d-6a16c380641e
ms.date: 06/08/2017
---


# CalculatedMembers Object (Excel)

A collection of all the  **[CalculatedMember](calculatedmembers-object-excel.md)** objects on the specified PivotTable.


## Remarks

 Each **CalculatedMember** object represents a calculated member or calculated measure.

Use the  **[CalculatedMembers](pivottable-calculatedmembers-property-excel.md)** property of the **[PivotTable](pivottable-object-excel.md)** object to return a **CalculatedMembers** collection.

There are three supported types of calculated members:  _Named Sets_ , _Calculated Measures_ , and _Calculated Members_ . Object model support has been available for all three of these types since Excel 2010. User interface support was made available for Named Sets in Excel 2010. In Excel 2013, the OLAP Calculated Members and Calculated Measures feature was created to build a user interface for the calculated members and measures object model.

 **Named Sets** are used exactly the same as in Excel 2010. Named Sets should continue to use the method CalculatedMembers.[CalculatedMembers.Add Method (Excel)](calculatedmembers-add-method-excel.md) and the type[XlCalculatedMemberType Enumeration (Excel)](xlcalculatedmembertype-enumeration-excel.md).

 **Calculated Members** have the following changes for Excel 2013:


- They now use the method called CalculatedMembers.[CalculatedMembers.AddCalculatedMember Method (Excel)](calculatedmembers-addcalculatedmember-method-excel.md).
    
- They support the property [CalculatedMember.ParentHierarchy Property (Excel)](calculatedmember-parenthierarchy-property-excel.md).
    
- They support the property [CalculatedMember.ParentMember Property (Excel)](calculatedmember-parentmember-property-excel.md).
    
- They support the property [CalculatedMember.NumberFormat Property (Excel)](calculatedmember-numberformat-property-excel.md).
    
 **Calculated Measures** have the following changes for Excel 2013:


- They now use the method called CalculatedMembers.[CalculatedMembers.AddCalculatedMember Method (Excel)](calculatedmembers-addcalculatedmember-method-excel.md).
    
- They now use the type [XlCalculatedMemberType Enumeration (Excel)](xlcalculatedmembertype-enumeration-excel.md).
    
- They support the property [CalculatedMember.DisplayFolder Property (Excel)](calculatedmember-displayfolder-property-excel.md).
    
- They support the property [CalculatedMember.NumberFormat Property (Excel)](calculatedmember-numberformat-property-excel.md).
    

## Example

The following example adds a set to a PivotTable, assuming a PivotTable from the FoodMart SQL database exists on the active worksheet.


```vb
Sub UseCalculatedMember() 
 Dim pvtTable As PivotTable 
 Set pvtTable = ActiveSheet.PivotTables(1)
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _ 
 Formula:="'{[Product].[All Products].Children}'", _ 
 Type:=xlCalculatedSet 
 
End Sub
```


 **Note**  For the  **Add** method in the previous example, the **Formula** argument must have a valid MDX syntax statement. The **Name** argument has to be acceptable to the Online Analytical Processing (OLAP) provider and the **Type** argument has to be defined.


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


