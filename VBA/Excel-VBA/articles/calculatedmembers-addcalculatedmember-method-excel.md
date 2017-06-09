---
title: CalculatedMembers.AddCalculatedMember Method (Excel)
keywords: vbaxl10.chm684079
f1_keywords:
- vbaxl10.chm684079
ms.prod: excel
ms.assetid: 61e3fdf5-f7e3-9d7f-4449-1f4408251422
ms.date: 06/08/2017
---


# CalculatedMembers.AddCalculatedMember Method (Excel)

Adds a calculated field or calculated item to a PivotTable.


## Syntax

 _expression_ . **AddCalculatedMember**_(Name,_ _Formula,_ _SolveOrder,_ _Type,_ _DisplayFolder,_ _MeasureGroup,_ _ParentHierarchy,_ _ParentMember,_ _NumberFormat)_

 _expression_ A variable that represents a **CalculatedMembers** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|STRING|The name of the calculated member.|
| _Formula_|Required|VARIANT|The formula of the calculated member.|
| _SolveOrder_|Optional|VARIANT|The solve order for the calculated member.|
| _Type_|Optional|VARIANT|The type of calculated member.|
| _DisplayFolder_|Optional|VARIANT|A folder that exists to display calculated measures.|
| _MeasureGroup_|Optional|VARIANT|The group to which the calculated belongs.|
| _ParentHierarchy_|Optional|VARIANT|The parent path of the  **ParentMember** .|
| _ParentMember_|Optional|VARIANT|The parent of the calculated member.|
| _NumberFormat_|Optional|VARIANT|The format of numbers used for calculated members.|

### Return value

 **CALCULATEDMEMBER**


### Remarks

 The **Formula** argument must have a valid MDX (multidimensional expression) syntax statement. The **Name** argument has to be acceptable to the Online Analytical Processing (OLAP) provider.


-  **DisplayFolder** : Display folders are only valid for calculated measures. They are not valid for calculated members. The string can have semicolons **;** in it. Semicolons designate multiple display folders. For example, if you use the string my **folder1;myfolder2** then the calculated measure will show in two display folders one named **myfolder1** the other named ?myfolder2?. The string can have backslashes **\**. This designates a hierarchical path for the display folder. For example, if you use the string  **welcome\to\seattle** there will be a display folder called **welcome** that contains a display folder called **to** which contains a display folder called **seattle**. Display folders are virtual folders.? They do not really exist in the same sense that we think of system folders. They only exist for purposes of displaying the calculated measures.
    
-  **NumberFormat** : The number formats can only be set by macros. There is no user interface for setting them. This is the only property that cannot be set via the user interface. The type is always **xlNumberFormatTypeDefault** when a calculated member is created via the user interface. The number formats are only valid for calculated members. They are not valid for calculated measures.
    
-  **ParentHierarchy** : The parent hierarchy can be any valid MDX hierarchy. Parent hierarchies are only valid for calculated members. They are not valid for calculated measures. If a parent member is chosen that is in a different parent hierarchy, the parent hierarchy will be automatically changed to match the parent hierarchy of the parent member. For example, assume the following macro for a calculated member:
    
```vb
     OLEDBConnection.CalculatedMembers.AddCalculatedMember Name:="[UK+US]", _
     Formula:= _
    "[Customer].[Customer Geography].[Country].&;[United Kingdom] + [Customer].[Customer Geography].[Country].&;[United States] " _
     , Type:=xlCalculatedMember, SolveOrder:=0, ParentHierarchy:= _
     "[Account].[Accounts]", ParentMember:= _
    "[Customer].[Customer Geography].[Australia]", NumberFormat:= _
     xlNumberFormatTypePercent
```


    In this case you have specified that the parent member is from the  **[Customer].[Customer Geography]** hierarchy yet you have given the parent hierarchy as **[Account].[Accounts]**. When the member is created, it will use the parent hierarchy of the parent member which is  **[Customer].[Customer Geography]** and when you look in the **Manage Calculations** dialog in the UI, it will show **[Customer].[Customer Geography]** as the parent hierarchy rather than the one specified in the macro, i.e. **[Account].[Accounts]**.
    

### Example

The following code sample adds a  _calculated measure_ to a PivotTable.


 **Note**  In both of these samples, the PivotTable must be refreshed after creating the calculation in order to view it in the user interface.


```vb
Sub AddCalculatedMeasure()

Dim pvt As PivotTable
Dim strName As String
Dim strFormula As String
Dim strDisplayFolder As String
Dim strMeasureGroup As String

Set pvt = Sheet1.PivotTables("PivotTable1")
strName = "[Measures].[Internet Sales Amount 25 %]"
strFormula = "[Measures].[Internet Sales Amount]*1.25"
strDisplayFolder = "My Folder\Percent Calculations" 
strMeasureGroup = "Internet Sales"

pvt.CalculatedMembers. AddCalculatedMember Name:=strName, Formula:=strFormula, Type:=xlCalculatedMeasure, DisplayFolder:=strDisplayFolder, MeasureGroup:=strMeasureGroup, NumberFormat:=xlNumberFormatTypePercent

End Sub
```

The following code sample adds a  _calculated member_ to a PivotTable.




```vb
Sub AddCalculatedMember()

Dim pvt As PivotTable
Dim strName As String
Dim strFormula As String
Dim strParentHierarchy As String
Dim strParentMember As String

Set pvt = Sheet1.PivotTables("PivotTable1")
strName = "[Customer].[Customer Geography].[All Customers].[North America]"
strFormula = "[Customer].[Customer Geography].[Country].&;[United States] + [Customer].[Customer Geography].[Country].&;[Canada]"
strParentHierarchy = "[Customer].[Customer Geography]"
strParentMember = "[Customer].[Customer Geography].[All Customers]"

pvt.CalculatedMembers. AddCalculatedMember Name:=strName, Formula:=strFormula, Type:=xlCalculatedMember, ParentHierarchy:=strParentHierarchy, ParentMember:=strParentMember, SolveOrder:=0, NumberFormat:=xlNumberFormatTypeDefault

End Sub
```


## See also


#### Concepts


[CalculatedMembers Collection](calculatedmembers-object-excel.md)

