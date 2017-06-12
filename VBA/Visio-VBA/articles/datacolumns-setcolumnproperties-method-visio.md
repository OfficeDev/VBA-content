---
title: DataColumns.SetColumnProperties Method (Visio)
keywords: vis_sdr.chm16660390
f1_keywords:
- vis_sdr.chm16660390
ms.prod: visio
api_name:
- Visio.DataColumns.SetColumnProperties
ms.assetid: 453de04e-3def-11d1-67a4-127da4459564
ms.date: 06/08/2017
---


# DataColumns.SetColumnProperties Method (Visio)

Sets one or more data-column properties for one or more data columns.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **SetColumnProperties**( **_ColumnNames()_** , **_Properties()_** , **_Values()_** )

 _expression_ An expression that returns a **DataColumns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ColumnNames()_|Required| **String**|An array of strings that represent data column names.|
| _Properties()_|Required| **Long**|An array of data-column properties, as  **VisDataColumnProperties** . See Remarks for possible values.|
| _Values()_|Required| **Variant**|An array of values to be assigned to the properties. See Remarks for possible values.|

### Return Value

Nothing


## Remarks

The  **SetColumnProperties** method is a more efficient way to set properties for multiple data columns simultaneously than is setting properties one column at a time. Depending on the items you place in each of the three parameter arrays, you can change multiple properties of the same data column or one or more properties of different data columns. For each change you want to make, pass related column-name/property/value triplets at corresponding positions of all three arrays. Note that the size of all three arrays that you pass to the method must be the same, or the method will return an error.

For the ColumnNames() parameter, pass an array of the names of the data columns whose properties you want to change. If you want to change multiple properties of the same column, you can either place the same name in multiple array positions, or you can place the column name in one array position and place empty strings in the succeeding positions that correspond to the array positions of the properties you want to change. 

Possible values for items in the Properties() parameter array are declared in  **VisDataColumnProperties** , and are shown in the following table.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| ** visDataColumnPropertyCalendar**|3|Calendar of the data-column property.|
| **visDataColumnPropertyCurrency**|5|Currency of the data-column property.|
| **visDataColumnPropertyDisplayName**|6|Display name of the data-column property in the UI.|
| **visDataColumnPropertyHyperlink**|8|Whether the data-column value becomes a hyperlink in the Visio UI when it is linked to a shape.|
| **visDataColumnPropertyLangID**|2|Language ID of the data-column property.|
| **visDataColumnPropertyType**|1|Data type of the data-column property.|
| **visDataColumnPropertyUnits**|4|Units of the data-column property.|
| **visDataColumnPropertyVisible**|7|Whether the data-column property is visible in the UI.|
Possible values for items in the Values() parameter array depend on the corresponding Property() array parameter values. The table in the  **[DataColumn.SetProperty](datacolumn-setproperty-method-visio.md)** topic shows valid data-column property values for each data-column property, depending on the data-column data type.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to set the value of a single property for two different columns in the most recently added recordset in the  **DataRecordsets** collection of the active document. The macro assumes that the most recently added recordset is one based on data in the ORGDATA.xls spreadsheet that is shipped with Visio. Sample code for adding that data recordset is shown in the **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method topic. However, you can use this code with any data recordset that has at least two columns.

The macro changes the display name of the first column to "Dept." and sets the  **Hyperlink** property of the second column to **True** . Thereafter (if you used ORGDATA as your data source), the e-mail address of shapes linked to data in the data recordset will act as a hyperlink.

Note that changing the display name of a data column changes only its  **[DisplayName](datacolumn-displayname-property-visio.md)** property, and does not change the column's programmatic name, which is specified by its **[Name](datacolumn-name-property-visio.md)** property.




```vb
 
Public Sub SetColumnProperties_Example() 
 
    Dim vsoDataRecordset As Visio.DataRecordset 
    Dim intCount As Integer 
     
    intCount = Visio.ActiveDocument.DataRecordsets.Count 
    Set vsoDataRecordset = Visio.ActiveDocument.DataRecordsets(intCount) 
     
    Dim astrColumnNames(1) As String 
    Dim alngProperties(1) As Long 
    Dim avarValues(1) As Variant 
     
    astrColumnNames(0) = vsoDataRecordset.DataColumns(1).DisplayName 
    astrColumnNames(1) = vsoDataRecordset.DataColumns(2).DisplayName 
        
    alngProperties(0) = visDataColumnPropertyDisplayName 
    alngProperties(1) = visDataColumnPropertyHyperlink 
        
    avarValues(0) = "Dept." 
    avarValues(1) = True 
         
    vsoDataRecordset.DataColumns.SetColumnProperties astrColumnNames, alngProperties, avarValues 
 
End Sub
```


