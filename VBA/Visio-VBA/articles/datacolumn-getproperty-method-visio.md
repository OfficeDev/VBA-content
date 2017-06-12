---
title: DataColumn.GetProperty Method (Visio)
keywords: vis_sdr.chm16760400
f1_keywords:
- vis_sdr.chm16760400
ms.prod: visio
api_name:
- Visio.DataColumn.GetProperty
ms.assetid: 8fa134e8-320d-546b-1de1-e19607a60c49
ms.date: 06/08/2017
---


# DataColumn.GetProperty Method (Visio)

Gets the value of the specified data-column property.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **GetProperty**( **_Property_** )

 _expression_ An expression that returns a **DataColumn** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Property_|Required| **VisDataColumnProperties**|The data column property to get. See Remarks for possible values.|

### Return Value

Variant


## Remarks

When you link shapes in a Microsoft Visio drawing to data in a data recordset, Visio maps columns in the data recordset to rows in the Shape Data section of the ShapeSheet spreadsheet, each of which corresponds to a shape-data item. 


 **Note**  In some previous versions of Visio, shape data were called custom properties.

Data-column properties map data columns to certain cells in the Shape Data section of the ShapeSheet. For example, by passing the  **GetProperty** method the DisplayName property, which is represented by the enumerated value **visDataColumnPropertyDisplayName** , you can get the value of the Label cell in the Shape Data section of the ShapeSheet for a particular shape data item. In addition, that property sets the label of the shape data item in the **Shape Data** dialog box, as well as the name of the data column that is displayed in the **External Data** window in the Visio user interface.

Possible values for the Property parameter are declared in  **VisDataColumnProperties** , and are shown in the following table.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| ** visDataColumnPropertyCalendar**|3|Calendar of the data-column property.|
| **visDataColumnPropertyCurrency**|5|Currency of the data-column property.|
| **visDataColumnPropertyDisplayName**|6|Display name of the data-column property in the UI.|
| **visDataColumnPropertyHyperlink**|8|Whether the data-column value becomes a hyperlink in the Visio UI when it is linked to a shape.|
| **visDataColumnPropertyLangID**|2|Language ID of the data-column property.|
| **visDataColumnPropertyType**|1|Type of the data-column property.|
| **visDataColumnPropertyUnits**|4|Units of the data-column property.|
| **visDataColumnPropertyVisible**|7|Whether the data-column property is visible in the UI.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetProperty** method to get the value of the Label cell in the Shape Data section for the first column in the data recordset passed to the method and display it in the **Immediate** window. Then it uses the **SetProperty** method to set the value and displays the new value. Changing this value changes the label of the shape data item in the **Shape Data** dialog box for all shapes linked to rows in the data recordset.

To get and set the Label cell value, the macro passes the  **visDataColumnPropertyDisplayName** value from the **VisDataColumnProperties** enumeration to the **DataColumn.GetProperty** and **DataColumn.SetProperty** methods.

Before running this macro, create at least one data recordset in your VBA project to pass to the macro.




```vb
 
Public Sub GetProperty_Example(vsoDataRecordset As Visio.DataRecordset) 
    Dim strPropertyName As String 
    Dim strNewName As String 
    Dim vsoDataColumn As Visio.DataColumn 
 
    strNewName = "New Property Name" 
    Set vsoDataColumn = vsoDataRecordset.DataColumns(1) 
 
    strPropertyName = vsoDataColumn.GetProperty(visDataColumnPropertyDisplayName) 
    Debug.Print strPropertyName 
 
    vsoDataColumn.SetProperty visDataColumnPropertyDisplayName, strNewName 
    strPropertyName = vsoDataColumn.GetProperty(visDataColumnPropertyDisplayName) 
    Debug.Print strPropertyName 
End Sub
```


