---
title: DataColumn.SetProperty Method (Visio)
keywords: vis_sdr.chm16760405
f1_keywords:
- vis_sdr.chm16760405
ms.prod: visio
api_name:
- Visio.DataColumn.SetProperty
ms.assetid: 5851daa0-e2e0-7073-7e26-f0fc73586b9b
ms.date: 06/08/2017
---


# DataColumn.SetProperty Method (Visio)

Sets the value of the specified data-column property.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **SetProperty**( **_Property_** , **_Value_** )

 _expression_ An expression that returns a **DataColumn** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Property_|Required| **VisDataColumnProperties**|The data-column property whose value you want set. See Remarks for possible values.|
| _Value_|Required| **Variant**|The value to assign the data-column property. See Remarks for possible values.|

### Return Value

Nothing


## Remarks

When you link shapes in a Microsoft Visio drawing to data in a data recordset, Visio maps columns in the data recordset to rows in the Shape Data section of the ShapeSheet spreadsheet, each of which corresponds to a shape-data item. 


 **Note**  In some previous versions of Visio, shape data were called custom properties.

Data-column properties map data columns to certain cells in the Shape Data section of the ShapeSheet. For example, by passing the  **SetProperty** method a new value for the **DisplayName** property, which is represented by the enumerated value **visDataColumnPropertyDisplayName** , you set the value of the Label cell in the Shape Data section of the ShapeSheet for a particular shape data item. In addition, setting that property sets the label of the shape data item in the **Shape Data** dialog box, as well as the name of the data column that is displayed in the **External Data** window in the Visio user interface. These settings correspond to those that you can set in the **Column Settings** dialog box in the Visio user interface (right-click in the **External Data** window and then click **Column Settings**), as well as those that you can make in the  **Types and Units** dialog box for each column (click **Data Types** in the **Column Settings** dialog box.)

Possible values for the Property parameter are declared in  **VisDataColumnProperties** , and are shown in the following table.



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
Possible values for the Value parameter depend on the Property parameter value. The following table shows valid data-column property values for each data-column property, depending on the data-column data type.



|****|****|**Data Column Data Type**|****|****|****|****|****|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|||Number (VisCellVals.visPropTypeNumber)|Date (VisCellVals.visPropTypeDate)|Currency (VisCellVals.visPropTypeCurrency)|Duration (VisCellVals.visPropTypeDuration)|String (VisCellVals.visPropTypeString)|Boolean (VisCellVals.visPropTypeBoolean)|
|Data Column Property |Type| **visPropTypeNumber**| **visPropTypeDate**| **visPropTypeCurrency**| **visPropTypeDuration**| **visPropTypeString**| **visPropTypeBoolean**|
||Visible| **Boolean**| **Boolean**| **Boolean**| **Boolean**| **Boolean**| **Boolean**|
||DisplayName| **String**| **String**| **String**| **String**| **String**| **String**|
||LangID||Valid LCID number|||||
||Currency|||Valid 3-letter currency-constant string as used in the CY function in the Visio ShapeSheet spreadsheet.||||
||Calendar||One of the members of  **VisCellVals** , depending on the LangID value. (See table below.)|||||
||Units|One of the following members of  **VisUnitsCodes** :
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>visAcre</b></p></li><li><p><b>visAngleUnits</b></p></li><li><p><b>visCentimeters</b></p></li><li><p><b>visCiceros</b></p></li><li><p><b>visCicerosAndDidots</b></p></li><li><p><b>visDegreeMinSec</b></p></li><li><p><b>visDegrees</b></p></li><li><p><b>visDrawingUnits</b></p></li><li><p><b>visFeet</b></p></li><li><p><b>visFeetAndInches</b></p></li><li><p><b>visHectare</b></p></li><li><p><b>visDidots</b></p></li><li><p><b>visInches</b></p></li><li><p><b>visInchFrac</b></p></li><li><p><b>visKilometers</b></p></li><li><p><b>visMeters</b></p></li><li><p><b>visMileFrac</b></p></li><li><p><b>visMiles</b></p></li><li><p><b>visMillimeters</b></p></li><li><p><b>visMin</b></p></li><li><p><b>visNautMiles</b></p></li><li><p><b>visPageUnits</b></p></li><li><p><b>visPicas</b></p></li><li><p><b>visPicasAndPoints</b></p></li><li><p><b>visPoints</b></p></li><li><p><b>visRadians</b></p></li><li><p><b>visSec</b></p></li><li><p><b>visYards</b></p></li><li><p><b>visNumber</b>  (special behavior ? this constant makes the value unitless)  
</p></li></ul>OR Descriptive string?a string used for units, such as  _cm_ or _sq cm_ . This string will be validated so that it is one of the supported Visio units. Passing invalid strings causes the method to fail.|||One of the following members of  **VisUnitsCodes** :
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>visDurationUnits</b></p></li><li><p><b>visElapsedDay</b></p></li><li><p><b>visElapsedHour</b></p></li><li><p><b>visElapsedMin</b></p></li><li><p><b>visElapsedSec</b></p></li><li><p><b>visElapsedWeek</b></p></li></ul>OR Descriptive string?a string used for units such as  _ew_ . This string will be validated so that it is one of the supported Visio units. Passing an invalid string will cause this method to fail.|||
||HyperLink||||| **Boolean**||
The LangID and Calendar properties are bound by the validation rules shown in the following table. Languages not shown use the Western calendar only.



|**Calendar**|**Hirji**|**Western**|**French Transliterated**|**English Transliterated**|**Hebrew Lunar**|**Saka Era**|**Japanese Emperor Era**|**Korean Danki**|**Thai Buddhist**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Language||||||||||
|All Arabic|x|x|x|x||||||
|Bengali(Bangladesh)|x|x||||||||
|Divehi|x|x||||||||
|All English|x|x|||x|x||||
|Persian|x|x||||||||
|Hebrew|x||||x|||||
|Hindi|x|||||x||||
|Japanese||x|||||x|||
|Korean||x||||||x||
|Kashmiri (Arabic)|x|x||||||||
|Punjabi (Pakistan)|x|x||||||||
|Pashto|x|x||||||||
|Sindhi|x|x||||||||
|Thai||||||||||
|Urdu|x|x||||||||
|Tamzight|x|x||||||||

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetProperty** method to get the value of the Label cell in the Shape Data section for the first column in the data recordset passed to the method and display it in the **Immediate** window. Then it uses the **SetProperty** method to set the value and displays the new value. Changing this value changes the label of the shape data item in the **Shape Data** dialog box for all shapes linked to rows in the data recordset.

To get and set the Label cell value, the macro passes the  **visDataColumnPropertyDisplayName** value from the **VisDataColumnProperties** enumeration to the **DataColumn.GetProperty** and **DataColumn.SetProperty** methods.

Before running this macro, create at least one data recordset in your VBA project to pass to the macro.




```vb
 
Public Sub SetProperty_Example(vsoDataRecordset As Visio.DataRecordset) 
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


