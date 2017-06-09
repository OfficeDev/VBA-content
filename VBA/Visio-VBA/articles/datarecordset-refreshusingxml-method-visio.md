---
title: DataRecordset.RefreshUsingXML Method (Visio)
keywords: vis_sdr.chm16460325
f1_keywords:
- vis_sdr.chm16460325
ms.prod: visio
api_name:
- Visio.DataRecordset.RefreshUsingXML
ms.assetid: 345935ab-b269-61dd-9ebe-e1f87b89bb11
ms.date: 06/08/2017
---


# DataRecordset.RefreshUsingXML Method (Visio)

Updates linked shapes with data contained in the string that conforms to the ADO classic XML schema passed to the method as a parameter.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **RefreshUsingXML**( **_NewDataAsXML_** )

 _expression_ An expression that returns a **DataRecordset** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewDataAsXML_|Required| **String**|An XML string that contains new data to refresh the data recordset and that conforms to the classic ADO schema.|

### Return Value

Nothing


## Remarks

For the XMLString parameter, pass an XML string that conforms to the ADO classic XML schema and that describes the data you want to import. A sample XML string is shown in the example later in this topic. 

The data in the XML string you pass to the  **RefreshUsingXML** method should be structured in a manner similar to that of the data in the data recordset you want to update. At a minimum, the primary key columns should be the same in both sets of data. The _primary key_ identifies the name of the data column or columns that contain unique identifiers for each row. The value in the primary key column for each row uniquely identifies that row in the data recordset.

When you create a data recordset, Microsoft Visio assigns row IDs to all the rows in the recordset based on the existing order of the rows in the data source. 

If the XML string you pass to the  **RefreshUsingXML** method contains a column consisting of Visio row IDs (as it would, for example, if you exported it from Visio by getting the **[DataAsXML ](datarecordset-dataasxml-property-visio.md)** property value of the data recordset), the **RefreshUsingXML** method attempts to validate the row IDs in the string. If the method finds the row IDs to be valid, it reuses them in the updated data recordset. If it finds them to be invalid, it returns an error.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **RefreshUsingXML** method to update an existing data recordset with data contained in an ADO classic XML string.

 A sample XML string is shown here. Before running this macro, open a new Visio drawing and run the macro in the **[DataRecordsets.AddFromXML](datarecordsets-addfromxml-method-visio.md)** method topic.

When you pass it to the  **RefreshUsingXML** method, this string will update the data recordset that the **AddFromXML** method created, changing the city names.




```
<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' 
xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' 
xmlns:rs='urn:schemas-microsoft-com:rowset' 
xmlns:z='#RowsetSchema'> 
<s:Schema id='RowsetSchema'> 
<s:ElementType name='row' content='eltOnly' rs:updatable='true'> 
<s:AttributeType name='c1' rs:name='Cities' 
rs:number='2' rs:nullable='true' rs:maydefer='true' rs:write='true'> 
<s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/> 
</s:AttributeType> 
<s:extends type='rs:rowbase'/> 
</s:ElementType> 
</s:Schema> 
<rs:data> 
<z:row c1='New York' /> 
<z:row c1='London' /> 
</rs:data> 
</xml>
```

In the following sample code, we pass the  **RefreshUsingXML** method the name of an XML string containing the updated data.




```vb
Public Sub RefreshUsingXML_Example() 
 
    Dim strXML As String 
    Dim intCount As Integer 
    Dim vsoDataRecordset As Visio.DataRecordset 
 
intCount = ThisDocument.DataRecordsets.Count 
 
    strXML = "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'" + Chr(10) _ 
    &; "xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'" + Chr(10) _ 
    &; "xmlns:rs='urn:schemas-microsoft-com:rowset'" + Chr(10) _ 
    &; "xmlns:z='#RowsetSchema'>" + Chr(10) _ 
    &; "<s:Schema id='RowsetSchema'>" + Chr(10) _ 
    &; "<s:ElementType name='row' content='eltOnly' rs:updatable='true'>" + Chr(10) _ 
    &; "<s:AttributeType name='c1' rs:name='Cities'" + Chr(10) _ 
    &; "rs:number='2' rs:nullable='true' rs:maydefer='true' rs:write='true'>" + Chr(10) _ 
    &; "<s:datatype dt:type='string' dt:maxLength='255' rs:precision='0'/>" + Chr(10) _ 
    &; "</s:AttributeType>" + Chr(10) _ 
    &; "<s:extends type='rs:rowbase'/>" + Chr(10) _ 
    &; "</s:ElementType>" + Chr(10) _ 
    &; "</s:Schema>" + Chr(10) _ 
    &; "<rs:data>" + Chr(10) _ 
    &; "<z:row c1='New York'/>" + Chr(10) _ 
    &; "<z:row c1='London'/>" + Chr(10) _ 
    &; "</rs:data>" + Chr(10) _ 
    &; "</xml>" 
 
    ThisDocument.DataRecordsets(intCount).RefreshUsingXML(strXML) 
 
End Sub
```


