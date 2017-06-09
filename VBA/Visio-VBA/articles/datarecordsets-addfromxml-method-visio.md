---
title: DataRecordsets.AddFromXML Method (Visio)
keywords: vis_sdr.chm16360270
f1_keywords:
- vis_sdr.chm16360270
ms.prod: visio
api_name:
- Visio.DataRecordsets.AddFromXML
ms.assetid: b75d7ecc-98d2-ae9b-608f-a9ec2b736ea6
ms.date: 06/08/2017
---


# DataRecordsets.AddFromXML Method (Visio)

Adds a  **[DataRecordset](datarecordset-object-visio.md)** object to the **[DataRecordsets](datarecordsets-object-visio.md)** collection, and fills the resulting data recordset with data supplied in the form of an XML string.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **AddFromXML**( **_XMLString_** , **_AddOptions_** , **_Name_** )

 _expression_ An expression that returns a **DataRecordsets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XMLString_|Required| **String**|An XML string that conforms to the Microsoft ActiveX? Data Objects (ADO) classic XML schema and that describes the data you want to import.|
| _AddOptions_|Required| **Long**|Options that determine properties of the data recordset to add. A combination of one or more enumerated value from  **[VisDataRecordsetAddOptions](visdatarecordsetaddoptions-enumeration-visio.md)** . For more information, see Remarks.|
| _Name_|Optional| **String**|Assigns a display name to the  **DataRecordset** object being added.|

### Return Value

DataRecordset


## Remarks

For the XMLString parameter, pass an XML string that conforms to the ADO classic XML schema and that describes the data you want to import. A simple XML string is shown in the example later in this topic.

The AddOptions parameter can be a combination of one or more of the following values from the  **VisDataRecordsetAddOptions** enumeration, which is declared in the Microsoft Visio type library. The default is zero (0), which specifies that none of the options be set.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDataRecordsetNoExternalDataUI**|1|Prevents data in the new data recordset from being displayed in the  **External Data** window.|
| **visDataRecordsetNoAdvConfig**|4|Prevents the data recordset from being displayed in the  **Configure Refresh** dialog box.|
| **visDataRecordsetDontCopyLinks**|16|Adds a data recordset, but shape-data links are not cut or copied.|
 Once you assign these values, you cannot change them for the life of the **DataRecordset** object.

The Name argument is an optional string that lets you assign the data recordset a display name. If you specify that the  **External Data** window display in the Visio UI, the name you pass for this argument appears on the tab of the **External Data** window that corresponds to the data recordset added.

In contrast with data recordsets created by using the  **Add** or **AddFromConnectionFile** methods, data recordsets created by using the **AddFromXML** method are not associated with a **[DataConnection](dataconnection-object-visio.md)** object.

What's more, Visio never refreshes a data recordset you created by using the  **AddFromXML** method automatically, regardless of the setting of the **[DataRecordset.RefreshInterval](datarecordset-refreshinterval-property-visio.md)** property. To refresh the data in such a data recordset, you must call the **[DataRecordset.RefreshUsingXML](datarecordset-refreshusingxml-method-visio.md)** method.

If the  **AddFromXML** method succeeds, it performs the following actions:


- Creates an  **DataRecordset** object and assigns it the name you specify in the Name parameter. If you do not specify a name, Visio assigns the data recordset the name of the database table that is the source of the data.
    
- Maps the data types of the columns of the data source to equivalent Visio data types, while filtering the results to remove data-source columns that cannot be linked to Visio shapes because they have no equivalent Visio data type. 
    
-  Assigns a Visio data-row ID to each row in the data recordset, unless the imported data already contains valid Visio data-row IDs. For more information about Visio data-row IDs, see the **[DataRecordset.GetDataRowIDs ](datarecordset-getdatarowids-method-visio.md)** topic.
    

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how you can use the  **AddFromXML** method to connect a Visio drawing to data contained in an ADO XML string.

A sample XML string is shown here. When it is passed to the  **AddFromXML** method, this string creates a data recordset that contains one column, named "Cities", and two data rows with entries in that column consisting of city names.




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
<z:row c1='Seattle' /> 
<z:row c1='Redmond' /> 
</rs:data> 
</xml>
```

In the following sample code, we pass the  **AddFromXML** method the name of an XML string containing data and the name of a string containing the display name we want to assign to the new data recordset being created.




```vb
Public Sub AddFromXML_Example() 
 
    Dim strXML As String 
    Dim strName As String 
    Dim vsoDataRecordset As Visio.DataRecordset 
 
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
    &; "<z:row c1='Seattle'/>" + Chr(10) _ 
    &; "<z:row c1='Redmond'/>" + Chr(10) _ 
    &; "</rs:data>" + Chr(10) _ 
    &; "</xml>" 
 
    strName = "City Names" 
 
    Set vsoDataRecordset = ThisDocument.DataRecordsets.AddFromXML(strXML, 0, strName) 
 
End Sub
```


