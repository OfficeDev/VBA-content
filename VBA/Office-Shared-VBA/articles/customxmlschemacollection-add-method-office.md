---
title: CustomXMLSchemaCollection.Add Method (Office)
keywords: vbaof11.chm292005
f1_keywords:
- vbaof11.chm292005
ms.prod: office
api_name:
- Office.CustomXMLSchemaCollection.Add
ms.assetid: d5df782b-0e8b-e827-4cb4-40ddb9731e9b
ms.date: 06/08/2017
---


# CustomXMLSchemaCollection.Add Method (Office)

Allows you to add one or more schemas to a schema collection that can then be added to a stream in the data store and to the Schema Library. 


## Syntax

 _expression_. **Add**( **_NamespaceURI_**, **_Alias_**, **_FileName_**, **_InstallForAllUsers_** )

 _expression_ An expression that returns a **CustomXMLSchemaCollection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NamespaceURI_|Optional|**String**|Contains the namespace of the schema to be added to the collection. If the schema already exists in the Schema Library, the method will retrieve it from there.|
| _Alias_|Optional|**String**|Contains the alias of the schema to be added to the collection. If the alias already exists in the Schema Library, the method can find it using this argument.|
| _FileName_|Optional|**String**|Contains the location of the schema on a disk. If this parameter is specified, the schema is added to the collection and to to the Schema Library.|
| _InstallForAllUsers_|Optional|**Boolean**|Specifies whether, in the case where the method is adding the schema to the Schema Library, the Schema Library keys should be written to the registry(HKey_Local_Machine for all users or HKey_Current_User for just the current user). The parameter defaults to  **False** and writes to HKey_Current_User.|

### Return Value

CustomXMLSchema


## Example

The following example adds a schema to the schema collection, selects a single node from it, and then returns the node to the calling procedure.


```
Function AddSchema() 
    On Error GoTo Err 
 
    Dim objCustomXMLSchemaCollection As CustomXMLSchemaCollection 
    Dim cxp1 As  CustomXMLSchema 
    Dim cxn As CustomXMLNode 
 
    ' Adds a schema to the collection. 
    cxp1 = objCustomXMLSchemaCollection.Add("urn:invoice:namespace", "coreDefinitions", wdCore.xsd", True) 
 
... 
 
    Set cxn = cxp4.SelectSingleNode("//*[@quantity < 4]") 
 
    AddSchema = cxn 
      
    Exit Function 
                 
' Exception handling. Show the message and resume. 
Err: 
        MsgBox (Err.Description) 
        Resume Next 
End Function 

```


## See also


#### Concepts


[CustomXMLSchemaCollection Object](customxmlschemacollection-object-office.md)
#### Other resources


[CustomXMLSchemaCollection Object Members](customxmlschemacollection-members-office.md)

