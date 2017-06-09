---
title: CustomXMLSchemaCollection.AddCollection Method (Office)
keywords: vbaof11.chm292006
f1_keywords:
- vbaof11.chm292006
ms.prod: office
api_name:
- Office.CustomXMLSchemaCollection.AddCollection
ms.assetid: d3b49c57-9a5b-9b5b-0003-d09240d227c1
ms.date: 06/08/2017
---


# CustomXMLSchemaCollection.AddCollection Method (Office)

Adds an existing schema collection to the current schema collection. 


## Syntax

 _expression_. **AddCollection**( **_SchemaCollection_** )

 _expression_ An expression that returns a **CustomXMLSchemaCollection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchemaCollection_|Required|**CustomXMLSchemaCollection**|Represents a collection of schemas to be imported into the current schema collection.|

## Remarks

If there is a conflict between namespaces while importing the collection, for example if x.xsd is already linked to "urn:invoice:namespace" but the incoming collection has z.xsd for the same namespace, the incoming collection wins.


## Example

The following example receives the target schema collection and incoming schema collection arguments and then adds the one collection to the other.


```
Sub AddSchema(objTargetCustomXMLSchemaCollection As CustomXMLSchemaCollection, _ 
  objTargetCustomXMLSchemaCollection As CustomXMLSchemaCollection) 
 
    ' Adds a schema collection to another schema the collection. 
    objTargetCustomXMLSchemaCollection.AddCollection(objIncomingCustomXMLSchemaCollection) 
                
End Sub
```


## See also


#### Concepts


[CustomXMLSchemaCollection Object](customxmlschemacollection-object-office.md)
#### Other resources


[CustomXMLSchemaCollection Object Members](customxmlschemacollection-members-office.md)

