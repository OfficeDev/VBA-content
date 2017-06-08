---
title: CustomXMLSchemaCollection.Validate Method (Office)
keywords: vbaof11.chm292007
f1_keywords:
- vbaof11.chm292007
ms.prod: office
api_name:
- Office.CustomXMLSchemaCollection.Validate
ms.assetid: c1358676-9df7-83fe-1b6c-8ef70f9d9c4b
ms.date: 06/08/2017
---


# CustomXMLSchemaCollection.Validate Method (Office)

Specifies whether the schemas in a schema collection are valid (conforms to the syntactic rules of XML and the rules for a specified vocabulary; a standard for structuring XML).


## Syntax

 _expression_. **Validate**

 _expression_ An expression that returns a **CustomXMLSchemaCollection** object.


### Return Value

Boolean


## Remarks

In addition to determining whether the schemas are valid, this method also traverses the  **include** statements for each schema in the collection and adds the referenced schemas to the source schema.


## Example

The following example validates the schema collection and returns the  **Boolean** results to the calling procedure.


```
Function ValidateSchemas(objSourceCustomXMLSchemaCollection As CustomXMLSchemaCollection) 
Dim boolValid As Boolean 
 
' Validates the schemas in a schema collection. 
boolValid = objSourceCustomXMLSchemaCollection.Validate 
 
ValidateSchemas = boolValid 
   
End Function
```


## See also


#### Concepts


[CustomXMLSchemaCollection Object](customxmlschemacollection-object-office.md)
#### Other resources


[CustomXMLSchemaCollection Object Members](customxmlschemacollection-members-office.md)

