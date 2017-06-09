---
title: CustomXMLSchema.Reload Method (Office)
keywords: vbaof11.chm291005
f1_keywords:
- vbaof11.chm291005
ms.prod: office
api_name:
- Office.CustomXMLSchema.Reload
ms.assetid: 963b941a-0b93-fc02-c150-747975005561
ms.date: 06/08/2017
---


# CustomXMLSchema.Reload Method (Office)

Reloads a schema from a file.


## Syntax

 _expression_. **Reload**

 _expression_ An expression that returns a **CustomXMLSchema** object.


## Remarks

Typically, this method is used to update the location of the schema or to determine if the schema is still valid. It is also useful for reloading a schema that frequently changes. If this action is attempted on a schema in a collection that is already validated or tied to a data stream, then the operation is not performed and an error message is displayed.


## Example

The following example specifies the location of the schema and then reloads it.


```
Dim objCustomXMLSchema As  CustomXMLSchema 
Dim strSchemaLocation As String 
' Set the location of the schema.. 
objCustomXMLSchema.Location = "c:\mySchema.xsd" 
 
' Reload the schema. 
objCustomXMLSchema.Reload 

```


## See also


#### Concepts


[CustomXMLSchema Object](customxmlschema-object-office.md)
#### Other resources


[CustomXMLSchema Object Members](customxmlschema-members-office.md)

