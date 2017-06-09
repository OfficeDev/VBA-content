---
title: CustomXMLSchema.Delete Method (Office)
keywords: vbaof11.chm291004
f1_keywords:
- vbaof11.chm291004
ms.prod: office
api_name:
- Office.CustomXMLSchema.Delete
ms.assetid: bdd79a25-7f2f-c810-13b0-9d7dc34e9a3d
ms.date: 06/08/2017
---


# CustomXMLSchema.Delete Method (Office)

Deletes the specified schema from the  **CustomXMLSchema** collection.


## Syntax

 _expression_. **Delete**

 _expression_ An expression that returns a **CustomXMLSchema** object.


## Remarks

If this operation is attempted on a schema in a collection that is already validated or attached to a data stream, then the operation is not performed and an error message is displayed.


## Example

The following example adds a schema to the collection and then deletes the schema.


```
Sub DeleteSchema() 
    On Error GoTo Err 
 
    Dim objCustomXMLSchemaCollection As CustomXMLSchemaCollection 
    Dim objCustomXMLSchema As  CustomXMLSchema 
 
    ' Adds a schema to the collection. 
    objCustomXMLSchema.Add("urn:invoice:namespace")  
 
    ... 
 
    ' Deletes the schema. 
    objCustomXMLSchema.Delete 
      
    Exit Sub 
                 
' Exception handling. Show the message and resume. 
Err: 
        MsgBox (Err.Description) 
        Resume Next 
End Sub
```


## See also


#### Concepts


[CustomXMLSchema Object](customxmlschema-object-office.md)
#### Other resources


[CustomXMLSchema Object Members](customxmlschema-members-office.md)

