---
title: Document.ExecuteLine Method (Visio)
keywords: vis_sdr.chm10516260
f1_keywords:
- vis_sdr.chm10516260
ms.prod: visio
api_name:
- Visio.Document.ExecuteLine
ms.assetid: 0443c879-b569-c35b-e28c-77d0bf4b23ba
ms.date: 06/08/2017
---


# Document.ExecuteLine Method (Visio)

Executes a line of Microsoft Visual Basic code.


## Syntax

 _expression_ . **ExecuteLine**( **_Line_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Line_|Required| **String**|A string that will be interpreted as Microsoft Visual Basic for Applications (VBA) code.|

### Return Value

Nothing


## Remarks

The VBA project of the  **Document** object is told to execute the supplied string. VBA treats the string as it would treat the same string typed into its Immediate window.


## Example

The following are some possible uses of the  **ExecuteLine** method:


```vb
'Executes the macro (procedure without an argument) named "SomeMacro" 
 'that is in some module of the Visual Basic project of ThisDocument. 
 ThisDocument.ExecuteLine("SomeMacro ") 
 
 'Executes the procedure named SomeProcedure and passes it 3 arguments. 
 ThisDocument.ExecuteLine("SomeProcedure  1, 2, 3") 
 
 'Same as previous example, but procedure name qualified 
 'with module name. 
 ThisDocument.ExecuteLine("Module1.SomeProcedure  1, 2, 3") 
 
 'Shows the form UserForm1. 
 ThisDocument.ExecuteLine("UserForm1.Show") 
 
 'Prints "some string" to the Immediate window. 
 ThisDocument.ExecuteLine("Debug.Print ""some string """) 
 
 'Prints number of open documents to the Immediate window. 
 ThisDocument.ExecuteLine("Debug.Print Documents.Count") 
 
 'Tells ThisDocument to save itself. 
 ThisDocument.ExecuteLine("ThisDocument.Save") 

```


