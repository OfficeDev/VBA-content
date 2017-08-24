---
title: WebCommandButton.DataFileFormat Property (Publisher)
keywords: vbapb10.chm3932169
f1_keywords:
- vbapb10.chm3932169
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.DataFileFormat
ms.assetid: 7594b575-b39f-3cd4-d0b9-c13c04299345
ms.date: 06/08/2017
---


# WebCommandButton.DataFileFormat Property (Publisher)

Sets or returns a  **PbSubmitDataFormatType** constant that represents the format to use when saving Web form data to a file. Read/write.


## Syntax

 _expression_. **DataFileFormat**

 _expression_A variable that represents a  **WebCommandButton** object.


### Return Value

PbSubmitDataFormatType


## Remarks

The  **DataFileFormat** property value can be one of the **[PbSubmitDataFormatType](pbsubmitdataformattype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example sets Microsoft Publisher to process Web form data by saving it to a comma-delimited text file on the same Web server as the form is stored. (Note that Filename must be replaced with a valid file name for this example to work.)


```vb
Sub WebDataFile() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalSaveOnServer 
 .DataFileFormat = pbSubmitDataFormatCSV 
 .DataFileName = "Filename" 
 End With 
End Sub
```


