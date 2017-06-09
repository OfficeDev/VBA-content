---
title: DataFeedConnection.SaveAsODC Method (Excel)
keywords: vbaxl10.chm928088
f1_keywords:
- vbaxl10.chm928088
ms.prod: excel
ms.assetid: e66ff66c-9b19-a479-0afa-4f7e307113ac
ms.date: 06/08/2017
---


# DataFeedConnection.SaveAsODC Method (Excel)

Saves the data feed connection as a Microsoft Office Data Connection file.


## Syntax

 _expression_ . **SaveAsODC**_(ODCFileName,_ _Description,_ _Keywords)_

 _expression_ A variable that represents a[DataFeedConnection Object (Excel)](datafeedconnection-object-excel.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ODCFileName_|Required|STRING|Location to save the file.|
| _Description_|Optional|VARIANT|Description that will be saved in the file.|
| _Keywords_|Optional|VARIANT|Space-separated keywords that can be used to search for this file.|

### Example

The following example saves the connection as an ODC file titled "ODCFile". This example assumes data feed connection exists on the active worksheet. 


```vb
Sub UseSaveAsODC() 
 
   Application.ActiveWorkbook.Connections("Datafeed1").DataFeedConnection.SaveAsODC ("ODCFile")
 
End Sub
```


### Return value

 **VOID**


## See also


#### Other resources



[DataFeedConnection Object](datafeedconnection-object-excel.md)

