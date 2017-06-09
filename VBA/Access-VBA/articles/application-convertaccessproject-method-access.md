---
title: Application.ConvertAccessProject Method (Access)
keywords: vbaac10.chm12598
f1_keywords:
- vbaac10.chm12598
ms.prod: access
api_name:
- Access.Application.ConvertAccessProject
ms.assetid: 49b865f5-30b6-7b28-efe8-df2cc67951b0
ms.date: 06/08/2017
---


# Application.ConvertAccessProject Method (Access)

Converts the specified Microsoft Access file from one version to another.


## Syntax

 _expression_. **ConvertAccessProject**( ** _SourceFilename_**, ** _DestinationFilename_**, ** _DestinationFileFormat_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceFilename_|Required|**String**|The name of the Access file to convert. If a path isn't specified, Access looks for the file in the current directory.|
| _DestinationFilename_|Required|**String**|e name of the file where Access saves the converted file. If a path isn't specified, Access saves the file in the current directory.|
| _DestinationFileFormat_|Required|**AcFileFormat**|An  **[AcFileFormat](acfileformat-enumeration-access.md)** constant that specifies the format of the converted file.|

### Return Value

Nothing


## Remarks

The file specified by  _DestinationFilename_ cannot already exist, or an error occurs.


## Example

The following example converts the specified Access 97 file to an Access 2000 file in the same directory.


```vb
Application.ConvertAccessProject _ 
 SourceFilename:="C:\My Documents\Sales-Access97.mdb", _ 
 DestinationFilename:="C:\My Documents\Sales-Access2000.mdb", _ 
 DestinationFileFormat:=acFileFormatAccess2000 

```


## See also


#### Concepts


[Application Object](application-object-access.md)

