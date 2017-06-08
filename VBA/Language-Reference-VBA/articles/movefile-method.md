---
title: MoveFile Method
keywords: vblr6.chm2182059
f1_keywords:
- vblr6.chm2182059
ms.prod: office
api_name:
- Office.MoveFile
ms.assetid: 1b5dec21-8333-1bc6-0088-6999051beaa4
ms.date: 06/08/2017
---


# MoveFile Method



 **Description**
Moves one or more files from one location to another.
 **Syntax**
 _object_. **MoveFile**_source_, _destination_
The  **MoveFile** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _source_|Required. The path to the file or files to be moved. The  _source_ argument string can contain wildcard characters in the last path component only.|
| _destination_|Required. The path where the file or files are to be moved. The  _destination_ argument can't contain wildcard characters.|
 **Remarks**
If  _source_ contains wildcards or _destination_ ends with a path separator ( **\** ), it is assumed that _destination_ specifies an existing folder in which to move the matching files. Otherwise, _destination_ is assumed to be the name of a destination file to create. In either case, three things can happen when an individual file is moved:


- If  _destination_ does not exist, the file gets moved. This is the usual case.
    
- If  _destination_ is an existing file, an error occurs.
    
- If desti _n_ ation is a directory, an error occurs.
    

An error also occurs if a wildcard character that is used in  _source_ doesn't match any files. The **MoveFile** method stops on the first error it encounters. No attempt is made to roll back any changes made before the error occurs.


 **Important**  This method allows moving files between volumes only if supported by the operating system.



