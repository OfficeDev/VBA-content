---
title: MoveFolder Method
keywords: vblr6.chm2182060
f1_keywords:
- vblr6.chm2182060
ms.prod: office
api_name:
- Office.MoveFolder
ms.assetid: 08a088c1-6e3c-d2a2-7708-f1682cafd91e
ms.date: 06/08/2017
---


# MoveFolder Method



 **Description**
Moves one or more folders from one location to another.
 **Syntax**
 _object_. **MoveFolder**_source_, _destination_
The  **MoveFolder** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _source_|Required. The path to the folder or folders to be moved. The  _source_ argument string can contain wildcard characters in the last path component only.|
| _destination_|Required. The path where the folder or folders are to be moved. The  _destination_ argument can't contain wildcard characters.|
 **Remarks**
If  _source_ contains wildcards or _destination_ ends with a path separator ( **\** ), it is assumed that _destination_ specifies an existing folder in which to move the matching files. Otherwise, _destination_ is assumed to be the name of a destination folder to create. In either case, three things can happen when an individual folder is moved:


- If  _destination_ does not exist, the folder gets moved. This is the usual case.
    
- If  _destination_ is an existing file, an error occurs.
    
- If  _destination_ is a directory, an error occurs.
    

An error also occurs if a wildcard character that is used in  _source_ doesn't match any folders. The **MoveFolder** method stops on the first error it encounters. No attempt is made to roll back any changes made before the error occurs.


 **Important**  This method allows moving folders between volumes only if supported by the operating system.



