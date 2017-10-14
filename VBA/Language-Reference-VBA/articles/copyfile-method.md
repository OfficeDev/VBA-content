---
title: CopyFile Method
keywords: vblr6.chm2182032
f1_keywords:
- vblr6.chm2182032
ms.prod: office
api_name:
- Office.CopyFile
ms.assetid: 2ab700b1-0827-c277-6af5-93a86ed05cc1
ms.date: 06/08/2017
---


# CopyFile Method



 **Description**
Copies one or more files from one location to another.
 **Syntax**
 _object_. **CopyFile**_source_, _destination_ [, _overwrite_ ]
The  **CopyFile** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. The  _object_ is always the name of a **FileSystemObject**.|
| _source_|Required. Character string file specification, which can include wildcard characters, for one or more files to be copied.|
| _destination_|Required. Character string destination where the file or files from  _source_ are to be copied. Wildcard characters are not allowed.|
| _overwrite_|Optional.  **Boolean** value that indicates if existing files are to be overwritten. If **True**, files are overwritten; if **False**, they are not. The default is **True**. Note that **CopyFile** will fail if _destination_ has the read-only attribute set, regardless of the value of _overwrite_.|
 **Remarks**
Wildcard characters can only be used in the last path component of the  _source_ argument. For example, you can use:



```vb
FileSystemObject.CopyFile "c:\mydocuments\letters\*.doc", "c:\tempfolder\"

```

But you can't use:



```vb
FileSystemObject.CopyFile "c:\mydocuments\*\R1???97.xls", "c:\tempfolder"


```

If  _source_ contains wildcard characters or _destination_ ends with a path separator ( **\** ), it is assumed that _destination_ is an existing folder in which to copy matching files. Otherwise, _destination_ is assumed to be the name of a file to create. In either case, three things can happen when an individual file is copied.


- If  _destination_ does not exist, _source_ gets copied. This is the usual case.
    
- If  _destination_ is an existing file, an error occurs if _overwrite_ is **False**. Otherwise, an attempt is made to copy _source_ over the existing file.
    
- If  _destination_ is a directory, an error occurs.
    

An error also occurs if a  _source_ using wildcard characters doesn't match any files. The **CopyFile** method stops on the first error it encounters. No attempt is made to roll back or undo any changes made before an error occurs.

