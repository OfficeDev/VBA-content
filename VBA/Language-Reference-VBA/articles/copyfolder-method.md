---
title: CopyFolder Method
keywords: vblr6.chm2182033
f1_keywords:
- vblr6.chm2182033
ms.prod: office
api_name:
- Office.CopyFolder
ms.assetid: d94788b4-9a92-77ea-6591-5ea2b4603233
ms.date: 06/08/2017
---


# CopyFolder Method



 **Description**
Recursively copies a folder from one location to another.
 **Syntax**
 _object_. **CopyFolder**_source_, _destination_ [, ove _r_ write]
The  **CopyFolder** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _source_|Required. Character string folder specification, which can include wildcard characters, for one or more folders to be copied.|
| _destination_|Required. Character string destination where the folder and subfolders from  _source_ are to be copied. Wildcard characters are not allowed.|
| _overwrite_|Optional.  **Boolean** value that indicates if existing folders are to be overwritten. If **True**, files are overwritten; if **False**, they are not. The default is **True**.|
 **Remarks**
Wildcard characters can only be used in the last path component of the  _source_ argument. For example, you can use:



```
FileSystemObject.CopyFolder "c:\mydocuments\letters\*", "c:\tempfolder\"

```

But you can't use:



```
FileSystemObject.CopyFolder "c:\mydocuments\*\*", "c:\tempfolder\"


```

If  _source_ contains wildcard characters or _destination_ ends with a path separator (\), it is assumed that _destination_ is an existing folder in which to copy matching folders and subfolders. Otherwise, _destination_ is assumed to be the name of a folder to create. In either case, four things can happen when an individual folder is copied.


- If  _destination_ does not exist, the _source_ folder and all its contents gets copied. This is the usual case.
    
- If  _destination_ is an existing file, an error occurs.
    
- If  _destination_ is a directory, an attempt is made to copy the folder and all its contents. If a file contained in _source_ already exists in _destination_, an error occurs if _overwrite_ is **False**. Otherwise, it will attempt to copy the file over the existing file.
    
- If  _destination_ is a read-only directory, an error occurs if an attempt is made to copy an existing read-only file into that directory and _overwrite_ is **False**.
    

An error also occurs if a  _source_ using wildcard characters doesn't match any folders.
The  **CopyFolder** method stops on the first error it encounters. No attempt is made to roll back any changes made before an error occurs.

