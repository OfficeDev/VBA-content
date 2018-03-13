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


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                            |
|:----------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. The  <em>object</em> is always the name of a <strong>FileSystemObject</strong>.                                                                                                                                                                                                                                                                                               |
| <em>source</em>       | Required. Character string file specification, which can include wildcard characters, for one or more files to be copied.                                                                                                                                                                                                                                                               |
| <em>destination</em>  | Required. Character string destination where the file or files from  <em>source</em> are to be copied. Wildcard characters are not allowed.                                                                                                                                                                                                                                             |
| <em>overwrite</em>    | Optional.  <strong>Boolean</strong> value that indicates if existing files are to be overwritten. If <strong>True</strong>, files are overwritten; if <strong>False</strong>, they are not. The default is <strong>True</strong>. Note that <strong>CopyFile</strong> will fail if <em>destination</em> has the read-only attribute set, regardless of the value of <em>overwrite</em>. |

 **Remarks**
Wildcard characters can only be used in the last path component of the  _source_ argument. For example, you can use:



```vb
FileSystemObject.CopyFile "c:\mydocuments\letters\*.doc", "c:\tempfolder\"
```

But you can't use:



```vb
FileSystemObject.CopyFile "c:\mydocuments\*\R1???97.xls", "c:\tempfolder"
```

If  <em>source</em> contains wildcard characters or <em>destination</em> ends with a path separator ( <strong>\</strong> ), it is assumed that <em>destination</em> is an existing folder in which to copy matching files. Otherwise, <em>destination</em> is assumed to be the name of a file to create. In either case, three things can happen when an individual file is copied.


- If  _destination_ does not exist, _source_ gets copied. This is the usual case.

- If  _destination_ is an existing file, an error occurs if _overwrite_ is **False**. Otherwise, an attempt is made to copy _source_ over the existing file.

- If  _destination_ is a directory, an error occurs.


An error also occurs if a  _source_ using wildcard characters doesn't match any files. The **CopyFile** method stops on the first error it encounters. No attempt is made to roll back or undo any changes made before an error occurs.

