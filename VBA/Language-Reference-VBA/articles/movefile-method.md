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


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                               |
|:----------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. Always the name of a  <strong>FileSystemObject</strong>.                                                                                         |
| <em>source</em>       | Required. The path to the file or files to be moved. The  <em>source</em> argument string can contain wildcard characters in the last path component only. |
| <em>destination</em>  | Required. The path where the file or files are to be moved. The  <em>destination</em> argument can't contain wildcard characters.                          |

 <strong>Remarks</strong>
If  
<em>source</em> contains wildcards or <em>destination</em> ends with a path separator ( <strong>\</strong> ), it is assumed that <em>destination</em> specifies an existing folder in which to move the matching files. Otherwise, <em>destination</em> is assumed to be the name of a destination file to create. In either case, three things can happen when an individual file is moved:


- If  _destination_ does not exist, the file gets moved. This is the usual case.

- If  _destination_ is an existing file, an error occurs.

- If desti _n_ ation is a directory, an error occurs.


An error also occurs if a wildcard character that is used in  _source_ doesn't match any files. The **MoveFile** method stops on the first error it encounters. No attempt is made to roll back any changes made before the error occurs.


 **Important**  This method allows moving files between volumes only if supported by the operating system.



