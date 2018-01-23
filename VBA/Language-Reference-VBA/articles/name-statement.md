---
title: Name Statement
keywords: vblr6.chm1008979
f1_keywords:
- vblr6.chm1008979
ms.prod: office
ms.assetid: c248e962-1265-b871-3ef7-36effb070d2b
ms.date: 06/08/2017
---


# Name Statement

Renames a disk file, directory, or folder.

 **Syntax**

 **Name**_oldpathname_**As**_newpathname_

The  **Name** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _oldpathname_|Required. [String expression](vbe-glossary.md) that specifies the existing file name and location — may include directory or folder, and drive.|
| _newpathname_|Required. String expression that specifies the new file name and location — may include directory or folder, and drive. The file name specified by  _newpathname_ can't already exist.|
 **Remarks**
The Name statement renames a file and moves it to a different directory or folder, if necessary. Name can move a file across drives, but it can only rename an existing directory or folder when both newpathname and oldpathname are located on the same drive. Name cannot create a new file, directory, or folder.
Using  **Name** on an open file produces an error. You must close an open file before renaming it. **Name**[arguments](vbe-glossary.md) cannot include multiple-character ( **\*** ) and single-character ( **?** ) wildcards.

## Example

This example uses the  **Name** statement to rename a file. For purposes of this example, assume that the directories or folders that are specified already exist. On the Macintosh, "HD:" is the default drive name and portions of the pathname are separated by colons instead of backslashes.


```vb
Dim OldName, NewName 
OldName = "OLDFILE": NewName = "NEWFILE" ' Define file names. 
Name OldName As NewName ' Rename file. 
 
OldName = "C:\MYDIR\OLDFILE": NewName = "C:\YOURDIR\NEWFILE" 
Name OldName As NewName ' Move and rename file. 

```


