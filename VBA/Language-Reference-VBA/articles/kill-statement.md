---
title: Kill Statement
keywords: vblr6.chm1008955
f1_keywords:
- vblr6.chm1008955
ms.prod: office
ms.assetid: 31ca6ed1-7f34-30a1-c990-96759d0f6c32
ms.date: 06/08/2017
---


# Kill Statement

Deletes files from a disk.

 **Syntax**

 **Kill**_pathname_

The required  _pathname_[argument](vbe-glossary.md) is a[string expression](vbe-glossary.md) that specifies one or more file names to be deleted. The _pathname_ may include the directory or folder, and the drive.
 **Remarks**
In Microsoft Windows,  **Kill** supports the use of multiple-character ( **\*** ) and single-character ( **?** ) wildcards to specify multiple files. However, on the Macintosh, these characters are treated as valid file name characters and can't be used as wildcards to specify multiple files.
Since the Macintosh doesn't support the wildcards, use the file type to identify groups of files to delete. You can use the  **MacID** function to specify file type instead of repeating the command with separate file names. For example, the following statement deletes all TEXT files in the current folder.



```vb
Kill MacID("TEXT") 

```

If you use the  **MacID** function with **Kill** in Microsoft Windows, an error occurs.
An error occurs if you try to use  **Kill** to delete an open file.

 **Note**  To delete directories, use the  **RmDir** statement.


## Example

This example uses the  **Kill** statement to delete a file from a disk.


```vb
' Assume TESTFILE is a file containing some data. 
Kill "TestFile" ' Delete file. 
 
' Delete all *.TXT files in current directory. 
Kill "*.TXT" 

```


