---
title: ChDir Statement
keywords: vblr6.chm1008864
f1_keywords:
- vblr6.chm1008864
ms.prod: office
ms.assetid: a2ad61c8-5b69-8096-3176-52e7670f58ab
ms.date: 06/08/2017
---


# ChDir Statement

Changes the current directory or folder.


## Syntax

**ChDir** _path_

The required  _path_ [argument](vbe-glossary.md) is a [string expression](vbe-glossary.md) that identifies which directory or folder becomes the new default directory or folder. The _path_ may include the drive. If no drive is specified, **ChDir** changes the default directory or folder on the current drive.


## Remarks

The  **ChDir** statement changes the default directory or folder but does not change the default drive. A different statement, **[ChDrive](chdrive-statement.md)**, changes the default drive. 


```
ChDir "D:\TMP" ' Make "D:\TMP" the current folder. 

ChDrive "D"    ' Make "D" the current drive. 

```


On the Power Macintosh, the default drive always changes to the drive specified in  _path_. Full path specifications begin with the volume name, and relative paths begin with a colon ( **:** ). **ChDir** resolves any aliases specified in the path:


```
ChDir "MacDrive:Tmp" ' On the Macintosh. 

```


Note that when making relative directory changes, different symbols are used in Microsoft Windows and on the Macintosh:


```
ChDir ".." ' Moves up one directory in Microsoft Windows. 
ChDir "::" ' Moves up one directory on the Macintosh.
```


On the Macintosh, the default drive name is "HD" and portions of the pathname are separated by colons instead of backslashes. Similarly, you would specify Macintosh folders instead of Windows. Finally, wildcard characters have no special meaning on the Macintosh and are treated simply as characters.


## Example

This example uses the  **ChDir** statement to change the current directory or folder. If the default drive is C, then default drive remains C, even though the default folder changes to a folder on drive D:

```
' Assume "C:" is the current drive. The following statement changes 
' the default directory on drive "D:". "C:" remains the current drive. 
ChDir "D:\WINDOWS\SYSTEM" 

```


## See also


#### Reference


[ChDrive](chdrive-statement.md)

[CurDir](curdir-function.md)

[MkDir](mkdir-statement.md)

[RmDir](rmdir-statement.md)

