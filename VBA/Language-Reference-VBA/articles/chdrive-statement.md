---
title: ChDrive Statement
keywords: vblr6.chm1008865
f1_keywords:
- vblr6.chm1008865
ms.prod: office
ms.assetid: b07d5925-fba0-9a50-8197-c782fda0bee5
ms.date: 06/08/2017
---


# ChDrive Statement

Changes the current drive.

 **Syntax**

 **ChDrive**_drive_

The required  _drive_[argument](vbe-glossary.md) is a[string expression](vbe-glossary.md) that specifies an existing drive. If you supply a zero-length string (""), the current drive doesn't change. If the _drive_ argument is a multiple-character string, **ChDrive** uses only the first letter.
On the Macintosh,  **ChDrive** changes the current folder to the root folder of the specified drive.

## Example

This example uses the  **ChDrive** statement to change the current drive. On the Macintosh, "HD:" is the default drive name and **ChDrive** would change the current folder to the root folder of the specified drive. The following example assumes the machine actually has a drive named D.


```
ChDrive "D" ' Make "D" the current drive. 

```


