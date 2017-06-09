---
title: Width  Statement
keywords: vblr6.chm1009060
f1_keywords:
- vblr6.chm1009060
ms.prod: office
ms.assetid: 655e73fc-c294-5f82-4c1a-59c2ebd71036
ms.date: 06/08/2017
---


# Width # Statement

Assigns an output line width to a file opened using the  **Open** statement.

 **Syntax**

 **Width #**_filenumber_, _width_

The  **Width #** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](vbe-glossary.md).|
| _width_|Required. [Numeric expression](vbe-glossary.md) in the range 0-255, inclusive, that indicates how many characters appear on a line before a new line is started. If _width_ equals 0, there is no limit to the length of a line. The default value for _width_ is 0.|

## Example

This example uses the  **Width #** statement to set the output line width for a file.


```vb
Dim I 
Open "TESTFILE" For Output As #1 ' Open file for output. 
VBA.Width 1, 5 ' Set output line width to 5. 
For I = 0 To 9 ' Loop 10 times. 
 Print #1, Chr(48 + I); ' Prints five characters per line. 
Next I 
Close #1 ' Close file. 

```


