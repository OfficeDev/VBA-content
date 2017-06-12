---
title: Compiler Constants
keywords: vblr6.chm1020792
f1_keywords:
- vblr6.chm1020792
ms.prod: office
ms.assetid: bde15ce4-af30-1bbf-7d34-4cfa7e396261
ms.date: 06/08/2017
---


# Compiler Constants

Visual Basic for Applications defines [constants](vbe-glossary.md) for exclusive use with the **#If...Then...#Else** directive. These constants are functionally equivalent to constants defined with the **#If...Then...#Else** directive except that they are global in[scope](vbe-glossary.md); that is, they apply everywhere in a [project](vbe-glossary.md).


 **Note**  Because  **Win32** returns true in both 32-bit and 64-bit development platforms it is important that the order within the **#If...Then...#Else** directive returns the desired results in your code. For example, because **Win32** returns True in 64-bit ( **Win32** is compatible in **Win64** environments) checking for **Win32** before **Win64** results in the **Win64** condition never running because **Win32** returns True. The following order returns predictable results:


```vb
#If Win64 Then 
' Win64=true, Win32=true, Win16= false 
#ElseIf Win32 Then 
' Win32=true, Win16=false 
#Else 
' Win16=true 
#End If
```

This applies to both Winx and VBAx constants.
On 16-bit development platforms, the compiler constants are defined as follows:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**Win16**|**True**|Indicates development environment is 16-bit compatible.|
|**Win32**|**False**|Indicates that the development environment is not 32-bit compatible.|
|**Win64**|**False**|Indicates that the development environment is not 64-bit compatible.|
On 32-bit development platforms, the compiler constants are defined as follows:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**Vba6**|**True**|Indicates that the development environment is Visual Basic for Applications, version 6.0 compatible.|
|**Vba6**|**False**|Indicates that the development environment is not Visual Basic for Applications, version 6.0 compatible.|
|**Vba7**|**True**|Indicates that the development environment is Visual Basic for Applications, version 7.0 compatible.|
|**Vba7**|**False**|Indicates that the development environment is not Visual Basic for Applications, version 7.0 compatible.|
|**Win16**|**False**|Indicates that the development environment is not 16-bit compatible.|
|**Win32**|**True**|Indicates that the development environment is 32-bit compatible.|
|**Win64**|**False**|Indicates that the development environment is not 64-bit compatible.|
|**Mac**|**True**|Indicates that the development environment is Macintosh.|
|**Mac**|**False**|Indicates that the development environment is not Macintosh.|
On 64-bit development platforms, the compiler constants are defined as follows:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**Vba6**|**True**|Indicates that the development environment is Visual Basic for Applications, version 6.0 compatible.|
|**Vba6**|**False**|Indicates that the development environment is not Visual Basic for Applications, version 6.0 compatible.|
|**Vba7**|**True**|Indicates that the development environment is Visual Basic for Applications, version 7.0 compatible.|
|**Vba7**|**False**|Indicates that the development environment is not Visual Basic for Applications, version 7.0 compatible.|
|**Win16**|**False**|Indicates development environment is not 16-bit compatible.|
|**Win32**|**True**|Indicates that the development environment is 32-bit compatible.|
|**Win64**|**True**|Indicates that the development environment is 64-bit compatible.|
|**Mac**|**True**|Indicates that the development environment is Macintosh.|
|**Mac**|**False**|Indicates that the development environment is not Macintosh.|

 **Note**  These constants are provided by Visual Basic, so you cannot define your own constants with these same names at any level.


