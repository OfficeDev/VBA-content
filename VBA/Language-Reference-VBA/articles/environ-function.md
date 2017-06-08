---
title: Environ Function
keywords: vblr6.chm1013110
f1_keywords:
- vblr6.chm1013110
ms.prod: office
ms.assetid: ad8cb911-5dab-a327-bd9c-ee57818a713a
ms.date: 06/08/2017
---


# Environ Function



Returns the  **String** associated with an operating system environment variable. Not available on the Macintosh
 **Syntax**
 **Environ(** { **_envstring_** |**_number_** } **)**
The  **Environ** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_envstring_**|Optional. [String expression](vbe-glossary.md) containing the name of an environment variable.|
|**_number_**|Optional. [Numeric expression](vbe-glossary.md) corresponding to the numeric order of the environment string in the environment-string table. The **_number_**[argument](vbe-glossary.md) can be any numeric expression, but is rounded to a whole number before it is evaluated.|
 **Remarks**
If  **_envstring_** can't be found in the environment-string table, a zero-length string ("") is returned. Otherwise, **Environ** returns the text assigned to the specified **_envstring_**; that is, the text following the equal sign (=) in the environment-string table for that environment variable.
If you specify  **_number_**, the string occupying that numeric position in the environment-string table is returned. In this case, **Environ** returns all of the text, including **_envstring_**. If there is no environment string in the specified position, **Environ** returns a zero-length string.

## Example

This example uses the  **Environ** function to supply the entry number and length of the `PATH` statement from the environment-string table. Not available on the Macintosh.


```vb
Dim EnvString, Indx, Msg, PathLen    ' Declare variables.
Indx = 1    ' Initialize index to 1.
Do
    EnvString = Environ(Indx)    ' Get environment 
                ' variable.
    If Left(EnvString, 5) = "PATH=" Then    ' Check PATH entry.
        PathLen = Len(Environ("PATH"))    ' Get length.
        Msg = "PATH entry = " &; Indx &; " and length = " &; PathLen
        Exit Do
    Else
        Indx = Indx + 1    ' Not PATH entry,
    End If    ' so increment.
Loop Until EnvString = ""
If PathLen > 0 Then
    MsgBox Msg    ' Display message.
Else
    MsgBox "No PATH environment variable exists."
End If

```


