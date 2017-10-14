---
title: Seek Function
ms.prod: office
ms.assetid: 870aba03-b7ad-c931-928d-33aaf9cf5ab6
ms.date: 06/08/2017
---


# Seek Function



Returns a [Long](vbe-glossary.md) specifying the current read/write position within a file opened using the **Open** statement.
 **Syntax**
 **Seek(**_filenumber_**)**
The required  _filenumber_[argument](vbe-glossary.md) is an[Integer](vbe-glossary.md) containing a valid[file number](vbe-glossary.md).
 **Remarks**
 **Seek** returns a value between 1 and 2,147,483,647 (equivalent to 2^31 - 1), inclusive.
The following describes the return values for each file access mode.


|**Mode**|**Return Value**|
|:-----|:-----|
|**Random**|Number of the next record read or written|
|**Binary**, **Output**, **Append**, **Input**|Byte position at which the next operation takes place. The first byte in a file is at position 1, the second byte is at position 2, and so on.|

## Example

This example uses the  **Seek** function to return the current file position. The example assumes `TESTFILE` is a file containing records of the user-defined type is a file containing records of the user-defined type `Record`.


```vb
Type Record    ' Define user-defined type.
    ID As Integer
    Name As String * 20
End Type
```

For files opened in Random mode,  **Seek** returns number of next record.




```vb
Dim MyRecord As Record    ' Declare variable.
Open "TESTFILE" For Random As #1 Len = Len(MyRecord)
Do While Not EOF(1)    ' Loop until end of file.
    Get #1, , MyRecord    ' Read next record.
    Debug.Print Seek(1)    ' Print record number to the 
            ' Immediate window.
Loop
Close #1    ' Close file.

```

For files opened in modes other than Random mode,  **Seek** returns the byte position at which the next operation takes place. Assume `TESTFILE` is a file containing a few lines of text.




```vb
Dim MyChar
Open "TESTFILE" For Input As #1    ' Open file for reading.
Do While Not EOF(1)    ' Loop until end of file.
    MyChar = Input(1, #1)    ' Read next character of data.
    Debug.Print Seek(1)    ' Print byte position to the
            ' Immediate window.
Loop
Close #1    ' Close file.
```


