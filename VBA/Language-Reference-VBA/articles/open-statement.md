---
title: Open Statement
keywords: vblr6.chm1008800
f1_keywords:
- vblr6.chm1008800
ms.prod: office
ms.assetid: 359a24b9-6dbb-3648-0ce4-98ec38441ccf
ms.date: 06/08/2017
---


# Open Statement

Enables input/output (I/O) to a file.

 **Syntax**

 **Open**_pathname_**For**_mode_ [ **Access**_access_ ] [ _lock_ ] **As** [ **#** ] _filenumber_ [ **Len** = _reclength_ ]

The  **Open** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _pathname_|Required. [String expression](vbe-glossary.md) that specifies a file name — may include directory or folder, and drive.|
| _mode_|Required. [Keyword](vbe-glossary.md) specifying the file mode: **Append**, **Binary**, **Input**, **Output**, or **Random**. If unspecified, the file is opened for **Random** access.|
| _access_|Optional. Keyword specifying the operations permitted on the open file:  **Read**, **Write**, or **Read Write**.|
| _lock_|Optional. Keyword specifying the operations restricted on the open file by other processes:  **Shared**, **Lock Read**, **Lock Write**, and **Lock Read Write**.|
| _filenumber_|Required. A valid [file number](vbe-glossary.md) in the range 1 to 511, inclusive. Use the **FreeFile** function to obtain the next available file number.|
| _reclength_|Optional. Number less than or equal to 32,767 (bytes). For files opened for random access, this value is the record length. For sequential files, this value is the number of characters buffered.|
 **Remarks**
You must open a file before any I/O operation can be performed on it.  **Open** allocates a buffer for I/O to the file and determines the mode of access to use with the buffer.
If the file specified by  _pathname_ doesn't exist, it is created when a file is opened for **Append**, **Binary**, **Output**, or **Random** modes.
If the file is already opened by another process and the specified type of access is not allowed, the  **Open** operation fails and an error occurs.
The  **Len** clause is ignored if _mode_ is **Binary**.


 **Important**  In  **Binary**, **Input**, and **Random** modes, you can open a file using a different file number without first closing the file. In **Append** and **Output** modes, you must close a file before opening it with a different file number.



## Example

This example illustrates various uses of the  **Open** statement to enable input and output to a file.

The following code opens the file in sequential-input mode.




```vb
Open "TESTFILE" For InputAs#1 
' Close before reopening in another mode. 
Close #1 

```

This example opens the file in Binary mode for writing operations only.




```vb
Open "TESTFILE" For Binary Access Write As #1 
' Close before reopening in another mode. 
Close #1 

```

The following example opens the file in Random mode. The file contains records of the user-defined type .




```vb
Type Record ' Define user-defined type. 
 ID As Integer 
 Name As String * 20 
End Type 
 
Dim MyRecord As Record ' Declare variable. 
Open "TESTFILE" For Random As #1 Len = Len(MyRecord) 
' Close before reopening in another mode. 
Close #1 

```

This code example opens the file for sequential output; any process can read or write to file.




```vb
Open "TESTFILE" For Output Shared As #1 
' Close before reopening in another mode. 
Close #1 
```

This code example opens the file in Binary mode for reading; other processes can't read file.




```vb
Open "TESTFILE" For Binary Access Read Lock Read As #1 
```


