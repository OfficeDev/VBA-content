
# Reset Statement

 **Last modified:** July 28, 2015

Closes all disk files opened using the  **Open** statement.

 **Syntax**

 **Reset**
 **Remarks**
The  **Reset** statement closes all active files opened by the **Open** statement and writes the contents of all file buffers to disk.

## Example

This example uses the  **Reset** statement to close all open files and write the contents of all file buffers to disk. Note the use of the **Variant** variable as both a string and a number.


```
Dim FileNumber 
For FileNumber = 1 To 5 ' Loop 5 times. 
 ' Open file for output. FileNumber is concatenated into the string 
 ' TEST for the file name, but is a number following a #. 
 Open "TEST" &amp; FileNumber For Output As #FileNumber 
 Write #FileNumber, "Hello World" ' Write data to file. 
Next FileNumber 
Reset ' Close files and write contents 
 ' to disk. 

```

