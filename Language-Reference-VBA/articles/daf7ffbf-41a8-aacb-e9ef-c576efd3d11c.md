
# ProcOfLine Property

 **Last modified:** July 28, 2015


Returns the name of the  [procedure](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) that the specified line is in.
 **Syntax**
 _object_**.ProcOfLine(**_line_,  _prockind_**) As String**
The  **ProcOfLine** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An  [object expression](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) that evaluates to an object in the Applies To list.|
| _line_|Required. A  [Long](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) specifying the line to check.|
| _prockind_|Required. Specifies the kind of procedure to locate. Because  [property procedures](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) can have multiple representations in the [module](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is,  **Sub** and **Function** procedures) use **vbext_pk_Proc**.|
You can use one of the following  [constants](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) for the _prockind_ [argument](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md):


|**Constant**|**Description**|
|:-----|:-----|
| **vbext_pk_Get**|Specifies a procedure that returns the value of a property.|
| **vbext_pk_Let**|Specifies a procedure that assigns a value to a property.|
| **vbext_pk_Set**|Specifies a procedure that sets a reference to an object.|
| **vbext_pk_Proc**|Specifies all procedures other than property procedures.|
 **Remarks**
A line is within a procedure if it's a blank line or comment line preceding the procedure declaration and, if the procedure is the last procedure in a  [code module](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), a blank line or lines following the procedure.
