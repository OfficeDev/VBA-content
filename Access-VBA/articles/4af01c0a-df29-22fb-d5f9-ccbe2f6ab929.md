
# ODBCConnection.Creator Property (Excel)

 **Last modified:** July 28, 2015

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.

## Syntax

 _expression_. **Creator**

 _expression_A variable that represents an  **ODBCConnection** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


 [ODBCConnection Object](b880ebec-15a4-5a3d-ef02-db73106db9c9.md)
#### Other resources


 [ODBCConnection Object Members](d13b91f3-a89f-7dd7-7a98-f1d952f3b047.md)
