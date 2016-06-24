
# OLEDBError.Number Property (Excel)

Returns a numeric value that specifies an error. The error number corresponds to a unique trap number corresponding to an error condition that resulted after the most recent OLE DB query. Read-only  **Long** .


## Syntax

 _expression_ . **Number**

 _expression_ A variable that represents an **OLEDBError** object.


## Example

This example displays the error number and other error information returned by the most recent OLE DB query.


```vb
Set objEr = Application.OLEDBErrors(1) 
MsgBox "The following error occurred:" &; _ 
 objEr.Number &; ", " &; objEr.Native &; ", " &; _ 
 objEr.ErrorString &; " : " &; objEr.SqlState
```


## See also


#### Concepts


[OLEDBError Object](6bcbf721-f2c8-f784-361b-e1a298bb2ecb.md)
#### Other resources


[OLEDBError Object Members](52181252-dd6f-b267-fa21-4ad8175b7346.md)
