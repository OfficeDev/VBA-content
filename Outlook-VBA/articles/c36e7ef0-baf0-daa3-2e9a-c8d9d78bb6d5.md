
# DistListItem.Session Property (Outlook)

 **Last modified:** July 28, 2015

Returns the  ** [NameSpace](f0dcaa19-07f5-5d42-a3bf-2e42b7885644.md)**object for the current session. Read-only.

## Syntax

 _expression_. **Session**

 _expression_A variable that represents a  **DistListItem** object.


## Remarks

The  **Session** property and the ** [GetNamespace](6175d0d9-5a61-ce45-35c0-b70895d757b3.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```
Set objNamespace = Application.GetNamespace("MAPI") 
```


```
Set objSession = Application.Session
```


## See also


#### Concepts


 [DistListItem Object](027c3986-abff-d9b1-ecc2-26d60805e952.md)
#### Other resources


 [DistListItem Object Members](3ba4af84-ce84-61d9-1bc9-fab41bf6f125.md)
