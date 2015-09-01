
# ContactItem.Session Property (Outlook)

 **Last modified:** July 28, 2015

Returns the  ** [NameSpace](f0dcaa19-07f5-5d42-a3bf-2e42b7885644.md)**object for the current session. Read-only.

## Syntax

 _expression_. **Session**

 _expression_A variable that represents a  **ContactItem** object.


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


 [ContactItem Object](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)
#### Other resources


 [ContactItem Object Members](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)
