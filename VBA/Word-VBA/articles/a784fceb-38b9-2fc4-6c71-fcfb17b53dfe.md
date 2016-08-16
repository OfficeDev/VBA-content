
# ProtectedViewWindow.Activate Method (Word)

Activates the specified protected view window.


## Syntax

 _expression_ . **Activate**

 _expression_ An expression that returns a **[ProtectedViewWindow Object](d77e80e7-c54e-5954-1586-dacd3c9f7434.md)** object.


### Return Value

Nothing


## Example

The following code example activates the next protected view window in the [ProtectedViewWindows](62c2f4d5-1080-548e-730b-388308144dfe.md) collection.


```vb
Dim pvWindow As ProtectedViewWindow 
 
' At least one document must be open in protected view for this statement to execute. 
ProtectedViewWindows(1).Activate
```


## See also


#### Concepts


[ProtectedViewWindow Object](d77e80e7-c54e-5954-1586-dacd3c9f7434.md)
