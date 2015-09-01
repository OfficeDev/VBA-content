
# ProtectedViewWindow.Activate Method (Word)

 **Last modified:** July 28, 2015

Activates the specified protected view window.

## Syntax

 _expression_. **Activate**

 _expression_An expression that returns a  ** [ProtectedViewWindow Object](d77e80e7-c54e-5954-1586-dacd3c9f7434.md)** object.


### Return Value

Nothing


## Example

The following code example activates the next protected view window in the  [ProtectedViewWindows](62c2f4d5-1080-548e-730b-388308144dfe.md) collection.


```
Dim pvWindow As ProtectedViewWindow 
 
' At least one document must be open in protected view for this statement to execute. 
ProtectedViewWindows(1).Activate
```


## See also


#### Concepts


 [ProtectedViewWindow Object](d77e80e7-c54e-5954-1586-dacd3c9f7434.md)
#### Other resources


 [ProtectedViewWindow Object Members](03a8f0c3-f76b-f933-9cae-5a159234c289.md)
