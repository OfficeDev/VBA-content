
# Global.IsSandboxed Property (Word)

 **Last modified:** July 28, 2015

 **True** if the application window is a protected view window. Read-only.

## Syntax

 _expression_. **IsSandboxed**

 _expression_An expression that returns a  ** [Global](b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c.md)** object.


## Example

The following code example displays whether or not the document referenced by  _doc_ is in a protected view window.


```
If doc.Application.IsSandboxed Then 
 MsgBox "The document " &amp; _ 
 """" &amp; doc.Name &amp; """" &amp; _ 
 " is in a protected view window." 
Else 
 MsgBox "The document " &amp; _ 
 """" &amp; doc.Name &amp; """" &amp; _ 
 " is not in a protected view window." 
End If
```


## See also


#### Concepts


 [Global Object](b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c.md)
#### Other resources


 [Global Object Members](35050f7b-bc46-4795-ec17-f68e263c8af0.md)
