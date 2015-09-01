
# Application.ProtectedViewWindowActivate Event (Word)

 **Last modified:** July 28, 2015

Occurs when any protected view window is activated.

## Syntax

 _expression_. **ProtectedViewWindowActivate**( **_PvWindow_**, )

 _expression_An expression that returns a  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PvWindow|Required| ** [ProtectedViewWindow](d77e80e7-c54e-5954-1586-dacd3c9f7434.md)**|The protected view window that is activated.|

## Example

The following code example maximizes any protected view window when it is activated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work. For more information about how to do this, see  [Using Events with the Application Object](784c4c61-7e47-3dbf-46f6-da655f786ca1.md).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```
Private Sub App_ProtectedViewWindowActivate(ByVal PvWindow As ProtectedViewWindow) 
 PvWindow.WindowState = wdWindowStateMaximize 
End Sub
```


## See also


#### Concepts


 [Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


 [Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
