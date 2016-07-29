
# Application.Version Property (Publisher)

Returns a  **String** indicating the version number of the currently-installed copy of Microsoft Publisher. Read-only.


## Syntax

 _expression_. **Version**

 _expression_A variable that represents a  **Application** object.


### Return Value

String


## Example

The following example displays the version and build number of the currently-installed copy of Publisher.


```vb
MsgBox "You are currently running Microsoft Publisher, " _ 
 &; " version " &; Application.Version &; ", build " _ 
 &; Application.Build &; "." 

```


## See also


#### Concepts


 [Application Object](acfc7efb-e6a5-a89a-3aee-3cb4af2f3508.md)
#### Other resources


 [Application Object Members](aa4d515b-f779-b8b5-968a-8e5f7466fb56.md)
