
# Application.Version Property (Publisher)

 **Last modified:** July 28, 2015

Returns a  **String** indicating the version number of the currently-installed copy of Microsoft Publisher. Read-only.

## Syntax

 _expression_. **Version**

 _expression_A variable that represents a  **Application** object.


### Return Value

String


## Example

The following example displays the version and build number of the currently-installed copy of Publisher.


```
MsgBox "You are currently running Microsoft Publisher, " _ 
 &amp; " version " &amp; Application.Version &amp; ", build " _ 
 &amp; Application.Build &amp; "." 

```


## See also


#### Concepts


 [Application Object](acfc7efb-e6a5-a89a-3aee-3cb4af2f3508.md)
#### Other resources


 [Application Object Members](aa4d515b-f779-b8b5-968a-8e5f7466fb56.md)
