
# Application.ProjectBeforeResourceDelete Event (Project)

 **Last modified:** July 28, 2015

Occurs before a resource is deleted.

## Syntax

 _expression_. **ProjectBeforeResourceDelete**( **_res_**,  **_Cancel_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|res|Required| **Resource**| The resource that is being deleted.|
|Cancel|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the resource is not deleted.|

### Return Value

nothing


## Remarks

Project events do not occur when the project is embedded in another document or application.

The  **ProjectBeforeResourceDelete** event doesn't occur when changes have been made using a custom form.

