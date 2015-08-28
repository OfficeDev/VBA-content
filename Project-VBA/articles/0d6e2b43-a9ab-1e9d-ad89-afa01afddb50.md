
# Application.DrawingMove Method (Project)

 **Last modified:** July 28, 2015

Moves the active drawing object forward or backward in the drawing layers.

## Syntax

 _expression_. **DrawingMove**( **_Forward_**,  **_Full_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Forward|Optional| **Boolean**| **True** if the active drawing object moves forward in the drawing layers. The default value is **False**.|
|Full|Optional| **Boolean**| **True** if the active drawing object moves the full extent of the direction specified with **Forward**.  **False** if the object moves only one layer. The default value is **False**.|

### Return Value

 **Boolean**

