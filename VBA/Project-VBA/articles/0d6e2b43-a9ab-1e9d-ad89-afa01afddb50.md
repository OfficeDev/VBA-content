
# Application.DrawingMove Method (Project)

Moves the active drawing object forward or backward in the drawing layers.


## Syntax

 _expression_. **DrawingMove**( ** _Forward_**, ** _Full_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Forward_|Optional|**Boolean**|**True** if the active drawing object moves forward in the drawing layers. The default value is **False**.|
| _Full_|Optional|**Boolean**|**True** if the active drawing object moves the full extent of the direction specified with **Forward**. **False** if the object moves only one layer. The default value is **False**.|

### Return Value

 **Boolean**

