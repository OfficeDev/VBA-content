
# Section.FormulaChanged Event (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Occurs after a formula changes in a cell in the object that receives the event.


## Syntax

Private Sub  _expression__**FormulaChanged**( **_ByVal Cell As [IVCELL]_**)

 _expression_A variable that represents a  **Section** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cell|Required| **[IVCELL]**|The cell whose formula changed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](de8f5c7a-421d-ebcf-22b6-4310a202ef64.md).




 **Note**  You can use VBA  **WithEvents** variables to sink the **FormulaChanged** event.

For performance considerations, the  **Document** object's event set does not include the **FormulaChanged** event. To sink the **FormulaChanged** event from a **Document** object (and the **ThisDocument** object in a VBA project), you must use the **AddAdvise** method.

