
# OLEFormat.Verb Method (Excel)

 **Last modified:** July 28, 2015

Sends a verb to the server of the specified OLE object.

## Syntax

 _expression_. **Verb**( **_Verb_**)

 _expression_A variable that represents an  **OLEFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Verb|Optional| ** [XlOLEVerb](56664ebc-d745-2279-3f6e-b4fdbc6f599a.md)**|The verb that the server of the OLE object should act on. If this argument is omitted, the default verb is sent. The available verbs are determined by the object's source application. Typical verbs for an OLE object are Open and Primary (represented by the  **XlOLEVerb** constants **xlOpen** and **xlPrimary**).|

## See also


#### Concepts


 [OLEFormat Object](96ee06d8-e922-c48c-4406-bb2f5cbaa02a.md)
#### Other resources


 [OLEFormat Object Members](18f0bbed-752a-5e01-51f1-c17435b3adea.md)
