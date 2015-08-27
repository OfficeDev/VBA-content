
# MasterShortcut.ImportIcon Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Imports the icon for a  **Master** object from a named file.


## Syntax

 _expression_. **ImportIcon**( **_FileName_**)

 _expression_A variable that represents a  **MasterShortcut** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|FileName|Required| **String**|The name of the file to import.|

### Return Value

Nothing


## Remarks

The  **ImportIcon** method can only import files that were produced by exporting a master icon in the application's internal icon format ( **visIconFormatVisio**)â€”it does not accept icons in other file formats.

