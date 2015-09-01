
# AddressEntry.Update Method (Outlook)

 **Last modified:** July 28, 2015

Posts a change to the  ** [AddressEntry](d4a0a85e-8bab-bc56-57bc-d70c3c570c8e.md)** object in the messaging system.

## Syntax

 _expression_. **Update**( **_MakePermanent_**,  **_Refresh_**)

 _expression_An expression that returns a  **AddressEntry** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|MakePermanent|Optional| **Variant**| A value of **True** indicates that the property cache is flushed and all changes are committed in the underlying address book. A value of **False** indicates that the property cache is flushed but not committed to persistent storage. The default value is **True**.|
|Refresh|Optional| **Variant**|A value of  **True** indicates that the property cache is reloaded from the values in the underlying address book. A value of **False** indicates that the property cache is not reloaded. The default value is **False**.|

## Remarks

New entries or changes to existing entries are not persisted in the collection until the  **Update** method has been called with itsMakePermanent parameter set to **True**.

To flush the cache and then reload the values from the address book, call  **Update** with theMakePermanent parameter set to **False** and theRefresh parameter set to **True**.


## See also


#### Concepts


 [AddressEntry Object](d4a0a85e-8bab-bc56-57bc-d70c3c570c8e.md)
#### Other resources


 [AddressEntry Object Members](74c88069-aec4-952b-556f-03873fbb488b.md)
