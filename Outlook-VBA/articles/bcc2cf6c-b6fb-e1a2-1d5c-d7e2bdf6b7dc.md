
# Items Members (Outlook)
Contains a collection of  [Outlook item objects](6ea4babf-facf-4018-ef5a-4a484e55153a.md) in a folder.

 **Last modified:** July 28, 2015

 **In this article**
 [Events](#sectionSection0)
 [Methods](#sectionSection1)
 [Properties](#sectionSection2)


## Events
<a name="sectionSection0"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [ItemAdd](e46f5958-aff8-3a6b-b3df-5c4352b6c3d9.md)|Occurs when one or more items are added to the specified collection. This event does not run when a large number of items are added to the folder at once. This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).|
| [ItemChange](6478357e-2a5a-300a-24e6-c125f8c81edd.md)|Occurs when an item in the specified collection is changed. This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).|
| [ItemRemove](c1b2d9cd-ab32-2c4a-85fa-9412c190ac4f.md)|Occurs when an item is deleted from the specified collection.|

## Methods
<a name="sectionSection1"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [Add](0ee68068-1452-0f29-b85a-88b801ac0448.md)|Creates a new Outlook item in the  ** [Items](3a99730b-e62a-5ca6-f6ec-911c95173242.md)** collection for the folder.|
| [Find](e7a791d8-b80b-df07-84a3-a85acabfcf80.md)|Locates and returns a Microsoft Outlook item object that satisfies the given  _Filter_.|
| [FindNext](2530f640-e024-3567-f539-6bdbf645401d.md)|After the  ** [Find](e7a791d8-b80b-df07-84a3-a85acabfcf80.md)** method runs, this method finds and returns the next Outlook item in the specified collection.|
| [GetFirst](142a6174-118e-6256-0511-8ae9e142e555.md)|Returns the first object in the collection. |
| [GetLast](d02a20be-19fc-fb6e-feff-b66ca0273beb.md)|Returns the last object in the collection. |
| [GetNext](01c49c21-d9f9-37c4-8c64-ff8e2b1f9462.md)|Returns the next object in the collection. |
| [GetPrevious](5dde47f8-2bd8-fdbe-d6e7-b1381e8a97a6.md)|Returns the previous object in the collection. |
| [Item](89a031e0-c0a3-fc22-f485-189df8db45f4.md)|Returns an Outlook item from a collection.|
| [Remove](d2838c82-d0ac-82cc-eed0-c34d55c67d63.md)|Removes an object from the collection.|
| [ResetColumns](0543dd17-1e65-5484-ab21-d4791b3b1194.md)|Clears the properties that have been cached with the  ** [SetColumns](90206a68-baf8-282c-5793-fee029fed452.md)** method.|
| [Restrict](e3b0cda1-e43d-cc5e-2942-0f54935d9dab.md)|Applies a filter to the  ** [Items](3a99730b-e62a-5ca6-f6ec-911c95173242.md)** collection, returning a new collection containing all of the items from the original that match the filter.|
| [SetColumns](90206a68-baf8-282c-5793-fee029fed452.md)|Caches certain properties for extremely fast access to those particular properties of each item in an  ** [Items](3a99730b-e62a-5ca6-f6ec-911c95173242.md)** collection.|
| [Sort](7cb248a2-6885-8be5-df7b-fd5683081e01.md)|Sorts the collection of items by the specified property. The index for the collection is reset to 1 upon completion of this method.|

## Properties
<a name="sectionSection2"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [Application](b55a6901-fbd4-36a1-47e7-2c1e37e0a31c.md)|Returns an  ** [Application](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)**object that represents the parent Outlook application for the object. Read-only.|
| [Class](783ed46a-fd40-c848-b440-8ea3c5d0e6b9.md)|Returns an  ** [OlObjectClass](33d724b3-df3c-2a7f-a80f-93b66d96f588.md)** constant indicating the object's class. Read-only.|
| [Count](c18b06be-3a21-3350-6d14-57c822a85d42.md)|Returns a  **Long** indicating the count of objects in the specified collection. Read-only.|
| [IncludeRecurrences](7d192112-889c-56ce-aab2-107d751c80c4.md)|Returns a  **Boolean** that indicates **True** if the ** [Items](3a99730b-e62a-5ca6-f6ec-911c95173242.md)** collection should include recurrence patterns. Read/write.|
| [Parent](8e99782a-5654-ae1d-c6d8-9dbfcbcf44ac.md)|Returns the parent  **Object** of the specified object. Read-only.|
| [Session](5c385dfc-042e-7649-0f32-5d34e53fca57.md)|Returns the  ** [NameSpace](f0dcaa19-07f5-5d42-a3bf-2e42b7885644.md)**object for the current session. Read-only.|
