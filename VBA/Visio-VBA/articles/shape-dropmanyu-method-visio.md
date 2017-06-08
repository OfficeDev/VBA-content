---
title: Shape.DropManyU Method (Visio)
keywords: vis_sdr.chm11251930
f1_keywords:
- vis_sdr.chm11251930
ms.prod: visio
api_name:
- Visio.Shape.DropManyU
ms.assetid: b3e18874-bb90-334f-e633-3e20133242a1
ms.date: 06/08/2017
---


# Shape.DropManyU Method (Visio)

Creates one or more new  **Shape** objects on a page, in a master, or in a group. It returns an array of the IDs of the **Shape** objects it produces.


## Syntax

 _expression_ . **DropManyU**( **_ObjectsToInstance()_** , **_xyArray()_** , **_IDArray()_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectsToInstance()_|Required| **Variant**|Identifies masters or other objects from which to make shapes by their universal names.|
| _xyArray()_|Required| **Double**|An array of alternating  _x_ and _y_ values specifying the positions for the new shapes.|
| _IDArray()_|Required| **Integer**|Out parameter. An array that returns the IDs of the created shapes.|

### Return Value

Integer


## Remarks

Using the  **DropManyU** method is like using the **Page** , **Master** , or **Shape** object's **Drop** method, except you can use the **DropManyU** method to create many new **Shape** objects at once, rather than one per method call. The **DropManyU** method creates new **Shape** objects on the page, in the master, or in the group shape to which it is applied (this shape is called the "target object" in the following discussion).

You can identify which master to drop by passing the  **DropManyU** method a **Master** object or the master's index or the master's name. When you pass an object, **DropManyU** isn't constrained to just dropping a master from the document stencil of the document onto which it is being dropped. The object can be a master from another document or another type of object.

Passing integers (master indices) or strings (master names) to  **DropManyU** is faster than passing objects, but integers or strings can identify only masters in the document stencil of the document onto which it is being dropped. Hence your program has to somehow get the masters in question into the document stencil in the first place, provided they weren't there already.

 _ObjectsToInstance()_ should be a one-dimensional array of _n_ >= 1 variants. Its entries identify objects from which you want to make new **Shape** objects. An entry often refers to a Microsoft Visio application **Master** object. It might also refer to a Visio application **Shape** object, **Selection** object, or even an object from another application. The application doesn't care what the lower and upper array bounds of the _ObjectsToInstance()_ entries are. Call these _vlb_ and _vub_ , respectively.




- If  _ObjectsToInstance(i)_ is the integer _j_ , an instance of the **Master** object in the document stencil of the target object's document whose 1-based index is _j_ is made. The EventDrop cell in the Events section of the new shape is not triggered. Use the **Drop** method instead if you want the EventDrop cell to trigger.
    
- If  _ObjectsToInstance(i)_ is the string _s_ (or a reference to the string _s_ ), an instance of the **Master** object with name _s_ in the document stencil of the target object's document is made; _s_ can equal either the **Master** object's **UniqueID** or **NameU** property. The EventDrop cell in the Events section of the new shape is not triggered. Use the **Drop** method instead if you want the EventDrop cell to be triggered.
    
- For  _vlb_ < _i_ <= _vub_ , if _ObjectsToInstance(i)_ is empty ( **Nothing** or uninitialized in Microsoft Visual Basic), entry _i_ will cause _ObjectsToInstance(j)_ to be instanced again, where _j_ is the largest value < _i_ such that _ObjectsToInstance(j)_ isn't empty. If you want to make _n_ instances of the same thing, only _ObjectsToInstance(vlb)_ needs to be provided.
    


The  _xyArray()_ argument should be a one-dimensional array of 2 _m_ doubles with lower bound _xylb_ and upper bound _xyub_ , where _m_ >= _n_ . The values in the array tell the **DropManyU** method where to position the **Shape** objects it produces. _ObjectsToInstance_( _vlb_ + ( _ i_ - 1)) is dropped at ( _xy_ [( _i_ - 1)2 + _xylb_ ], _xy_ [(i - 1)2 + _xylb_ + 1]) for 1 <= _i_ <= _n_ .

Note that  _m_ > _n_ is allowed. For _n_ < _i_ <= _m_ , the _i_ 'th thing instanced is the same thing as the _n_ 'th thing instanced. Thus to make _m_ >= 1 instances of the same thing, you can pass an _ObjectsToInstance()_ array with one entry and an _m_ entry _xyArray()_ array.

If the entity being instanced is a master, the pin of the new  **Shape** object is positioned at the given _xy_ . Otherwise, the center of the **Shape** objects is positioned at the given _xy_ .

The  **Integer** value returned by the **DropManyU** method is the number of _xy_ entries in _xyArray()_ that the **DropManyU** method successfully processed. If all entries were processed successfully, _m_ is returned. If some entries are successfully processed prior to an error occurring, the produced **Shape** objects are not deleted and this raises an exception but still returns a positive integer.

Presuming all  _m_ _xy_ entries are processed correctly, the number of new **Shape** objects produced by the **DropManyU** method is usually equal to _m_ . In rare cases (for example, if a **Selection** object gets instanced), more than _m_**Shape** objects may be produced. The caller can determine the number of produced **Shape** objects by comparing the number of shapes in the target object before and after the **DropManyU** method is executed. The caller can assert the new **Shape** objects are those with the highest indices in the target object's **Shapes** collection.

If the  **DropManyU** method returns zero (0), _IDArray()_ returns null ( **Nothing** ). Otherwise, it returns a one-dimensional array of _m_ integers indexed from 0 to _m_ - 1. _IDArray()_ is an out parameter that is allocated by the **DropManyU** method and ownership is passed to the program that called the **DropManyU** method. The caller should eventually perform the **SafeArrayDestroy** procedure on the returned array. (Microsoft Visual Basic and Microsoft Visual Basic for Applications take care of this for you.)

If  _IDArray()_ returns non-null (not **Nothing** ), _IDArray_( _i_ - 1), 1 <= _i_ <= _intReturned_ , returns the ID of the **Shape** object produced by the _i_ 'th _xyArray()_ entry, provided the _i_ 'th _xyArray()_ entry produced exactly one **Shape** object. If the _i_ 'th _xyArray()_ entry produced multiple **Shape** objects, -1 is returned in the entry. All entries _i_ , _intReturned_ <= _i_ < _m_ , return -1.




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **DropMany** method to drop more than one shape when you are using local names to identify the shapes. Use the **DropManyU** method to drop more than one shape when you are using universal names to identify the shapes.


## Example

The following example shows how to use the  **DropManyU** method. It drops one instance of every master in the document stencil of the macro's **Document** object onto Page1 of the macro's **Document** object. Before running this macro, make sure there is at least one master in the document stencil.


```vb
 
Public Sub DropManyU_Example() 
 
 On Error GoTo HandleError 
 
 Dim vsoMasters As Visio.Masters 
 Dim intMasterCount As Integer 
 Set vsoMasters = ThisDocument.Masters 
 intMasterCount = vsoMasters.Count 
 
 ReDim varObjectsToInstance(1 To intMasterCount) As Variant 
 ReDim adblXYArray(1 To (intMasterCount * 2)) As Double 
 Dim intCounter As Integer 
 For intCounter = 1 To intMasterCount 
 
 'Pass universal name of object to drop to DropManyU. 
 varObjectsToInstance(intCounter) = vsoMasters.ItemU(intCounter).NameU 
 
 'Set x components of where to drop to 2,4,6,2,4,6,2,4,6,... 
 adblXYArray (intCounter * 2 - 1) = (((intCounter - 1) Mod 3) + 1) * 2 
 
 'Set y components to 2,2,2,4,4,4,6,6,6,... 
 adblXYArray (intCounter * 2) = Int((intCounter + 2) / 3) * 2 
 
 Next intCounter 
 
 Dim aintIDArray() As Integer 
 Dim intProcessed As Integer 
 
 intProcessed = ThisDocument.Pages(1).DropManyU(varObjectsToInstance, _ 
 adblXYArray, aintIDArray) 
 Debug.Print intProcessed 
 
 For intCounter = LBound(aintIDArray) To UBound(aintIDArray) 
 Debug.Print intCounter; aintIDArray(intCounter) 
 Next intCounter 
 
 Exit Sub 
 
 HandleError: 
 MsgBox "Error" 
 
 Exit Sub 
 
End Sub
```


