---
title: Page.DrawNURBS Method (Visio)
keywords: vis_sdr.chm10916205
f1_keywords:
- vis_sdr.chm10916205
ms.prod: visio
api_name:
- Visio.Page.DrawNURBS
ms.assetid: f3c7e6fe-71a4-4809-b60a-a34cebd737b1
ms.date: 06/08/2017
---


# Page.DrawNURBS Method (Visio)

Creates a new shape whose path consists of a single NURBS (nonuniform rational B-spline) segment.


## Syntax

 _expression_ . **DrawNURBS**( **_degree_** , **_Flags_** , **_xyArray()_** , **_knots()_** , **_weights_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _degree_|Required| **Integer**|The spline's degree; an integer between 1 and 25.|
| _Flags_|Required| **Integer**|Flags that influence how the shape is drawn.|
| _xyArray()_|Required| **Double**|An array of alternating  _x_ and _y_ values that define the control points coordinates; use internal drawing units (inches).|
| _knots()_|Required| **Double**|An array of knots.|
| _weights_|Optional| **Variant**|An array of weights.|

### Return Value

Shape


## Remarks

The  **DrawNURBS** method creates a new shape whose path consists of a single NURBS segment as specified by the arguments.

The control points should be in internal drawing units (inches) with respect to the coordinate space of the page, master, or group in which the new shape is being created. The  _xyArray_,  _knots_, and  _weights_ arrays should be of type SAFEARRAY of 8-byte floating point values passed by reference (VT_R8|VT_ARRAY|VT_BYREF). This is how Microsoft Visual Basic passes arrays to Automation objects.

The  _knots_ argument is unit-less. The sequence of _knots_ should be non-decreasing. In other words, _knots_( _i_ + 1) < _knots_( _i_ ) is not acceptable. _knots_( _i_ + 1) = _knots_( _i_ ) is permitted, and then the value is repeated, but the following restrictions apply:




- The first knot may not be repeated more than  _degree_ + 1 times.
    
- The last knot may not be repeated.
    
- Any knot between the first and last may not be repeated more than  _degree_ times.
    
- If the first knot is repeated less than  _degree_ + 1 times, the spline is _periodic_ .
    
- The list of weights is optional. Its absence signals that the spline is  _non-rational_ . Weights are unit-less.
    


The following rules apply to the sizes of the lists. For a spline with n control points:




- If the spline is periodic,  _n_ > 2. Otherwise, _n_ > _degree_.
    
- The size of  _xyArray_ is 2 _n_ .
    
- The size of the  _weights_ array is _n_ (if present).
    
- The size of the  _knots_ array is _n_ + 1.
    


The conventional non-periodic spline requires  _n_ + _degree_ + 1 _knots_, but the application implies the repeated  _knots_ at the end. For example, the _degree_ 2 knot list (0,0,0,2,5,8) is interpreted in the application as the conventional knot sequence (0,0,0,2,5,8,8,8).

The  _Flags_ parameter is a bitmask that specifies options for drawing the new shape. Its value should be either zero (0) or **visSpline1D** (8). If _Flags_ is **visSpline1D** and if the first and last points in _xyArray_ don't coincide, the **DrawNURBS** method produces a shape with one-dimensional (1-D) behavior; otherwise, it produces a shape with two-dimensional (2-D) behavior.

If the first and last points in  _xyArray_ do coincide, the **DrawNURBS** method produces a filled shape.


