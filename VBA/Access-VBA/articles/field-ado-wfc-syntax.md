---
title: Field (ADO/WFC Syntax)
ms.prod: access
ms.assetid: 61d7028f-ed13-2a20-643d-68d43df91163
ms.date: 06/08/2017
---


# Field (ADO/WFC Syntax)

  

**Applies to:** Access 2013 | Access 2016

 **package com.ms.wfc.data**

 **Methods**



```
 
public void Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthappchunk_HV10294090.xml(byte[] bytes ) 
public void appendChunk(char[] chars ) 
public void appendChunk(String chars ) 
public byte[] getByteChunk(int len ) 
public char[] getCharChunk(int len ) 
public String getStringChunk(int len ) 

```

 **Properties**



```
 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactualsize_HV10293998.xml() 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproattributes_HV10294098.xml() 
public void setAttributes(int pl ) 
public com.ms.com.IUnknown getDataFormat() 
public void setDataFormat(com.ms.com.IUnknown format ) 

```

(For more information, see the Microsoft Visual J++ WFC Reference documentation for the com.ms.wfc.data.IDataFormat interface.)



```
 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprodefinedsize_HV10294289.xml() 
public void setDefinedSize(int pl ) 
public String Invalid DDUE based on source, error:link not allowed in code, link filename:mdproname_HV10294535.xml() 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdpronumericscale_HV10294551.xml() 
public void setNumericScale(byte pbNumericScale ) 
public Variant Invalid DDUE based on source, error:link not allowed in code, link filename:mdprooriginalvalue_HV10294583.xml() 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdproprecision_HV10294615.xml() 
public void setPrecision(byte pbPrecision ) 
public int Invalid DDUE based on source, error:link not allowed in code, link filename:mdprotype_HV10294866.xml() 
public void setType(int pDataType ) 
public Variant Invalid DDUE based on source, error:link not allowed in code, link filename:mdprounderlyingvalue_HV10294879.xml() 
public Variant Invalid DDUE based on source, error:link not allowed in code, link filename:mdprovalue_HV10294920.xml() 
public void setValue(Variant value ) 
public AdoProperties Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolproperties_HV10294633.xml() 

```

 **Field Accessor Methods**
The [Value](http://msdn.microsoft.com/library/ff21d122-98e3-2b48-d92f-e696b8079fc5%28Office.15%29.aspx) property of a[Field](http://msdn.microsoft.com/library/1dbd535e-48ad-a5c8-a1b2-6776c1e3e19d%28Office.15%29.aspx) object gets or sets the content of that object. The content is represented as a VARIANT, a type of object that can be assigned a value and any of several data types.
ADO/WFC implements the  **Value** property with the **getValue** method, which returns a VARIANT object; and the **setValue** method, which takes a VARIANT as an argument. VARIANTs are highly efficient in certain languages, such as Microsoft Visual Basic. However, you can attain better performance in Microsoft Visual J++ by using native Java data types.
In addition to the  **Value** property, ADO/WFC provides _accessor_ methods that use Java data types to get and set the content of **Field** objects. Most of these methods have names of the form **get** _DataType_ or **set** _DataType_.
There are two noteworthy exceptions: One of the  **getObject** methods returns an object coerced into a specified class. There is no **getNull** property; instead, there is an **isNull** property that returns a Boolean value indicating whether the field is null.



```js
 
public native boolean getBoolean(); 
public void setBoolean(boolean v ) 
public native byte getByte(); 
public void setByte(byte v ) 
public native byte[] getBytes(); 
public void setBytes(byte[] v ) 
public native double getDouble(); 
public void setDouble(double v ) 
public native float getFloat(); 
public void setFloat(float v ) 
public native int getInt(); 
public void setInt(int v ) 
public native long getLong(); 
public void setLong(long v ) 
public native short getShort(); 
public void setShort(short v ) 
public native String getString(); 
public void setString(String v ) 
public native boolean isNull(); 
public void setNull() 
public Object getObject() 
public Object getObject(Class c ) 
public void setObject(Object value ) 

```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

