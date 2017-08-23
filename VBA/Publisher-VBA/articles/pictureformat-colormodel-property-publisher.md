---
title: "Свойство PictureFormat.ColorModel (издатель)"
keywords: vbapb10.chm3604753
f1_keywords: vbapb10.chm3604753
ms.prod: publisher
api_name: Publisher.PictureFormat.ColorModel
ms.assetid: 8e3e259c-943d-c1a9-f090-2ee0f0bb29f2
ms.date: 06/08/2017
ms.openlocfilehash: 2cf3e6a60bf47122183eb6143f06b0586e3c6a3e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatcolormodel-property-publisher"></a>Свойство PictureFormat.ColorModel (издатель)

Возвращает константу **PbColorModel** , представляющий модель цвет рисунка. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ColorModel**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbColorModel


## <a name="remarks"></a>Заметки

Значение свойства **ColorModel** может иметь одно из **[PbColorModel](pbcolormodel-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

Следующий пример возвращает список изображений с режимом цвета RGB в активной публикации.


```vb
Sub ListRGBPictures() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
 For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 If .ColorModel = pbColorModelRGB Then 
 Debug.Print .Filename 
 End If 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
 Next pgLoop 
 
End Sub
```


