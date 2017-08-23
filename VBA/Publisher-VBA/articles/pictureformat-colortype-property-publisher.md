---
title: "Свойство PictureFormat.ColorType (издатель)"
keywords: vbapb10.chm3604737
f1_keywords: vbapb10.chm3604737
ms.prod: publisher
api_name: Publisher.PictureFormat.ColorType
ms.assetid: 439f9eb9-2593-d719-4ef6-0f14d1c7d0f4
ms.date: 06/08/2017
ms.openlocfilehash: 08996affa9fd1c2bfb7e5599c9650c7591f1eb85
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatcolortype-property-publisher"></a>Свойство PictureFormat.ColorType (издатель)

Возвращает или задает константой **MsoPictureColorType** , указывающее тип преобразования цветов, применяемые к указанной изображения или объекта OLE. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ColorType**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoPictureColorType


## <a name="remarks"></a>Заметки

Значение свойства **ColorType** может иметь одно из ** [MsoPictureColorType](http://msdn.microsoft.com/library/d11f2d08-2ac9-6cf4-34b8-7ffaabb5d4ae%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере задается преобразования цветовые оттенки серого для первой фигуры в активной публикации. Фигура должен быть изображения или объекта OLE.


```vb
ActiveDocument.Pages(1).Shapes(1).PictureFormat _ 
 .ColorType = msoPictureGrayScale
```


