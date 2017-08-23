---
title: "Свойство PictureFormat.CropRight (издатель)"
keywords: vbapb10.chm3604741
f1_keywords: vbapb10.chm3604741
ms.prod: publisher
api_name: Publisher.PictureFormat.CropRight
ms.assetid: b1c20de2-e2cf-708f-ddae-194c8b1b01c1
ms.date: 06/08/2017
ms.openlocfilehash: f1317a99a0cd6b5727d0f2cac21c7baafe569e10
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatcropright-property-publisher"></a>Свойство PictureFormat.CropRight (издатель)

Возвращает или задает **Variant** , показывающее, с помощью которого обрезается правого края рисунка или объекта OLE. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CropRight**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Отрицательные значения обрезки нижнего края от центра фрейма и положительные значения Обрезать по левому краю элемента frame.

Диапазон допустимых значений обрезки зависит от того, положение и размер кадра. Для исходное frame разрешенных низший отрицательным значением является расстояние между правым краем frame и правого края вспомогательной области. Наибольшее положительное значение разрешено является текущий ширина кадра.

Обрезка рассчитывается относительно исходного размера изображения. Например если вставить рисунок, который изначально — 100 точки широкий, размера, чтобы он был 200 точек широкий и свойства **CropRight** 50 100 точек (не 50) будет обрезки off в правой части рисунка.

Использование свойств **[CropLeft](pictureformat-cropleft-property-publisher.md)**, **[CropTop](pictureformat-croptop-property-publisher.md)**и **[CropBottom](pictureformat-cropbottom-property-publisher.md)** обрезать других края рисунка или объекта OLE.


## <a name="example"></a>Пример

В этом примере Кадрирование 20 точек отключена в правой части третий фигуры в активной публикации. Для обеспечения работы примера фигуры должен быть изображения или объекта OLE.


```vb
ActiveDocument.Pages(1).Shapes(3).PictureFormat _ 
 .CropRight = 20
```

В этом примере обрезает процент, указанный пользователем off справа от выбранной фигуры, независимо от того, был ли увеличен фигуры. Для обеспечения работы примера выбранной фигуры должен быть изображения или объекта OLE.




```vb
Dim sngPercent As Single 
Dim shpCrop As Shape 
Dim sngPoints As Single 
Dim sngWidth As Single 
 
sngPercent = InputBox("What percentage do you " &; _ 
 "want to crop off the right of this picture?") 
 
Set shpCrop = Selection.ShapeRange(1) 
With shpCrop.Duplicate 
 .ScaleWidth Factor:=1, _ 
 RelativeToOriginalSize:=True 
 sngWidth = .Width 
 .Delete 
End With 
 
sngPoints = sngWidth * sngPercent / 100 
 
shpCrop.PictureFormat.CropRight = sngPoints 

```


