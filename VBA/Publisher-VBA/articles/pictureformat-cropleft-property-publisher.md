---
title: "Свойство PictureFormat.CropLeft (издатель)"
keywords: vbapb10.chm3604740
f1_keywords: vbapb10.chm3604740
ms.prod: publisher
api_name: Publisher.PictureFormat.CropLeft
ms.assetid: f9fd2031-83f7-ea81-84eb-4f1ac6d65082
ms.date: 06/08/2017
ms.openlocfilehash: 273b5f3210f9d219e6939808761fb7275df10c05
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatcropleft-property-publisher"></a>Свойство PictureFormat.CropLeft (издатель)

Возвращает или задает **Variant** , показывающее, с помощью которого обрезается левого края рисунка или объекта OLE. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CropLeft**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Отрицательные значения обрезки нижнего края от центра фрейма и положительные значения Обрезать по правому краю фрейма.

Диапазон допустимых значений обрезки зависит от того, положение и размер кадра. Для исходное frame разрешенных низший отрицательное значение задается расстояние от левого края frame и левого края вспомогательной области. Наибольшее положительное значение разрешено является текущий ширина кадра.

Обрезка рассчитывается относительно исходного размера изображения. Например если вставить рисунок, который изначально — 100 точки широкий, размера, чтобы он был 200 точек широкий и свойства **CropLeft** 50 100 точек (не 50) будет обрезки off в левой части рисунка.

Использование свойств **[CropRight](pictureformat-cropright-property-publisher.md)**, **[CropTop](pictureformat-croptop-property-publisher.md)**и **[CropBottom](pictureformat-cropbottom-property-publisher.md)** обрезать других края рисунка или объекта OLE.


## <a name="example"></a>Пример

В этом примере Кадрирование 20 точек off слева от третьей фигуры в активной публикации. Для обеспечения работы примера фигуры должен быть изображения или объекта OLE.


```vb
ActiveDocument.Pages(1).Shapes(3).PictureFormat _ 
 .CropLeft = 20
```

В этом примере обрезает процент, указанный пользователем off слева от выбранной фигуры, независимо от того, был ли увеличен фигуры. Для обеспечения работы примера выбранной фигуры должен быть изображения или объекта OLE.




```vb
Dim sngPercent As Single 
Dim shpCrop As Shape 
Dim sngPoints As Single 
Dim sngWidth As Single 
 
sngPercent = InputBox("What percentage do you " &; _ 
 "want to crop off the left of this picture?") 
 
Set shpCrop = Selection.ShapeRange(1) 
With shpCrop.Duplicate 
 .ScaleWidth Factor:=1, _ 
 RelativeToOriginalSize:=True 
 sngWidth = .Width 
 .Delete 
End With 
 
sngPoints = sngWidth * sngPercent / 100 
 
shpCrop.PictureFormat.CropLeft = sngPoints 

```


