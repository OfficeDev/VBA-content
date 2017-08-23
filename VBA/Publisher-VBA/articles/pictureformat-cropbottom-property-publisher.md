---
title: "Свойство PictureFormat.CropBottom (издатель)"
keywords: vbapb10.chm3604739
f1_keywords: vbapb10.chm3604739
ms.prod: publisher
api_name: Publisher.PictureFormat.CropBottom
ms.assetid: 8c504221-11da-f6f1-8fbb-75dc5c62b953
ms.date: 06/08/2017
ms.openlocfilehash: 4f6207a6cd0028b4f1fee04db5f6866293b7c785
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatcropbottom-property-publisher"></a>Свойство PictureFormat.CropBottom (издатель)

Возвращает или задает **Variant** , показывающее, с помощью которого обрезается нижнего края рисунка или объекта OLE. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CropBottom**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Отрицательные значения обрезки нижнего края от центра фрейма и положительные значения обрезать направить верхнего края фрейма.

Диапазон допустимых значений обрезки зависит от того, положение и размер кадра. Для исходное frame разрешенных низший отрицательным значением является расстояние между нижнего края frame и нижний край вспомогательной области. Наибольшее положительное значение разрешено является текущий высота кадра.

Обрезка рассчитывается относительно исходного размера изображения. Например если вставить рисунок, который изначально — 100 точки высокой, размера, чтобы он был 200 точек высокой и свойства **CropBottom** 50 100 точек (не 50) будет обрезки в нижней части рисунка.

Использование свойств **[CropLeft](pictureformat-cropleft-property-publisher.md)**, **[CropRight](pictureformat-cropright-property-publisher.md)**и **[CropTop](pictureformat-croptop-property-publisher.md)** обрезать других края рисунка или объекта OLE.


## <a name="example"></a>Пример

В этом примере Кадрирование 20 точек отключена в нижней части третий фигуры в активной публикации. Для обеспечения работы примера фигуры должен быть изображения или объекта OLE.


```vb
ActiveDocument.Pages(1).Shapes(3).PictureFormat _ 
 .CropBottom = 20
```

В этом примере обрезает процент, указанный пользователем в нижней части выбранной фигуры, независимо от того, был ли увеличен фигуры. Для обеспечения работы примера выбранной фигуры должен быть изображения или объекта OLE.




```vb
Dim sngPercent As Single 
Dim shpCrop As Shape 
Dim sngPoints As Single 
Dim sngHeight As Single 
 
sngPercent = InputBox("What percentage do you " &; _ 
 "want to crop off the bottom of this picture?") 
 
Set shpCrop = Selection.ShapeRange(1) 
With shpCrop.Duplicate 
 .ScaleHeight Factor:=1, _ 
 RelativeToOriginalSize:=True 
 sngHeight = .Height 
 .Delete 
End With 
 
sngPoints = sngHeight * sngPercent / 100 
 
shpCrop.PictureFormat.CropBottom = sngPoints 

```


