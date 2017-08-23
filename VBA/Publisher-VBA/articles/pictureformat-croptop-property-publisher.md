---
title: "Свойство PictureFormat.CropTop (издатель)"
keywords: vbapb10.chm3604742
f1_keywords: vbapb10.chm3604742
ms.prod: publisher
api_name: Publisher.PictureFormat.CropTop
ms.assetid: b235898d-addf-6a4c-5693-229431545e6c
ms.date: 06/08/2017
ms.openlocfilehash: 871b8d3acdadfd396fc998a944ca70b086b41909
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatcroptop-property-publisher"></a>Свойство PictureFormat.CropTop (издатель)

Возвращает или задает **Variant** , показывающее, с помощью которого обрезается верхнего края рисунка или объекта OLE. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CropTop**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Отрицательные значения обрезки верхнего края от центра фрейма и положительные значения обрезать к границам снизу кадра.

Диапазон допустимых значений обрезки зависит от того, положение и размер кадра. Для исходное frame разрешенных низший отрицательным значением является расстояние от верхнего края frame и верхнего края вспомогательной области. Наибольшее положительное значение разрешено является текущий высота кадра.

Обрезка рассчитывается относительно исходного размера изображения. Например если вставить рисунок, который изначально — 100 точки высокой, размера, чтобы он был 200 точек высокой и свойства **CropTop** 50 100 точек (не 50) будет обрезки off в верхней части рисунка.

Использование свойств **[CropLeft](pictureformat-cropleft-property-publisher.md)**, **[CropRight](pictureformat-cropright-property-publisher.md)**и **[CropBottom](pictureformat-cropbottom-property-publisher.md)** обрезать других края рисунка или объекта OLE.


## <a name="example"></a>Пример

В этом примере Кадрирование 20 точек отключена в верхней части третий фигуры в активной публикации. Для обеспечения работы примера фигуры должен быть изображения или объекта OLE.


```vb
ActiveDocument.Pages(1).Shapes(3).PictureFormat _ 
 .CropTop = 20
```

В этом примере обрезает процент, указанный пользователем off в верхней части выбранной фигуры, независимо от того, был ли увеличен фигуры. Для обеспечения работы примера выбранной фигуры должен быть изображения или объекта OLE.




```vb
Dim sngPercent As Single 
Dim shpCrop As Shape 
Dim sngPoints As Single 
Dim sngHeight As Single 
 
sngPercent = InputBox("What percentage do you " &; _ 
 "want to crop off the top of this picture?") 
 
Set shpCrop = Selection.ShapeRange(1) 
With shpCrop.Duplicate 
 .ScaleHeight Factor:=1, _ 
 RelativeToOriginalSize:=True 
 sngHeight = .Height 
 .Delete 
End With 
 
sngPoints = sngHeight * sngPercent / 100 
 
shpCrop.PictureFormat.CropTop = sngPoints 

```


