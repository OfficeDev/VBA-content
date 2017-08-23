---
title: "Свойство ShapeRange.LockAspectRatio (издатель)"
keywords: vbapb10.chm2293827
f1_keywords: vbapb10.chm2293827
ms.prod: publisher
api_name: Publisher.ShapeRange.LockAspectRatio
ms.assetid: 8ed4f41f-3395-dd59-29d4-f66afd19ac51
ms.date: 06/08/2017
ms.openlocfilehash: f3c04244a4ed66c8947eb19af278a3d56d81c2a3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangelockaspectratio-property-publisher"></a>Свойство ShapeRange.LockAspectRatio (издатель)

Возвращает или задает константой **MsoTriState**, указывающее, является ли указанный фигуры сохраняются исходные пропорции при изменении размера. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LockAspectRatio**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Значение свойства **LockAspectRatio** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Высота и ширина формы изменения независимо друг от друга при изменении размера.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Указанный фигуры сохраняются исходные пропорции при изменении размера.|

## <a name="example"></a>Пример

В этом примере добавляется куба active публикацию. Куб можно переместить и размера, но не reproportioned.


```vb
Dim shp As Shape 
 
Set shp = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeCube, _ 
 Left:=50, Top:=50, Width:=100, Height:=200) _ 
 
shp.LockAspectRatio = msoTrue
```


