---
title: "Метод ThreeDFormat.ResetRotation (издатель)"
keywords: vbapb10.chm3801106
f1_keywords: vbapb10.chm3801106
ms.prod: publisher
api_name: Publisher.ThreeDFormat.ResetRotation
ms.assetid: 91e3943a-0087-fcb9-e33f-d41b60b869a7
ms.date: 06/08/2017
ms.openlocfilehash: 4148aaff4aaff328a46875649dab70bf0c91a83a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatresetrotation-method-publisher"></a>Метод ThreeDFormat.ResetRotation (издатель)

Сбрасывает придания объема поворот вокруг оси x (по горизонтали) и y (по вертикали) нуль (0), чтобы переадресовывать сталкивается передней части изменяется.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ResetRotation**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


## <a name="remarks"></a>Заметки

Этот метод не Отменить поворот относительно оси z (расширяет наружу плоскости публикации).

Чтобы задать придания объема поворот вокруг оси x и y, отличное от 0, используйте свойства **[RotationX](threedformat-rotationx-property-publisher.md)** и **[RotationY](threedformat-rotationy-property-publisher.md)** объекта **ThreeDFormat** .

Чтобы установить для придания объема поворот вокруг оси z, используйте свойство **[Вращение](shape-rotation-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** , представляющий вытянутый фигуры.


## <a name="example"></a>Пример

В этом примере восстанавливаются значения по умолчанию поворот вокруг оси x и y нуль для придания объема первой фигуры в активной публикации.


```vb
ActiveDocument.Pages(1).Shapes(1).ThreeD _ 
 .ResetRotation
```


