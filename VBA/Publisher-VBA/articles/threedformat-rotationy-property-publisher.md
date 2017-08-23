---
title: "Свойство ThreeDFormat.RotationY (издатель)"
keywords: vbapb10.chm3801360
f1_keywords: vbapb10.chm3801360
ms.prod: publisher
api_name: Publisher.ThreeDFormat.RotationY
ms.assetid: 571f090b-71a8-c92e-b4d8-4f21a4c383ed
ms.date: 06/08/2017
ms.openlocfilehash: b8f84c10d33508254cb56744e2cebe00c097e781
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatrotationy-property-publisher"></a>Свойство ThreeDFormat.RotationY (издатель)

Возвращает или задает поворот вытянутый фигуры относительно оси y в градусов. Может быть в диапазоне от - 90 до 90. Положительное значение указывает вращение слева; отрицательное значение указывает вращение справа. Чтение и запись **одного**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RotationY**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Чтобы задать поворот вытянутый фигуры относительно оси x, используйте свойство **[RotationX](threedformat-rotationx-property-publisher.md)** объекта **ThreeDFormat** . Чтобы задать поворот вытянутый фигуры относительно оси z, используйте свойство **[Вращение](shape-rotation-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** . Чтобы изменить направление придания объема конечный путь без вращающимся лицевой изменяется, используйте метод **[SetExtrusionDirection](threedformat-setextrusiondirection-method-publisher.md)** .


## <a name="example"></a>Пример

В этом примере добавляется три идентичные вытянутый овалов в активный документ и устанавливает их поворот вокруг оси y - 30, 0 до 30 градусов, соответственно.


```vb
Sub SetRotationY() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddShape(Type:=msoShapeOval, Left:=30, _ 
 Top:=120, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationY = -30 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=90, _ 
 Top:=120, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationY = 0 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=150, _ 
 Top:=120, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationY = 30 
 End With 
 End With 
End Sub
```


