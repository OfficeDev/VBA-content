---
title: "Свойство ThreeDFormat.RotationX (издатель)"
keywords: vbapb10.chm3801353
f1_keywords: vbapb10.chm3801353
ms.prod: publisher
api_name: Publisher.ThreeDFormat.RotationX
ms.assetid: 1ee394cb-746b-02f0-f2af-aa4a6fffd172
ms.date: 06/08/2017
ms.openlocfilehash: 5f49eb209157abb46b99e8797eb55fb3f327b3a6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatrotationx-property-publisher"></a>Свойство ThreeDFormat.RotationX (издатель)

Возвращает или задает поворот вытянутый фигуры относительно оси x в степени. Может быть в диапазоне от - 90 до 90. Положительное значение указывает вверх цикл; отрицательное значение указывает вращение вниз. Чтение и запись **одного**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RotationX**

 переменная _expression_A, представляет собой объект- **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Чтобы задать поворот вытянутый фигуры относительно оси y, используйте свойство **[RotationY](threedformat-rotationy-property-publisher.md)** объекта **ThreeDFormat** . Чтобы задать поворот вытянутый фигуры относительно оси z, используйте свойство **[Вращение](shape-rotation-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** . Чтобы изменить направление придания объема конечный путь без вращающимся лицевой изменяется, используйте метод **[SetExtrusionDirection](threedformat-setextrusiondirection-method-publisher.md)** .


## <a name="example"></a>Пример

В этом примере добавляется три идентичные вытянутый овалов в активный документ и устанавливает их поворот вокруг оси x - 30, 0 до 30 градусов, соответственно.


```vb
Sub SetRotationX() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddShape(Type:=msoShapeOval, Left:=30, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = -30 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=90, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = 0 
 End With 
 With .AddShape(Type:=msoShapeOval, Left:=150, _ 
 Top:=60, Width:=50, Height:=25).ThreeD 
 .Visible = True 
 .RotationX = 30 
 End With 
 End With 
End Sub
```


