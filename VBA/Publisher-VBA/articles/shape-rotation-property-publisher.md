---
title: "Свойство Shape.Rotation (издатель)"
keywords: vbapb10.chm2228294
f1_keywords: vbapb10.chm2228294
ms.prod: publisher
api_name: Publisher.Shape.Rotation
ms.assetid: 3cb55e8c-83fa-2f20-caac-a1e897e9a369
ms.date: 06/08/2017
ms.openlocfilehash: f1789cd4dd5aac3bce930d3e3240c0821dd0d5bc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperotation-property-publisher"></a>Свойство Shape.Rotation (издатель)

Возвращает или задает **единого** , представляющее номер указанный указанные форму — это вращаться вокруг оси z. Положительное значение указывает часовой стрелки; отрицательное значение указывает против вращение. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вращение**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Чтобы задать поворот объемной фигуры вокруг оси x или y, используйте свойство **[RotationX](threedformat-rotationx-property-publisher.md)** или свойство **[RotationY](threedformat-rotationy-property-publisher.md)** объекта **[ThreeDFormat](threedformat-object-publisher.md)** .


## <a name="example"></a>Пример

В этом примере сопоставляет вращение всех фигур на первой странице активная публикация цикл первую фигуру. В этом примере предполагается, что имеется по крайней мере двух фигур на первой странице active публикации.


```vb
Sub SetShapeRotation() 
 Dim sngRotation As Single 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes 
 sngRotation = .Item(1).Rotation 
 For intCount = 1 To .Count 
 .Item(intCount).Rotation = sngRotation 
 Next intCount 
 End With 
End Sub
```


