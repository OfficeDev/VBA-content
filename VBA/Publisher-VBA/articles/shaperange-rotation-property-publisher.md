---
title: "Свойство ShapeRange.Rotation (издатель)"
keywords: vbapb10.chm2293830
f1_keywords: vbapb10.chm2293830
ms.prod: publisher
api_name: Publisher.ShapeRange.Rotation
ms.assetid: 0239aaae-18c7-56ef-f2b1-82f82660370a
ms.date: 06/08/2017
ms.openlocfilehash: 5dc75e8ef22551ccbb56948311e9ec91ba5e4730
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangerotation-property-publisher"></a>Свойство ShapeRange.Rotation (издатель)

Возвращает или задает **единого** , представляющее номер указанный указанные форму — это вращаться вокруг оси z. Положительное значение указывает часовой стрелки; отрицательное значение указывает против вращение. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вращение**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


