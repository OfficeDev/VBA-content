---
title: "Свойство ThreeDFormat.ExtrusionColor (издатель)"
keywords: vbapb10.chm3801345
f1_keywords: vbapb10.chm3801345
ms.prod: publisher
api_name: Publisher.ThreeDFormat.ExtrusionColor
ms.assetid: 209a47fd-a219-9533-1a4a-572dfa4312f2
ms.date: 06/08/2017
ms.openlocfilehash: 62c93198ebac7ab74688ba39a1927efe32a75fcb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="threedformatextrusioncolor-property-publisher"></a>Свойство ThreeDFormat.ExtrusionColor (издатель)

Возвращает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющее цвет придания объема фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ExtrusionColor**

 переменная _expression_A, представляющий объект **ThreeDFormat** .


### <a name="return-value"></a>Возвращаемое значение

ColorFormat


## <a name="example"></a>Пример

В этом примере добавляется овала active публикации и затем указывает, что овала быть вытянутый глубина 50 точек и выбирать должен быть фиолетовым.


```vb
Dim shpNew As Shape 
 
' Set a reference to a new oval. 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=90, Top:=90, Width:=90, Height:=40) 
 
' Format the 3-D properties of the oval. 
With shpNew.ThreeD 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With 

```


