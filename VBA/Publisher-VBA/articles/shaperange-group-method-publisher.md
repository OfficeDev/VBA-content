---
title: "Метод ShapeRange.Group (издатель)"
keywords: vbapb10.chm2294018
f1_keywords: vbapb10.chm2294018
ms.prod: publisher
api_name: Publisher.ShapeRange.Group
ms.assetid: ca3e011f-72ea-904e-da3f-cac7fe24341d
ms.date: 06/08/2017
ms.openlocfilehash: d79f3071ddde6ed8811335f7403c7f16b4a1edb6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangegroup-method-publisher"></a>Метод ShapeRange.Group (издатель)

Группы фигур в диапазон указанной фигуры. Возвращает группы фигур в виде одного объекта **[Shape](shape-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Группа**

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Указанный диапазон должен содержать более одного фигуры или возникает ошибка.

Так как в группы фигур обрабатывается как одну форму, Группировка и разгруппировка фигур изменения количество элементов в коллекции **[фигур](shapes-object-publisher.md)** и изменяет номера индекса элементов, следующие за затронутых элементов в коллекции.


## <a name="example"></a>Пример

В этом примере добавляет две фигуры в первой страницы публикации, active, группирует две новые фигуры, задает заполнения группы, поворот группы и отправляет группы на задней слой графики.


```vb
With ActiveDocument.Pages(1).Shapes 
 
 ' Add two shapes to the page. 
 .AddShape(Type:=msoShapeCan, _ 
 Left:=50, Top:=10, Width:=100, Height:=200).Name = "shpOne" 
 .AddShape(Type:=msoShapeCube, _ 
 Left:=150, Top:=250, Width:=100, Height:=200).Name = "shpTwo" 
 
 ' Group the shapes and change the formatting for the whole group. 
 With .Range(Index:=Array("shpOne", "shpTwo")).Group 
 .Fill.PresetTextured PresetTexture:=msoTextureBlueTissuePaper 
 .Rotation = 45 
 .ZOrder ZOrderCmd:=msoSendToBack 
 End With 
 
End With 

```


