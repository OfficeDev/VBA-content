---
title: "Метод Document.FindShapesByTag (издатель)"
keywords: vbapb10.chm196689
f1_keywords: vbapb10.chm196689
ms.prod: publisher
api_name: Publisher.Document.FindShapesByTag
ms.assetid: 405a0f39-5892-23da-904a-5188a4340b00
ms.date: 06/08/2017
ms.openlocfilehash: 0b4ea229a08d6d063fadd018fdca05073d02dc65
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentfindshapesbytag-method-publisher"></a>Метод Document.FindShapesByTag (издатель)

Возвращает объект **[ShapeRange](shaperange-object-publisher.md)** , представляющий фигур с помощью указанного тега.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FindShapesByTag** ( **_TagName_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|TagName|Обязательное свойство.| **String**|Имя тега.|

### <a name="return-value"></a>Возвращаемое значение

ShapeRange


## <a name="example"></a>Пример

В этом примере добавляет две фигуры в первой страницы публикации, active, назначает каждого тега и затем вводит имя каждого тега в рамку ее назначенный фигуры.


```vb
Sub FindShape() 
 Dim strTag1 As String 
 Dim strTag2 As String 
 
 With ActiveDocument.Pages(1).Shapes 
 With .AddShape(Type:=msoShape5pointStar, Left:=50, _ 
 Top:=50, Width:=75, Height:=75) 
 strTag1 = .Tags.Add(Name:="Star", _ 
 Value:="This is a star.").Name 
 End With 
 
 With .AddShape(Type:=msoShapeHeart, Left:=100, _ 
 Top:=100, Width:=75, Height:=75) 
 strTag2 = .Tags.Add(Name:="Heart", _ 
 Value:="This is a heart.").Name 
 End With 
 End With 
 
 With ActiveDocument 
 .FindShapesByTag(TagName:=strTag1).TextFrame _ 
 .TextRange.Text = strTag1 
 .FindShapesByTag(TagName:=strTag2).TextFrame _ 
 .TextRange.Text = strTag2 
 End With 
End Sub
```


