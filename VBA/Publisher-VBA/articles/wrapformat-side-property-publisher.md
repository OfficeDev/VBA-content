---
title: "Свойство WrapFormat.Side (издатель)"
keywords: vbapb10.chm786436
f1_keywords: vbapb10.chm786436
ms.prod: publisher
api_name: Publisher.WrapFormat.Side
ms.assetid: b7998643-216a-a294-bbee-e5f1947400a7
ms.date: 06/08/2017
ms.openlocfilehash: d52aa615c2bd35eccc44c7c831cc66af3f2b7e69
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wrapformatside-property-publisher"></a>Свойство WrapFormat.Side (издатель)

Возвращает или задает значение константы **PbWrapSideType** , которое указывает, следует ли переносить текст вокруг фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Со стороны**

 переменная _expression_A, представляет собой объект- **WrapFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbWrapSideType


## <a name="remarks"></a>Заметки

Значение свойства **со стороны** может иметь одно из **[PbWrapSideType](pbwrapsidetype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере добавляется овала для первой страницы публикации, активных и указывает, что текст окружающей слева и справа овала.


```vb
Sub SetTextWrapFormatProperties() 
 With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeOval, _ 
 Left:=36, Top:=36, Width:=100, Height:=35) 
 With .TextWrap 
 .Type = pbWrapTypeSquare 
 .Side = pbWrapSideBoth 
 End With 
 End With 
End Sub
```


