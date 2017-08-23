---
title: "Свойство WrapFormat.Type (издатель)"
keywords: vbapb10.chm786435
f1_keywords: vbapb10.chm786435
ms.prod: publisher
api_name: Publisher.WrapFormat.Type
ms.assetid: da53302c-ae95-5aa9-a4ce-32647a2569d6
ms.date: 06/08/2017
ms.openlocfilehash: ae860c27abf2b643c8533eda8d98628327efe46b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wrapformattype-property-publisher"></a>Свойство WrapFormat.Type (издатель)

Указывает, обтекания фигуры указанный текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Тип**

 переменная _expression_A, представляет собой объект- **WrapFormat** .


## <a name="remarks"></a>Заметки

Значение свойства **типа** может иметь одно из **[PbWrapType](pbwraptype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В следующем примере добавляет овала active публикации и указывает, что текст публикации обтекания слева и справа прямоугольника, охватывающего овала.


```vb
Sub SetTextWrapType() 
 Dim shpOval As Shape 
 
 Set shpOval = ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeOval, Left:=36, Top:=36, _ 
 Width:=100, Height:=35) 
 
 With shpOval.TextWrap 
 .Type = pbWrapTypeSquare 
 .Side = pbWrapSideBoth 
 End With 
End Sub
```


