---
title: "Свойство TabStop.Alignment (издатель)"
keywords: vbapb10.chm5636100
f1_keywords: vbapb10.chm5636100
ms.prod: publisher
api_name: Publisher.TabStop.Alignment
ms.assetid: 59b35d9a-d53b-88cd-952b-6324d1ee7c01
ms.date: 06/08/2017
ms.openlocfilehash: efe49db2025e456bd9b1cc73406f044c0f4d49d8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabstopalignment-property-publisher"></a>Свойство TabStop.Alignment (издатель)

Возвращает или задает значение константы **PbTabAlignmentType** , представляющий выравнивание для заданной позиции табуляции. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выравнивание**

 переменная _expression_A, представляет собой объект- **TabStop** .


## <a name="remarks"></a>Заметки

Значение свойства **Alignment** может иметь одно из **[PbTabAlignmentType](pbtabalignmenttype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере переходит в список с вкладками и задает выравнивание для двух табуляции. В этом примере предполагается, что указанные форму — фрагмент текста и не другого типа фигуры и задать, что по крайней мере два табуляции уже существует.


```vb
Sub CustomDecimalTabStop() 
 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .InsertAfter Newtext:="Pencils" &; vbTab &; _ 
 "Each" &; vbTab &; "1.50" &; vbLf 
 .InsertAfter Newtext:="Pens" &; vbTab &; _ 
 "Each" &; vbTab &; "4.95" &; vbLf 
 .InsertAfter Newtext:="Folders" &; vbTab &; _ 
 "Box" &; vbTab &; "35.28" &; vbLf 
 .InsertAfter Newtext:="Envelopes" &; vbTab &; _ 
 "Case" &; vbTab &; "150.69" &; vbLf 
 With .Paragraphs(Start:=1).ParagraphFormat 
 .Tabs(1).Alignment = pbTabAlignmentCenter 
 .Tabs(2).Alignment = pbTabAlignmentDecimal 
 End With 
 End With 
End Sub
```


