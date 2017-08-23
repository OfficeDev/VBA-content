---
title: "Свойство ParagraphFormat.Alignment (издатель)"
keywords: vbapb10.chm5439491
f1_keywords: vbapb10.chm5439491
ms.prod: publisher
api_name: Publisher.ParagraphFormat.Alignment
ms.assetid: db66f8b8-a813-418c-2735-e5299e6a6045
ms.date: 06/08/2017
ms.openlocfilehash: 84e1ff5b8b4ddf759c29a42ccacd1c3b43a2c757
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatalignment-property-publisher"></a>Свойство ParagraphFormat.Alignment (издатель)

Возвращает или задает значение константы **PbParagraphAlignmentType** , представляющий выравнивание для указанного абзацев. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выравнивание**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


## <a name="remarks"></a>Заметки

Значение свойства **Alignment** может иметь одно из **[PbParagraphAlignmentType](pbparagraphalignmenttype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере добавляется новое текстовое поле для первой страницы active публикации и затем добавить текст и задает выравнивание абзаца и форматирование шрифта.


```vb
Sub NewTextFrame() 
 Dim shpTextBox As Shape 
 Set shpTextBox = ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=468, Height:=72) 
 With shpTextBox.TextFrame.TextRange 
 .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
 .Text = "Hello World" 
 With .Font 
 .Name = "Snap ITC" 
 .Size = 30 
 .Bold = msoTrue 
 End With 
 End With 
End Sub
```


