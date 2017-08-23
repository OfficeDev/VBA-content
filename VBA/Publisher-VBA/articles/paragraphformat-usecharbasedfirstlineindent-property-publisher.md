---
title: "Свойство ParagraphFormat.UseCharBasedFirstLineIndent (издатель)"
keywords: vbapb10.chm5439529
f1_keywords: vbapb10.chm5439529
ms.prod: publisher
api_name: Publisher.ParagraphFormat.UseCharBasedFirstLineIndent
ms.assetid: c2ac44ab-6671-5851-ac62-7449fd646cc5
ms.date: 06/08/2017
ms.openlocfilehash: 5d7f0cc808a9e591d4c2de2d828e0249f50c9937
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatusecharbasedfirstlineindent-property-publisher"></a>Свойство ParagraphFormat.UseCharBasedFirstLineIndent (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, является ли абзац с отступом с помощью ширину восточно-азиатских символов. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **UseCharBasedFirstLineIndent**

 переменная _expression_A, представляющий объект **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **UseCharBasedFirstLineIndent** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Значение **UseCharBasedFirstLineIndent** можно задать только в том случае, если на клиентском компьютере включены восточно-азиатских языков, а значение может быть возвращено независимо от того, включен ли восточно-азиатских языков. Обратите внимание, что перед **[CharBasedFirstLineIndent](paragraphformat-charbasedfirstlineindent-property-publisher.md)** свойство должно быть задано **UseCharBasedFirstLineIndent** можно возвращаются или задаются. Если **UseCharBasedFirstLineIndent** установлено сначала возвращается ошибку времени выполнения «отказано в разрешении».

Если **UseCharBasedFirstLineIndent** **msoTrue**, абзац с отступом с помощью ширину знаков восточно-азиатских языков, и если это **msoFalse** не. Значение по умолчанию — **msoFalse**.


## <a name="example"></a>Пример

В следующем примере создается текстовое поле на странице четвертый active публикации. После **UseCharBasedFirstLineIndent** задано значение **True**, ширина отступ первой строки задано 15 точек с помощью свойства **CharBasedFirstLineIndent** . Задайте свойства шрифта и вставки текста в абзац.


```vb
Dim theTextBox As Shape 
 
Set theTextBox = ActiveDocument.Pages(4).Shapes _ 
 .AddShape(msoShapeRectangle, 100, 100, 300, 200) 
 
With theTextBox 
 .TextFrame.TextRange.ParagraphFormat _ 
 .UseCharBasedFirstLineIndent = msoTrue 
 .TextFrame.TextRange.ParagraphFormat _ 
 .CharBasedFirstLineIndent = 15 
 .TextFrame.TextRange.Font.Name = "Verdana" 
 .TextFrame.TextRange.Font.Size = 12 
 .TextFrame.TextRange.Text = "This is a test sentence." _ 
 &; Chr(13) &; "This is another test sentence." 
End With
```


