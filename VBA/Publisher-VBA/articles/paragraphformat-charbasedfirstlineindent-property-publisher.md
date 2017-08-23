---
title: "Свойство ParagraphFormat.CharBasedFirstLineIndent (издатель)"
keywords: vbapb10.chm5439528
f1_keywords: vbapb10.chm5439528
ms.prod: publisher
api_name: Publisher.ParagraphFormat.CharBasedFirstLineIndent
ms.assetid: d0432be6-2e6a-39fa-9e9a-0300a0437f35
ms.date: 06/08/2017
ms.openlocfilehash: 08186e2468bade5b281ea0a598f7c346a5de636e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatcharbasedfirstlineindent-property-publisher"></a>Свойство ParagraphFormat.CharBasedFirstLineIndent (издатель)

Возвращает или задает значение отступ первой строки (в ширину знаков). Чтение и запись **времени**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CharBasedFirstLineIndent**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Значение **CharBasedFirstLineIndent** может возвращаются или задаются только после установки **[UseCharBasedFirstLineIndent](paragraphformat-usecharbasedfirstlineindent-property-publisher.md)** . Если **UseCharBasedFirstLineIndent** установлено сначала возвращается ошибку времени выполнения «отказано в разрешении». Обратите внимание, что **UseCharBasedFirstLineIndent** можно задать только в том случае, если на клиентском компьютере, (независимо от того, включен ли восточно-азиатских языков может быть возвращаемое значение) включены восточно-азиатских языков. Эффективно, это означает, что **CharBasedFirstLineIndent** не может использоваться, если на клиентском компьютере не разрешены восточно-азиатских языков.

Значение **CharBasedFirstLineIndent** находится в диапазоне от 0 (ноль) до 80.


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


