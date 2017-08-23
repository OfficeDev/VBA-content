---
title: "Свойство Document.UndoActionsAvailable (издатель)"
keywords: vbapb10.chm196726
f1_keywords: vbapb10.chm196726
ms.prod: publisher
api_name: Publisher.Document.UndoActionsAvailable
ms.assetid: 1dd20295-3987-c36d-ccc1-9e18a7887f33
ms.date: 06/08/2017
ms.openlocfilehash: 129b8ec3d9fc7e9a677c55dce28adae1d9612c52
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentundoactionsavailable-property-publisher"></a>Свойство Document.UndoActionsAvailable (издатель)

Возвращает число действий, доступных в стеке отмены. Только для чтения **времени**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **UndoActionsAvailable**

 переменная _expression_A, представляющий объект **документа** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В следующем примере добавляется прямоугольника, который содержит фрагмент текста на страницу четвертый active публикации. Некоторые свойства шрифта и текст надписи зависят от выбора. Затем запускается для определения, является ли шрифт в элементе frame текст Courier. Если это так, метод **[Отменить](document-undo-method-publisher.md)** используется со значением свойства **UndoActionsAvailable** , передается как параметр для указания, что все предыдущие действия быть отменены.

Метод **[Повторить](document-redo-method-publisher.md)** нажмите используется со значением свойства **[RedoActionsAvailable](document-redoactionsavailable-property-publisher.md)** минус 2, передается как параметр для возврата всех действий, за исключением последние два. Новый шрифт указан текст в текстовой рамки, в дополнение к новый текст.

В этом примере предполагается, что активный документ содержит по крайней мере четыре страницы.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
With theDoc 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 12 
 .Font.Name = "Courier" 
 .Text = "This font is Courier." 
 End With 
 End With 
 
 If thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Courier" Then 
 .Undo (.UndoActionsAvailable) 
 .Redo (.RedoActionsAvailable - 2) 
 With theShape.TextFrame.TextRange 
 .Font.Name = "Verdana" 
 .Text = "This font is Verdana." 
 End With 
 End If 
End With
```


