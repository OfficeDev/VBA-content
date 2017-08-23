---
title: "Метод Document.EndCustomUndoAction (издатель)"
keywords: vbapb10.chm196710
f1_keywords: vbapb10.chm196710
ms.prod: publisher
api_name: Publisher.Document.EndCustomUndoAction
ms.assetid: 5b703366-8d0e-1bbc-3320-a2fea99468c3
ms.date: 06/08/2017
ms.openlocfilehash: 77286813bf3ae1ae280a40c1ddd024e526d52a1f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentendcustomundoaction-method-publisher"></a>Метод Document.EndCustomUndoAction (издатель)

Конечная группа действий, реализуемые для создания единого отменить действие. ** [Метод BeginCustomUndoAction](document-begincustomundoaction-method-publisher.md)** метод используется для указания начальной точки и метки (текстовое описание) действия, используемые для создания единого отменить действие. Перенос группы действий можно отменить с помощью одной операции отмены.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EndCustomUndoAction**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

**BeginCustomUndoAction** метод необходимо вызывать до вызова метода **EndCustomUndoAction** . Если **EndCustomUndoAction** вызывается до **BeginCustomUndoAction**, возвращается ошибка во время выполнения.


## <a name="example"></a>Пример

Следующий пример содержит два действия настраиваемой отмены. Первый создается на странице четырех active публикации. Метод **BeginCustomUndoAction** используется для указания точки, в которой должно начаться настраиваемых отменить действие. Шесть отдельные действия выполняются, а затем они помещаются в одно действие при вызове **EndCustomUndoAction**. 

Чтобы определить, является ли шрифт Verdana протестирована текста в элементе frame текст, который был создан в первый настраиваемых отменить действие. В противном случае метод **[Отменить](document-undo-method-publisher.md)** вызывается с **[UndoActionsAvailable](document-undoactionsavailable-property-publisher.md)** передается как параметр. В этом случае имеется только один отменить действие. Таким образом вызов для **отмены** отменяет только одно действие, но это действие один переход шесть действий в одну.

Создается второй отменить действие, а также может быть отменено более поздней версии с помощью операции отмены одного.

В этом примере предполагается, что активная публикация содержит по крайней мере четыре страницы.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(4) 
 
With theDoc 
 ' The following six of actions are wrapped to create one 
 ' custom undo action named "Add Rectangle and Courier Text". 
 .BeginCustomUndoAction ("Add Rectangle and Courier Text") 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeRectangle, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 14 
 .Font.Bold = msoTrue 
 .Font.Name = "Courier" 
 .Text = "This font is Courier." 
 End With 
 End With 
 .EndCustomUndoAction 
 
 If Not thePage.Shapes(1).TextFrame.TextRange.Font.Name = "Verdana" Then 
 ' This call to Undo will undo all actions that are available. 
 ' In this case, there is only one action that can be undone. 
 .Undo (.UndoActionsAvailable) 
 ' A new custom undo action is created with a name of 
 ' "Add Balloon and Verdana Text". 
 .BeginCustomUndoAction ("Add Balloon and Verdana Text") 
 With thePage 
 Set theShape = .Shapes.AddShape(msoShapeBalloon, _ 
 75, 75, 190, 30) 
 With theShape.TextFrame.TextRange 
 .Font.Size = 11 
 .Font.Name = "Verdana" 
 .Text = "This font is Verdana." 
 End With 
 End With 
 .EndCustomUndoAction 
 End If 
End With
```


