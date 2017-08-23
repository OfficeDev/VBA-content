---
title: "Метод Document.BeginCustomUndoAction (издатель)"
keywords: vbapb10.chm196709
f1_keywords: vbapb10.chm196709
ms.prod: publisher
api_name: Publisher.Document.BeginCustomUndoAction
ms.assetid: 316f443e-6782-594b-b955-f5ab60140f6a
ms.date: 06/08/2017
ms.openlocfilehash: d90457576886837ee5ca2c80bffec48c09db0afd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentbegincustomundoaction-method-publisher"></a>Метод Document.BeginCustomUndoAction (издатель)

Указывает начальную точку и метка (текстовое описание) группы действий, реализуемые для создания единого отменить действие. Метод **[EndCustomUndoAction](document-endcustomundoaction-method-publisher.md)** используется для указания конечной точки действия, используемые для создания единого отменить действие. Перенос группы действий можно отменить с помощью одной операции отмены.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeginCustomUndoAction** ( **_Действие_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя действия|Обязательное свойство.| **String**|Метка, которая соответствует одной отменить действие. Эта метка отображается, если щелкнуть стрелку рядом с кнопкой кнопка "Отменить" на панели инструментов Стандартная.|

## <a name="remarks"></a>Заметки

Следующие методы объекта **Document** , недоступны в настраиваемых отменить действие. Если какие-либо из этих методов вызываются в настраиваемых отменить действие, возвращается ошибка во время выполнения:


-  **Document.Close**
    
-  **Document.MailMerge.DataSource.Close**
    
-  **Document.PrintOut**
    
-  **Document.Redo**
    
-  **Document.Save**
    
-  **Document.SaveAs**
    
-  **Document.Undo**
    
-  **Document.UndoClear**
    
-  **Document.UpdateOLEObjects**
    


**BeginCustomUndoAction** метод необходимо вызывать до вызова метода **EndCustomUndoAction** . Если **EndCustomUndoAction** вызывается до **BeginCustomUndoAction**, возвращается ошибка во время выполнения.

Вложение настраиваемой отменить действие в рамках другого настраиваемого отменить действие разрешено, но вложенных настраиваемых отменить действие не оказывает влияния. Только внешний настраиваемых отменить действие является активным.


## <a name="example"></a>Пример

Следующий пример содержит два действия настраиваемой отмены. На первой странице активная публикация создается первый из них. Метод **BeginCustomUndoAction** используется для указания точки, в которой должно начаться настраиваемых отменить действие. Шесть отдельные действия выполняются, а затем они помещаются в одно действие при вызове **EndCustomUndoAction**. 

Чтобы определить, является ли шрифт Verdana протестирована текста в элементе frame текст, который был создан в первый настраиваемых отменить действие. В противном случае метод **Отменить** вызывается с **[UndoActionsAvailable](document-undoactionsavailable-property-publisher.md)** передается как параметр. В этом случае имеется только один отменить действие. Таким образом, вызов ** [Отменить метод](document-undo-method-publisher.md)** отменяет только одно действие, но это действие один переход шесть действий в одну.

Создается второй отменить действие, а также может быть отменено более поздней версии с помощью операции отмены одного.




```vb
Dim thePage As page 
Dim theShape As Shape 
Dim theDoc As Publisher.Document 
 
Set theDoc = ActiveDocument 
Set thePage = theDoc.Pages(1) 
 
With theDoc 
 ' The following six actions are wrapped to create one 
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


