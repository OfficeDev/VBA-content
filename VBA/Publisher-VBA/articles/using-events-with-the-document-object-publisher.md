---
title: "Использование событий с помощью объекта Document (издатель)"
ms.prod: publisher
ms.assetid: 0f5cfe67-bfa1-0ec7-11c9-c4c1337ebe50
ms.date: 06/08/2017
ms.openlocfilehash: 1bd4dfc2cd847588f0dcd5b87b8e238fb950a2f5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="using-events-with-the-document-object-publisher"></a>Использование событий с помощью объекта Document (издатель)

Объект **Document** поддерживает семь события: **[BeforeClose](document-beforeclose-event-publisher.md)**, **[Open](document-open-event-publisher.md)**, **[Повторить](document-redo-event-publisher.md)**, **[ShapesAdded](document-shapesadded-event-publisher.md)**, **[ShapesRemoved](document-shapesremoved-event-publisher.md)**, **[Отменить](document-undo-event-publisher.md)**и **[WizardAfterChange](document-wizardafterchange-event-publisher.md)**. Написание процедуры реагировать на эти события в модуле класса с именем «ThisDocument». Используйте следующие шаги для создания процедуры обработки событий.


1. В разделе публикации проекта в окне **Обозреватель проектов** дважды щелкните **ThisDocument**. (В представлении **папки** **ThisDocument** расположен в папке **Объекты Microsoft Publisher** .)
    
2. Выберите **документ** из поля раскрывающегося списка **объектов** .
    
3. Выберите событие в раскрывающемся списке **процедуры** . Пустой подпрограмму добавляется в модуле класса.
    
4. Добавление команды Visual Basic, которые необходимо выполнить при возникновении события.
    

## <a name="example"></a>Пример

В этом примере показано **Открытие** процедуру события, которая отображает сообщение при открытии публикации.


```vb
Private Sub Document_Open() 
    MsgBox "This publication is copyrighted." 
End Sub
```

В следующем примере показано процедуру события **BeforeClose** , запрашивает у Да или нет ответа пользователя перед закрытием документа.




```vb
Private Sub Document_BeforeClose(Cancel As Boolean) 
    Dim intResponse As Integer 
 
    intResponse = MsgBox("Do you really want to close " _ 
        &; "the document?", vbYesNo) 
 
    If intResponse = vbNo Then Cancel = True 
End Sub
```


 **Примечание**  Сведения о создании процедур обработки событий для объекта **Application** содержатся в разделе [С помощью событий объекта](using-events-with-the-application-object-publisher.md) .


