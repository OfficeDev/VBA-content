---
title: "Свойство Application.ActiveDocument (издатель)"
keywords: vbapb10.chm131073
f1_keywords: vbapb10.chm131073
ms.prod: publisher
api_name: Publisher.Application.ActiveDocument
ms.assetid: c6293fa6-291c-d8ce-be54-f8a997b95d2e
ms.date: 06/08/2017
ms.openlocfilehash: a0581aad811846da9bda6187393fbc975d257fcb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationactivedocument-property-publisher"></a>Свойство Application.ActiveDocument (издатель)

Возвращает объект **[Document](document-object-publisher.md)** , представляющий active публикации. При наличии открыть доступ к документам, возникает ошибка.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ActiveDocument**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Документ


## <a name="example"></a>Пример

В этом примере пользователь может назначить имя файла для активной публикации и сохраните его с новым именем файла. Имя файла вместе с другими текст вставляется после текущего выбранного текста. (Обратите внимание на то, что имя файла, необходимо заменить имя допустимого публикации для работы этого примера.)


```vb
Sub NewsLetterSave() 
 
 Dim strFileName As String 
 
 ' Assign the explicit file name to a variable. 
 strFileName = "Filename" 
 Publisher.ActiveDocument.SaveAs strFileName 
 
 ' Insert the file name and supporting text after selected text. 
 Selection.TextRange.Collapse pbCollapseEnd 
 Selection.TextRange = _ 
 " This publication has been saved as " &; strFileName 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

