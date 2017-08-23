---
title: "Метод TextRange.InsertBarcode (издатель)"
keywords: vbapb10.chm5308502
f1_keywords: vbapb10.chm5308502
ms.prod: publisher
api_name: Publisher.TextRange.InsertBarcode
ms.assetid: ad613ca7-f056-55b0-1a96-51167555ce6f
ms.date: 06/08/2017
ms.openlocfilehash: 2ce5f4cdbf1e1ab057eaa4dd60c14bde6fc7a17b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeinsertbarcode-method-publisher"></a>Метод TextRange.InsertBarcode (издатель)

Вставка поля штрих-код в конец диапазона текста, представленного объектом **TextRange** родительского.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InsertBarcode**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

В идеальном случае следует создать надстройку для Microsoft Publisher для обработки событий, **[MailMergeGenerateBarcode](application-mailmergegeneratebarcode-event-publisher.md)** и **[MailMergeInsertBarcode](application-mailmergeinsertbarcode-event-publisher.md)** . Если вашей надстройки или код не содержит обработчики для этих событий, метод **InsertBarcode** возвращает ошибку.

В следующем примере показано, как обрабатывать следующие события с помощью Microsoft Visual Basic для приложений (VBA) в редакторе Visual Basic.

Если вы хотите включить вставку штрих-кодов в публикации из пользовательского интерфейса, вашей надстройки или кода VBA следует также необходимо задать значение свойства **[InsertBarcodeVisible](application-insertbarcodevisible-property-publisher.md)** значение **True**.


## <a name="example"></a>Пример

Следующем примере показано, как использовать метод **InsertBarcode** для вставки поля штрих кода в текстовом поле в публикации. Вставьте следующий код в проекта VBA и выполнить процедуру, **AttachToEvents** перед выполнением процедуры **InsertBarcode_Example** .

Прежде чем запускать код в следующем примере, используйте ** [MailMerge.OpenDataSource](mailmerge-opendatasource-method-publisher.md)** метод для подключения к источнику данных. Источник данных должен содержать столбец штрих кода, в котором приведены штрих-кодов для всех получателей слияния почты. Замените _barcodeColumnIndex_ в обработчике событий **MailMergeGenerateBarcode** в коде индекс столбца источника данных, который содержит штрих-код.

Запустите следующий код в окне **Редактора Visual Basic** и не в диалоговом окне **макросов** . (В меню **Сервис** выберите пункт **макрос**и нажмите кнопку макросы.)




```vb
Public WithEvents pubApplication As Publisher.Application 
 
Private Sub pubApplication_MailMergeGenerateBarcode(ByVal Doc As Document, bstrString As String) 
 
    bstrString = pubApplication.ActiveDocument.MailMerge.DataSource.DataFields.Item(barcodeColumnIndex).Value 
         
End Sub 
 
Private Sub pubApplication_MailMergeInsertBarcode(ByVal Doc As Document, OkToInsert As Boolean) 
 
    OkToInsert = True 
     
End Sub 
 
Public Sub InsertBarcode_Example() 
 
    Dim pubTextRange As Publisher.TextRange 
    Dim pubShape As Publisher.Shape 
     
    Set pubShape = ThisDocument.Pages(1).Shapes.AddTextbox(pbTextOrientationHorizontal, 100, 100, 500, 500) 
    Set pubTextRange = pubShape.TextFrame.TextRange 
     
    pubTextRange.InsertBarcode 
     
End Sub 
 
Public Sub AttachToEvents() 
 
    Set pubApplication = Application 
 
End Sub
```


