---
title: "Метод MailMergeDataField.AddToRecipientFields (издатель)"
keywords: vbapb10.chm6422562
f1_keywords: vbapb10.chm6422562
ms.prod: publisher
api_name: Publisher.MailMergeDataField.AddToRecipientFields
ms.assetid: eaf365f0-a9f4-c6e2-1267-d0a31b5934ce
ms.date: 06/08/2017
ms.openlocfilehash: fc21e1e7db16eabed890e087c14afdaafcae9973
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldaddtorecipientfields-method-publisher"></a>Метод MailMergeDataField.AddToRecipientFields (издатель)

Добавляет родительский объект **MailMergeDataField** из конкретного источника данных к источнику данных master (коллекцию полей данных) для публикации слияния почты.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddToRecipientFields**

 переменная _expression_A, представляет собой объект- **MailMergeDataField** .


## <a name="remarks"></a>Заметки

Этот метод работает только в том случае, если родительский объект **MailMergeDataField** еще не сопоставлен к полю получателя. Свойство **[IsMapped](mailmergedatafield-ismapped-property-publisher.md)** объекта **MailMergeDataField** можно использовать для определения, если объект уже сопоставлен.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **AddToRecipientFields** для добавления данных поля (столбца) для источника данных для основного источника данных (комбинированное списка получателей) для публикации.

Прежде чем запустить этот макрос, замените _datasourceindex_ номер индекса допустимый источник данных в коллекции источника данных активных документов и заменить _fieldname_ с именем поля в источнике данных, который требуется добавить к списку объединенный полями получателей.

В разделе **[элемент](mailmergedatasources-item-method-publisher.md)** метод пример того, как использовать свойство **Name** объекта **DataSource** для определения номера индекса требуемый источник данных.




```vb
Public Sub AddToRecipientFields_Example() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataField As Publisher.MailMergeDataField 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 Set pubMailMergeDataField = pubMailMergeDataSources.Item(datasourceindex).DataFields.Item("fieldname") 
 
 If pubMailMergeDataField.IsMapped Then 
 
 Debug.Print "This field is already mapped!" 
 
 Else 
 
 pubMailMergeDataField.AddToRecipientFields 
 Debug.Print "Field added successfully. (You can verify this by looking at the recipient or product list in the UI.)" 
 
 End If 
 
End Sub
```


