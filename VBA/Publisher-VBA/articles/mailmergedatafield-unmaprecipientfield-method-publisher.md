---
title: "Метод MailMergeDataField.UnMapRecipientField (издатель)"
keywords: vbapb10.chm6422564
f1_keywords: vbapb10.chm6422564
ms.prod: publisher
api_name: Publisher.MailMergeDataField.UnMapRecipientField
ms.assetid: 0063dfa7-1168-3701-56a3-f1908cf0d23a
ms.date: 06/08/2017
ms.openlocfilehash: aa2bdff4025091ce50372bc94187fe6eb1d56bc7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldunmaprecipientfield-method-publisher"></a>Метод MailMergeDataField.UnMapRecipientField (издатель)

Отменяет сопоставление между **MailMergeDataField** родительский объект для источника данных и поле получателя в источник данных (список получателей объединенный слияния почты), для которого выполняется в данный момент сопоставление.


## <a name="syntax"></a>Синтаксис

 _выражение_. **UnMapRecipientField**

 переменная _expression_A, представляет собой объект- **MailMergeDataField** .


## <a name="remarks"></a>Заметки

Этот метод работает только в том случае, если родительский объект **MailMergeDataField** сопоставляется к полю получателя. Свойство **[IsMapped](mailmergedatafield-ismapped-property-publisher.md)** объекта **MailMergeDataField** можно использовать для определения, если объект сопоставлен.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **UnmapRecipientField** для отмены сопоставление данных поля (столбца) для источника данных и поля в основных данных источника (объединенный список получателей) для публикации.

Прежде чем запустить этот макрос, замените _datasourceindex_ номер индекса допустимый источник данных в коллекции источника данных активных документов и заменить _fieldname_ с именем поля в источнике данных, который требуется удалить из объединенного списка полями получателей.

В разделе **[элемент](mailmergedatasources-item-method-publisher.md)** метод пример того, как использовать свойство **Name** объекта **DataSource** для определения номера индекса требуемый источник данных.




```vb
Public Sub UnmapRecipientField_Example() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataField As Publisher.MailMergeDataField 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 Set pubMailMergeDataField = pubMailMergeDataSources.Item(datasourceindex).DataFields.Item("fieldname") 
 
 If pubMailMergeDataField.IsMapped Then 
 
 pubMailMergeDataField.UnMapRecipientField 
 Debug.Print "Field unmapped succesfully." 
 
 Else 
 
 Debug.Print "This field is not mapped." 
 
 End If 
 
End Sub
```


