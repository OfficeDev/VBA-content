---
title: "Метод MailMergeDataField.MapToRecipientField (издатель)"
keywords: vbapb10.chm6422563
f1_keywords: vbapb10.chm6422563
ms.prod: publisher
api_name: Publisher.MailMergeDataField.MapToRecipientField
ms.assetid: d3da8a00-e2ca-b07b-cc8f-02d729cb149c
ms.date: 06/08/2017
ms.openlocfilehash: 0ee0311f935974f652dabffc998ad98a9abc5ed6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldmaptorecipientfield-method-publisher"></a>Метод MailMergeDataField.MapToRecipientField (издатель)

Сопоставляет поля (столбца) в источнике данных, представленный родительский объект **MailMergeDataField** к полю получателя (столбца) в источник данных (список получателей объединенный слияния почты).


## <a name="syntax"></a>Синтаксис

 _выражение_. **MapToRecipientField** ( **_bstrValue_**)

 переменная _expression_A, представляет собой объект- **MailMergeDataField** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|bstrValue|Необязательный| **String**|Имя поля получателя, который должен быть сопоставлен столбца источника данных.|

## <a name="remarks"></a>Заметки

Этот метод работает только в том случае, если родительский объект **MailMergeDataField** еще не сопоставлен к полю получателя. Свойство **[IsMapped](mailmergedatafield-ismapped-property-publisher.md)** объекта **MailMergeDataField** можно использовать для определения, если объект уже сопоставлен.

Если значение параметра необязательно bstrValue не передается, Microsoft Publisher предполагается, что поля для сопоставления имеет то же имя получателя поля в источник данных, к которому оно сопоставлено.

Если передать имя поля, которое не существует, Publisher, возвращается ошибка. 


 **Примечание**  Чтобы добавить поле, используйте метод **[AddToRecipientFields](mailmergedatafield-addtorecipientfields-method-publisher.md)** .


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **MapToRecipientField** для отображения данных поля (столбца) для источника данных в поле в источник данных (комбинированное списка получателей) для публикации.

Прежде чем запустить этот макрос, замените _datasourceindex_ номер индекса допустимый источник данных в коллекции источника данных активного документа, заменить _fieldname_ с именем поля в источнике данных, которое необходимо сопоставить к полю получателя и заменить _recipientfieldname_ с именем поля получателя.

В разделе **[элемент](mailmergedatasources-item-method-publisher.md)** метод пример того, как использовать свойство **Name** объекта **DataSource** для определения номера индекса требуемый источник данных.




```vb
Public Sub Map() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataField As Publisher.MailMergeDataField 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 Set pubMailMergeDataField = pubMailMergeDataSources.Item(datasourceindex).DataFields.Item("fieldname") 
 
 If pubMailMergeDataField.IsMapped Then 
 
 Debug.Print "This field is already mapped" 
 
 Else 
 
 pubMailMergeDataField.MapToRecipientField ("recipientfieldname") 
 Debug.Print "Field mapped successfully." 
 
 End If 
 
End Sub
```


