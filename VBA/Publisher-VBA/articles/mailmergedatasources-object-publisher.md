---
title: "Объект MailMergeDataSources (издатель)"
keywords: vbapb10.chm7274495
f1_keywords: vbapb10.chm7274495
ms.prod: publisher
api_name: Publisher.MailMergeDataSources
ms.assetid: 9eff8354-fbc3-7f55-ba6e-738a60f41259
ms.date: 06/08/2017
ms.openlocfilehash: ccf3d29f25c25d092fbce64dcded797d12ceb20f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasources-object-publisher"></a>Объект MailMergeDataSources (издатель)

Представляет коллекцию всех объектов **вывода** в активном документе Microsoft Publisher, каждый из которых представляет один из источников данных в ходе операции слияния почты.
 


## <a name="remarks"></a>Заметки

Элемент по умолчанию коллекции **MailMergeDataSources** — это метод **элемента** , который возвращает объект **вывода** в указанной позиции индекса.
 

 
Если имеется только один объект **вывода** в активном документе, коллекция **MailMergeDataSources** пуста. В этом случае при попытке получить значение свойства **DataSources** объекта **вывода** Publisher возвращает ошибку.
 

 

## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как получить имена всех подключенных источников данных в коллекции **MailMergeDataSources** в активный документ. Он используется свойство **IsDataSourceConnected** активного документа для определения, подключен ли в источнике данных.
 

 
Если подключение одного или нескольких источников данных, макрос использует свойство **Count** коллекции **MailMergeDataSources** для определения подключенные сколько источников данных.
 

 
Если подключен только один источник данных, макрос печатает имя этого источника данных в окне **интерпретации** ; Если подключено более одного источника данных, метод **элемента** коллекции **MailMergeDataSources** используется для итерации по коллекции и свойство **Name** объекта **вывода** для печати имя каждого источника данных в окне **интерпретации** .
 

 



```
Public Sub MailMergeDataSources_Example() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataSource As Publisher.MailMergeDataSource 
 Dim lngCount As Long 
 Dim intCounter As Integer 
 
 If ThisDocument.IsDataSourceConnected Then 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 
 lngCount = pubMailMergeDataSources.Count 
 
 If lngCount > 1 Then 
 
 ' More than one data source is connected. 
 For intCounter = 1 To lngCount 
 Debug.Print pubMailMergeDataSources.Item(intCounter).Name 
 Next 
 
 Else 
 
 ' Only one data source is connected. 
 Set pubMailMergeDataSource = ThisDocument.MailMerge.DataSource 
 Debug.Print "Only one data source ("; pubMailMergeDataSource.Name; ") is connected!" 
 
 End If 
 
 Else 
 
 Debug.Print "No data sources are connected!" 
 
 End If 
 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](mailmergedatasources-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](mailmergedatasources-application-property-publisher.md)|
|[Count](mailmergedatasources-count-property-publisher.md)|
|[Создатель](mailmergedatasources-creator-property-publisher.md)|
|[Родительский раздел](mailmergedatasources-parent-property-publisher.md)|

