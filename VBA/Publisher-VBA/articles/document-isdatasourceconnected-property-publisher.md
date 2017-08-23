---
title: "Свойство Document.IsDataSourceConnected (издатель)"
keywords: vbapb10.chm196722
f1_keywords: vbapb10.chm196722
ms.prod: publisher
api_name: Publisher.Document.IsDataSourceConnected
ms.assetid: b62422ab-12f7-1151-d8d1-1cb32de18160
ms.date: 06/08/2017
ms.openlocfilehash: 38690b1c1ab0a08f369369c8fdb0e8022c1552e0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentisdatasourceconnected-property-publisher"></a>Свойство Document.IsDataSourceConnected (издатель)

 **Значение true** , если указанная публикация подключена к источнику данных. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsDataSourceConnected**

 переменная _expression_A, представляющий объект **документа** .


## <a name="remarks"></a>Заметки

Публикации должен быть подключен к допустимым источником данных для выполнения слияния или объединения в каталог.


## <a name="example"></a>Пример

Следующий пример проверяет ли публикации подключается к источнику данных и, если это не так, указывает и подключается к публикации в источнике данных. 

Перед запуском этого примера необходимо заменить _PathToFile_ допустимый путь к файлу и _TableName_ с именем таблицы источника допустимых данных.




```vb
Dim strDataSource As String 
Dim strDataSourceTable As String 
 
 'Specify data source and table name 
 
 strDataSource = "PathToFile" 
 strDataSourceTable = "TableName" 
 
 'Connect to a datasource 
 If Not (ThisDocument.IsDataSourceConnected) Then 
 ThisDocument.MailMerge.OpenDataSource strDataSource, , strDataSourceTable 
 
 End If
```


