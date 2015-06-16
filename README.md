###Exportador de datos a Excel
Genera un Excel con una hoja de configuración que define el resto de hojas que contienen los datos exportados.
La estructura de la configuración es:
- WorksheetName
- ConfigurationTypeName
- TypeName
- HeaderRange
- DataRange

            var excelWriter = new ExcelWriter(FilePath, _items);
            excelWriter.Write();

###Importador de datos a Excel
Lee los datos del excel y por reflexión, los genera en memoria con los mismo tipos con los que fueron guardados.

            var excelReader = new ExcelReader(FilePath);
            var itemsRead = excelReader.Read();

