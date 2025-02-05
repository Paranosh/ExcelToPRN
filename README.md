Convierte un archivo de excel a archivo de texto plano PRN encolumnado de longitud fija, teniendo en cuenta la celda de mayor longitud del excel.
El funcionamiento es simple, arrastrar el archivo .xlsx al ejecutable y en el directorio del archivo de origen creará dos nuevos archivos, uno el prn y otro en el que indicará de la primera linea del prn, el texto del campo, la posición inicial de ese texto en la linea y su longitud.
Por defecto no se muestra la consola, por lo que en caso de error no proporciona ningún tipo de información, para que se muestra cambbiar la linea <OutputType>WinExe</OutputType> por <OutputType>Exe</OutputType>, del archivo ExcelToPRN.csproj

![imagen](https://github.com/user-attachments/assets/a5d335b8-ad9d-4063-be85-93f46001c9b8)
![imagen](https://github.com/user-attachments/assets/20cd661a-d2d9-4603-91b3-f75e1a93ed14)
![imagen](https://github.com/user-attachments/assets/83ab46b4-369f-4230-b2d9-26999391592b)

