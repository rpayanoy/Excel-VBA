Sub crearNuevoLibro()
    'Crear un nuevo Libro desde uno existente
    Dim nuevoLibro As Workbook
    Dim rutaLibroExistente As String
    Dim nombreNuevoLibro As String
    
    'Ruta del libro existente
    rutaLibroExistente = ThisWorkbook.Path
    
    'Asignar nombre para el nuevo libro
    nombreNuevoLibro = "Nombre del nuevo Libro"
    
    'Añadir nuevo libro
    Set nuevoLibro = Workbooks.Add
    
    'Guardar el nuevo Libro
    nuevoLibro.SaveAs Filename:=rutaLibroExistente & "\" & nombreNuevoLibro, _
    FileFormat:=51, Password:="", WriteResPassword:="", _
    ReadOnlyRecommended:=False, CreateBackup:=False
    'FileFormat -> 51: (.xlsx) | 52: (.xlsm)
    
    'Cerrar el nuevo Libro
    nuevoLibro.Close
End Sub
