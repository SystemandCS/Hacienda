Module Publices




    Public blnNuevo As Boolean
    Public blnModificar As Boolean
    Public idCuarteo As Long
    Public idRemitoExp As Long

    Public PosicionGrillaCuarteo As Long
    Public PosicionGrillaCuarteo2 As Long
    Public intIdCuarteo As Long
    Public IntIdDespo As Long
    Public intidMres As Long
    Public intNroCuarteo As Long
    Public intsecuencia As Long
    Public fechacuarteo As String
    Public intNroTropa As Long
    Public blnDesposte As Boolean
    Public intCantCabezas As Long
    Public intIdSC As Long
    Public intFichaactiva As Long
    Public intidUsrFrigorifico As Long
    Public intIdMr As Long
    Public intNroRomaneo As Double
    Public idfrigorifico As Long
    Public strFechaFaena As String
    Public intmaxRubro As Long
    Public blnAgregueRubro As Boolean
    Public idSector As Long
    Public intmaxCatGan As Long
    Public blnAgregueCatGan As Boolean
    Public intmaxCatIva As Long
    Public blnAgregueCatIva As Boolean
    Public intmaxCondPago As Long
    Public blnAgregueCondPago As Boolean
    Public intmaxCondCompra As Long
    Public blnAgregueCondCompra As Boolean
    Public blnCambioEstado As Boolean
    Public strNombreUsr As String
    Public strApeUsr As String
    Public blnCambioSector As Boolean
    Public strSector As String
    Public blnEntre As Boolean
    Public RecibeMerc As Boolean
    Public IntNroSol As Long
    Public IntNroPC As Long
    Public strProveedores As String
    Public CantSol As Double
    Public CantRecibida As Double
    Public IdEntrega As Long
    Public IdPcDet As Long
    Public strCantidadUnidad As String
    Public FlagCompras As Long
    Public AgregueProd As Boolean
    Public IdProd As Long
    Public strProd As String
    Public idFrigFiltro As Long
    Public idStockFiltro As Long
    Public fechaIngreso
    Public strnombrearchivo As String
    Public strnombrearchivodetalle As String
    Public blnVengoDeFaena As Boolean
    Public Jerarquia As String
    Public idLogin As Long
    Public descrip_jera As String
    Public sorden As Long
    Public sbande As Boolean
    Public habilitamenuUsr As Boolean
    Public nroCuarteo As Long
    Public nroTRopa As Long
    Public strFechaFaena1 As String
    Public intCorte As Long
    Public strFEchaCuarteo As String
    Public StockCuarteo As Long
    Public strTipoAccesoUsrHac As String
    Public strTipoAccesoUsrStock As String
    Public strTipoAccesoUsrUsu As String
    Public strFechaDespostada As String
    Public Proveedores() As String
    Public idFrigorificoFiltrado As Long
    Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)



End Module