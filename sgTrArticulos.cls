VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sgTrArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub ALTA()

    Load frmAbmArticulos

    frmAbmArticulos.Modo = "A"
    frmAbmArticulos.IniciarABM
    
    frmAbmArticulos.caption = "Agregar artículo ..."
    
    frmAbmArticulos.Show vbModal
    Screen.MousePointer = vbNormal

End Sub

Public Sub ALTAMED()

    Load frmAbmMedicamentos

    frmAbmMedicamentos.Modo = "A"
    frmAbmMedicamentos.IniciarABM
    
    frmAbmMedicamentos.caption = "Agregar Medicamento ..."
    
    frmAbmMedicamentos.Show vbModal
    Screen.MousePointer = vbNormal

End Sub

Public Sub ALTAPRODMED()

    Load frmAbmProdMedicos

    frmAbmProdMedicos.Modo = "A"
    frmAbmProdMedicos.IniciarABM
    
    frmAbmProdMedicos.caption = "Agregar Producto Médico ..."
    frmAbmProdMedicos.Show vbModal
    Screen.MousePointer = vbNormal

End Sub

Public Sub ALTAVARIOS()

    Load frmAbmVarios

    frmAbmVarios.Modo = "A"
    frmAbmVarios.IniciarABM
    
    frmAbmVarios.caption = "Agregar Varios ..."

    frmAbmVarios.Show vbModal
    Screen.MousePointer = vbNormal

End Sub

Public Sub ALTAVARIOSALIMENTACION()

    Load frmAbmVariosAlimentacion

    frmAbmVariosAlimentacion.Modo = "A"
    frmAbmVariosAlimentacion.IniciarABM
    
    frmAbmVariosAlimentacion.caption = "Agregar Varios Alimentación..."

    frmAbmVariosAlimentacion.Show vbModal
    Screen.MousePointer = vbNormal

End Sub

Public Sub ALTAVARIOS_SERVGRALES()

    Load frmABMVarios_ServGrales

    frmABMVarios_ServGrales.Modo = "A"
    frmABMVarios_ServGrales.IniciarABM
    
    frmABMVarios_ServGrales.caption = "Agregar Varios Servicios Generales..."

    frmABMVarios_ServGrales.Show vbModal
    Screen.MousePointer = vbNormal

End Sub
Public Function ArmarCriterios(ByVal Relevancia) As sgTRCriterios
    On Error Resume Next
    
    Dim Crit  As sgTRCriterio
    Dim sSQL As String

    Set ArmarCriterios = New sgTRCriterios

    'Nombre
    ArmarCriterios.Add "ARTI_NOMBRE", "Nombre", sgTexto, "ARTI_NOMBRE"
    
    'Estado
    sSQL = "SELECT INTE_TEXTO AS MOSTRAR, INTE_ITEM AS TAG FROM REFER_INTERNAS WHERE INTE_CODIGO=14011"
    Set Crit = ArmarCriterios.Add("ARTI_ESTADO", "Estado", sgNumerico, "ARTI_ESTADO")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL
    
    'Relevancia
    sSQL = "SELECT INTE_TEXTO AS MOSTRAR, INTE_ITEM AS TAG FROM REFER_INTERNAS WHERE INTE_CODIGO=14002"
    Set Crit = ArmarCriterios.Add("ARTI_RELEVANCIA", "Relevancia", sgNumerico, "ARTI_RELEVANCIA")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL
    
    'Linea
    sSQL = "SELECT GRAL_DESCRIPCION AS MOSTRAR, GRAL_IDREG AS TAG FROM TABLA_GENERAL WHERE GRAL_TABLA=52001"
    Set Crit = ArmarCriterios.Add("ARTI_LINEA", "Línea", sgNumerico, "ARTI_LINEA")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL
    
    'Fabricante
    sSQL = "Select LABO_DESCRIPCION as MOSTRAR, LABO_CODIGO  as TAG From SGEHELP.dbo.LABORATORIOS order by LABO_DESCRIPCION "
    Set Crit = ArmarCriterios.Add("ARTI_FABRICANTE", "Laboratorio", sgNumerico, "ARTI_FABRICANTE")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL
    
    'Proveedor
    sSQL = "Select PROVNOMBRE as MOSTRAR, PROVIDREG  as TAG From PROVEEDORES order by PROVNOMBRE "
    Set Crit = ArmarCriterios.Add("ARTI_PROVEEDOR", "Proveedor", sgNumerico, "ARTI_PROVEEDOR")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL
    
    'Grupo Compras
    sSQL = "Select GRUCMP_NOMBRE as MOSTRAR, GRUCMP_IDREG  as TAG From GRUPOS_COMPRAS order by GRUCMP_NOMBRE "
    Set Crit = ArmarCriterios.Add("ARTI_GRUPOCMP", "Grupo", sgNumerico, "ARTI_GRUPOCMP")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL
    
    'Clase
    sSQL = "SELECT GRAL_DESCRIPCION AS MOSTRAR, GRAL_IDREG AS TAG FROM TABLA_GENERAL WHERE GRAL_TABLA=52007"
    Set Crit = ArmarCriterios.Add("ARTI_CLASE", "Clase", sgNumerico, "ARTI_CLASE")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL

    'Familia
    sSQL = "SELECT GRAL_DESCRIPCION AS MOSTRAR, GRAL_IDREG AS TAG FROM TABLA_GENERAL WHERE GRAL_TABLA=52008"
    Set Crit = ArmarCriterios.Add("ARTI_FAMILIA", "Familia", sgNumerico, "ARTI_FAMILIA")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL

    'Tipo item
    sSQL = "SELECT INTE_TEXTO AS MOSTRAR, INTE_ITEM AS TAG FROM REFER_INTERNAS WHERE INTE_CODIGO=14004"
    Set Crit = ArmarCriterios.Add("ARTI_TIPOITEM", "Tipo Item", sgNumerico, "ARTI_TIPOITEM")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL

    'Tipo vencimiento
    sSQL = "SELECT INTE_TEXTO AS MOSTRAR, INTE_ITEM AS TAG FROM REFER_INTERNAS WHERE INTE_CODIGO=14005"
    Set Crit = ArmarCriterios.Add("ARTI_VENCIMIENTO", "Tipo Vencimiento", sgNumerico, "ARTI_VENCIMIENTO")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL
    
    'Clasificacion
    sSQL = "SELECT RTRIM(CLASARTI_NIVEL1) + ' - ' + RTRIM(CLASARTI_NIVEL2) + ' - ' + RTRIM(CLASARTI_NIVEL3) AS MOSTRAR, "
    sSQL = sSQL & " CLASARTI_IDREG FROM CLASIFIC_ARTICULO "
    If Relevancia <> 0 Then
        sSQL = sSQL & " WHERE CLASARTI_RELEVA = " & Relevancia
    End If
    sSQL = sSQL & " ORDER BY CLASARTI_NIVEL1, CLASARTI_NIVEL2, CLASARTI_NIVEL3"
    Set Crit = ArmarCriterios.Add("ARTI_CLASIFICACION", "Clasificacion", sgNumerico, "ARTI_CLASIFICACION")
    Crit.UsaCombo = True
    Crit.SQLCombo = sSQL
    
    'Fecha carga
    ArmarCriterios.Add "ARTI_FECPRO", "Fecha carga", sgFecha, "ARTI_FECPRO"
    
    'Usuario
    ArmarCriterios.Add "ARTI_USUARIO", "Usuario", sgTexto, "ARTI_USUARIO"
    
End Function
Public Function GrabarAlta(ByVal Arti As sgDArticulos, ByVal Catalogos As sgColArtiCata, ByVal Depositos As sgColArtiDepo) As Boolean
    On Error Resume Next

    Dim cn As ADODB.Connection
    Dim sSQL As String
    Dim lRes As Long
    Dim ih As Integer
    Dim idReg As Long

    Dim mIdRegArtiCata As Long
    Dim mIdRegArtiDepo As Long
    
    Dim i As Integer
    
    
    If z_EsCodigoArticuloAlta(Arti.Codigo) Then
        MsgBox "El artículo ya existe en la base, no puede ser ingresado nuevamente", vbCritical, "Error"
        GrabarAlta = False
        Exit Function
    End If
      
    
    On Error GoTo errorsql

    Set cn = New ADODB.Connection
    cn.Open user.sgEmpresas.Seleccionada.EMPRCONEXION
    cn.BeginTrans

    On Error GoTo cnError

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
            
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.ARTICULOS_INSERT"
    cmd.Parameters.Refresh
 
    Arti.idReg = ObtenerProximoID(user.sgEmpresas.Seleccionada.EMPRCONEXION, "ARTICULOS", "ARTI_IDREG")
    
    cmd.Parameters(1) = Arti.idReg
    cmd.Parameters(2) = Arti.Codigo
    cmd.Parameters(3) = Arti.CodigoBarras
    cmd.Parameters(4) = Arti.Etiqueta
    cmd.Parameters(5) = Arti.Nombre
    cmd.Parameters(6) = Arti.Descripcion
    cmd.Parameters(7) = Arti.Gravado
    cmd.Parameters(8) = Arti.Fabricante
    cmd.Parameters(9) = Arti.Observac
    cmd.Parameters(10) = Arti.Imagen
    
    cmd.Parameters(11) = Arti.FechaAlta
    cmd.Parameters(12) = Arti.FechaEstado
    cmd.Parameters(13) = Arti.Estado
    cmd.Parameters(14) = Arti.TipoDescri
    cmd.Parameters(15) = Arti.Relevancia
    cmd.Parameters(16) = Arti.Linea
    cmd.Parameters(17) = Arti.CodOficial
    cmd.Parameters(18) = Arti.Origen
    cmd.Parameters(19) = Arti.Tratamiento
    cmd.Parameters(20) = Arti.TipoItem
    
    cmd.Parameters(21) = Arti.Vencimiento
    cmd.Parameters(22) = Arti.Clasificacion
    cmd.Parameters(23) = Arti.Clase
    cmd.Parameters(24) = Arti.Familia
    cmd.Parameters(25) = Arti.ItemPrecio
    cmd.Parameters(26) = Arti.Utilidad
    cmd.Parameters(27) = Arti.Bonificacion
    cmd.Parameters(28) = Arti.IndUniPrecioCompra
    cmd.Parameters(29) = Arti.UniMedPreciocompra
    cmd.Parameters(30) = Arti.FactorUnidadPrecioCompra
    cmd.Parameters(31) = Arti.IndUniPrecioVenta
    
    cmd.Parameters(32) = Arti.UniMedPrecioVenta
    cmd.Parameters(33) = Arti.FactorUnidadPrecioVenta
    cmd.Parameters(34) = Arti.ModoCompra
    cmd.Parameters(35) = Arti.Proveedor
    cmd.Parameters(36) = Arti.ProvCodigoArticulo
    cmd.Parameters(37) = Arti.GrupoCompra
    cmd.Parameters(38) = Arti.IndUniDetalleCompra
    cmd.Parameters(39) = Arti.UniMedDetalleCompra
    cmd.Parameters(40) = Arti.FactorUniDetalleCompra
    cmd.Parameters(41) = Arti.FormaBultoCompra
    cmd.Parameters(42) = Arti.UnidadesxBultoCompra
    cmd.Parameters(43) = Arti.CompraMinima
    
    cmd.Parameters(44) = Arti.ItemGranel
    cmd.Parameters(45) = Arti.CantGranel
    cmd.Parameters(46) = Arti.IndUniDetalleVenta
    cmd.Parameters(47) = Arti.FactorUniDetalleVenta
    cmd.Parameters(48) = Arti.UniMedDetalleVenta
    cmd.Parameters(49) = Arti.FormaBultoVenta
    cmd.Parameters(50) = Arti.UnidadesxBultoVenta
    cmd.Parameters(51) = Arti.idReg
    cmd.Parameters(52) = Arti.ModoStock
    cmd.Parameters(53) = Arti.UniMedidaStock
    cmd.Parameters(54) = Arti.FormaBultoStock
    cmd.Parameters(55) = Arti.UnixBultoStock
    
    cmd.Parameters(56) = Arti.FactorConv
    cmd.Parameters(57) = Arti.ModoReposi
    cmd.Parameters(58) = Arti.StockMinimo
    cmd.Parameters(59) = Arti.StockMaximo
    cmd.Parameters(60) = Arti.StockReposi
    cmd.Parameters(61) = user.USUACODIGO
    cmd.Parameters(62) = ConvertirFecha(Date)
    cmd.Parameters(63) = Format(Time, "HH:mm")
    cmd.Execute lRes
    
    If lRes = 0 Then GoTo cnError
    
    mIdRegArtiCata = ObtenerProximoID(user.sgEmpresas.Seleccionada.EMPRCONEXION, "ARTICULO_CATALOGOS", "ARTICATA_IDREG")
    
    For i = 1 To Catalogos.Count
        If Not Catalogos.Item(i).idCatalogo = 0 Then
            sSQL = "INSERT INTO ARTICULO_CATALOGOS (ARTICATA_IDREG, ARTICATA_ARTICULO,"
            sSQL = sSQL & "ARTICATA_CATALOGO, ARTICATA_ADICIONAL, "
            sSQL = sSQL & "ARTICATA_USUARIO, ARTICATA_FECPRO, ARTICATA_HORPRO) VALUES ("
            sSQL = sSQL & mIdRegArtiCata & ","
            sSQL = sSQL & Arti.idReg & ","
            sSQL = sSQL & Catalogos.Item(i).idCatalogo & ", '"
            sSQL = sSQL & Catalogos.Item(i).Adicionales & "', "
            sSQL = sSQL & "'" & Arti.Usuario & "',"
            sSQL = sSQL & ConvertirFecha(Date) & ","
            sSQL = sSQL & "'" & Format(Now, "hh:mm:ss") & "')"
            cn.Execute sSQL, lRes
            If lRes = 0 Then GoTo cnError
            
            mIdRegArtiCata = mIdRegArtiCata + 1
        End If
    Next
    
    mIdRegArtiDepo = ObtenerProximoID(user.sgEmpresas.Seleccionada.EMPRCONEXION, "ARTICULO_DEPOSITOS", "ARTIDEPO_IDREG")
    
    For i = 1 To Depositos.Count
    
        sSQL = "INSERT INTO ARTICULO_DEPOSITOS (ARTIDEPO_IDREG, ARTIDEPO_ARTICULO,"
        sSQL = sSQL & "ARTIDEPO_DEPOSITO, ARTIDEPO_UBICACION, ARTIDEPO_MODALIDAD,ARTIDEPO_SINEGATIVO,"
        sSQL = sSQL & "ARTIDEPO_STOCKMINIMO, ARTIDEPO_STOCKMAXIMO, "
        sSQL = sSQL & "ARTIDEPO_STOCKREPOSI, ARTIDEPO_CONSUMO, ARTIDEPO_USUARIO,"
        sSQL = sSQL & "ARTIDEPO_FECPRO, ARTIDEPO_HORPRO) VALUES ("
        sSQL = sSQL & mIdRegArtiDepo & ","
        sSQL = sSQL & Arti.idReg & ","
        sSQL = sSQL & Depositos.Item(i).idDeposito & ", '"
        sSQL = sSQL & Depositos.Item(i).Ubicacion & "',"
        sSQL = sSQL & Depositos.Item(i).Modalidad & ",'"
        sSQL = sSQL & Depositos.Item(i).SiNegativo & "', "
        sSQL = sSQL & Depositos.Item(i).StockMinimo & ", "
        sSQL = sSQL & Depositos.Item(i).StockMaximo & ", "
        sSQL = sSQL & Depositos.Item(i).StockReposi & ", "
        sSQL = sSQL & "0, "
        sSQL = sSQL & "'" & Arti.Usuario & "',"
        sSQL = sSQL & ConvertirFecha(Date) & ","
        sSQL = sSQL & "'" & Format(Now, "hh:mm:ss") & "')"
        cn.Execute sSQL, lRes
        If lRes = 0 Then GoTo cnError
        
        mIdRegArtiDepo = mIdRegArtiDepo + 1
        
    Next
    
    sSQL = "EXEC p_BorrarUsuCodTabCaptura @NOMTABLA = 'ARTICULOS_CAPTURA'" & ", @COD_USUARIO='" & Trim(user.USUACODIGO) & "'" & ", @CODIGO='" & Trim(Arti.Codigo) & "'"
    cn.Execute sSQL, lRes

    'If lRes = 0 Then GoTo cnError
    
    MsgBox "El artículo " & Arti.Descripcion & " ha sido creado con exito.", vbOKOnly + vbInformation, "Incorporación de artículos ..."
    cn.CommitTrans
    GrabarAlta = True

    cn.Close
    Set cn = Nothing

Exit Function

errorsql:
    MsgBox "Ha ocurrido un error al intentar crear el artículo." & vbCrLf & "Comuniquese con mantenimiento.", vbOKOnly + vbCritical, "Fallo en el proceso de ingresos..."
    GuardarErrorTransacciones cn.Errors, "al crear el artículo - [SGTRARTICULOS]"
    cn.RollbackTrans
    cn.Close
    Set cn = Nothing
    GrabarAlta = False

    Exit Function

cnError:
    MostrarErrorDeTransacciones cn.Errors, "al crear un artículo"
    cn.RollbackTrans
    cn.Close
    Set cn = Nothing
    GrabarAlta = False

End Function


Public Sub Modificar()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmArticulos
    frmAbmArticulos.Modo = "M"
    'frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11).value

    frmAbmArticulos.IniciarABM
    
    frmAbmArticulos.caption = "Modificar artículos ..."
    frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmAbmArticulos.Show vbModal

End Sub
Public Sub CambiarEstadoArticulo()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmCambioEstadoArticulo
    frmCambioEstadoArticulo.Modo = "E"
    'frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11).value

    'frmCambioEstadoArticulo.CargarDatos
    
    frmCambioEstadoArticulo.caption = "Cambiar Estado artículos ..."
    frmCambioEstadoArticulo.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmCambioEstadoArticulo.Show vbModal

End Sub
Public Sub CambiarEstadoProdMedicos()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmCambioEstadoArticulo
    frmCambioEstadoArticulo.Modo = "E"
    'frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11).value

    'frmCambioEstadoArticulo.CargarDatos
    
    frmCambioEstadoArticulo.caption = "Cambiar Estado artículos ..."
    frmCambioEstadoArticulo.CargarDatos frmTratProdMedicos.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmCambioEstadoArticulo.Show vbModal

End Sub
Public Sub CambiarEstadoArticuloVarios()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmCambioEstadoArticulo
    frmCambioEstadoArticulo.Modo = "E"
    'frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11).value

    'frmCambioEstadoArticulo.CargarDatos
    
    frmCambioEstadoArticulo.caption = "Cambiar Estado artículos ..."
    frmCambioEstadoArticulo.CargarDatos frmTratVarios.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmCambioEstadoArticulo.Show vbModal

End Sub

Public Sub CambiarEstadoVariosAlimentacion()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmCambioEstadoArticulo
    frmCambioEstadoArticulo.Modo = "E"
    'frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11).value

    'frmCambioEstadoArticulo.CargarDatos
    
    frmCambioEstadoArticulo.caption = "Cambiar Estado artículos ..."
    frmCambioEstadoArticulo.CargarDatos frmTratVariosAlimentacion.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmCambioEstadoArticulo.Show vbModal

End Sub
Public Sub CambiarEstadoVariosServGrales()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmCambioEstadoArticulo
    frmCambioEstadoArticulo.Modo = "E"
    'frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11).value

    'frmCambioEstadoArticulo.CargarDatos
    
    frmCambioEstadoArticulo.caption = "Cambiar Estado artículos ..."
    frmCambioEstadoArticulo.CargarDatos frmTratVariosServGrales.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmCambioEstadoArticulo.Show vbModal

End Sub
Public Sub CambiarEstadoMedicamento()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmCambioEstadoArticulo
    frmCambioEstadoArticulo.Modo = "E"
    'frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11).value

    'frmCambioEstadoArticulo.CargarDatos
    
    frmCambioEstadoArticulo.caption = "Cambiar Estado Medicamentos ..."
    frmCambioEstadoArticulo.CargarDatos frmTratMedicamentos.sgGrilla.Rows.Current.Cells(10)

    Screen.MousePointer = vbNormal
    frmCambioEstadoArticulo.Show vbModal

End Sub
Public Sub ModificarMed()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmMedicamentos
    frmAbmMedicamentos.Modo = "M"
    frmAbmMedicamentos.IniciarABM
    
    frmAbmMedicamentos.caption = "Modificar artículos ..."
    frmAbmMedicamentos.CargarDatos frmTratMedicamentos.sgGrilla.Rows.Current.Cells(10)

    Screen.MousePointer = vbNormal
    frmAbmMedicamentos.Show vbModal

End Sub
Public Sub ModificarProdMed()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmProdMedicos
    frmAbmProdMedicos.Modo = "M"
    frmAbmProdMedicos.IniciarABM
    
    frmAbmProdMedicos.caption = "Modificar artículos ..."
    frmAbmProdMedicos.CargarDatos frmTratProdMedicos.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmAbmProdMedicos.Show vbModal

End Sub
Public Sub ModificarVarios()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmVarios
    frmAbmVarios.Modo = "M"
    frmAbmVarios.IniciarABM
    
    frmAbmVarios.caption = "Modificar artículos ..."
    frmAbmVarios.CargarDatos frmTratVarios.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmAbmVarios.Show vbModal

End Sub

Public Sub ModificarVariosAlimentacion()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmVariosAlimentacion
    frmAbmVariosAlimentacion.Modo = "M"
    frmAbmVariosAlimentacion.IniciarABM
    
    frmAbmVariosAlimentacion.caption = "Modificar Artículos Alimentación..."
    frmAbmVariosAlimentacion.CargarDatos frmTratVariosAlimentacion.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmAbmVariosAlimentacion.Show vbModal

End Sub

Public Sub ModificarVarios_ServGrales()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmABMVarios_ServGrales
    frmABMVarios_ServGrales.Modo = "M"
    frmABMVarios_ServGrales.IniciarABM
    
    frmABMVarios_ServGrales.caption = "Modificar Artículos Alimentación..."
    frmABMVarios_ServGrales.CargarDatos frmTratVariosServGrales.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmABMVarios_ServGrales.Show vbModal

End Sub

Public Sub ConsultarMed()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmMedicamentos
    frmAbmMedicamentos.Modo = "C"
    frmAbmMedicamentos.IniciarABM
    
    frmAbmMedicamentos.caption = "Consultar artículos ..."
    frmAbmMedicamentos.CargarDatos frmTratMedicamentos.sgGrilla.Rows.Current.Cells(10)

    Screen.MousePointer = vbNormal
    frmAbmMedicamentos.Show vbModal

End Sub
Public Sub ConsultarProdMed()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmProdMedicos
    frmAbmProdMedicos.Modo = "C"
    frmAbmProdMedicos.IniciarABM
    
    frmAbmProdMedicos.caption = "Consultar artículos ..."
    frmAbmProdMedicos.CargarDatos frmTratProdMedicos.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmAbmProdMedicos.Show vbModal

End Sub
Public Sub ConsultarVarios()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmVarios
    frmAbmVarios.Modo = "C"
    frmAbmVarios.IniciarABM
    
    frmAbmVarios.caption = "Consultar artículos ..."
    frmAbmVarios.CargarDatos frmTratVarios.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmAbmVarios.Show vbModal

End Sub
Public Sub ConsultarVariosAlimentacion()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmVariosAlimentacion
    frmAbmVariosAlimentacion.Modo = "C"
    frmAbmVariosAlimentacion.IniciarABM
    
    frmAbmVariosAlimentacion.caption = "Consultar artículos ..."
    frmAbmVariosAlimentacion.CargarDatos frmTratVariosAlimentacion.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmAbmVariosAlimentacion.Show vbModal

End Sub

Public Sub ConsultarVarios_ServGrales()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmABMVarios_ServGrales
    frmABMVarios_ServGrales.Modo = "C"
    frmABMVarios_ServGrales.IniciarABM
    
    frmABMVarios_ServGrales.caption = "Consultar artículos ..."
    frmABMVarios_ServGrales.CargarDatos frmTratVariosServGrales.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmABMVarios_ServGrales.Show vbModal

End Sub
Public Sub Consultar()

    On Error Resume Next
    Screen.MousePointer = vbHourglass

    Load frmAbmArticulos
    frmAbmArticulos.Modo = "C"
    frmAbmArticulos.IniciarABM
    
    frmAbmArticulos.caption = "Consultar artículos ..."
    frmAbmArticulos.CargarDatos frmTratArticulos.sgGrilla.Rows.Current.Cells(11)

    Screen.MousePointer = vbNormal
    frmAbmArticulos.Show vbModal

End Sub



 Public Function GrabarCambioEstado(ByVal Arti As sgDArticulos) As Boolean
    On Error Resume Next

    Dim cn As ADODB.Connection
    Dim sSQL As String
    Dim lRes As Long
    Dim ih As Integer
    Dim idReg As Long

        
    Dim i As Integer
    
    On Error GoTo errorsql

    Set cn = New ADODB.Connection
    cn.Open user.sgEmpresas.Seleccionada.EMPRCONEXION
    cn.BeginTrans

    On Error GoTo cnError

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
            
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.ARTICULOS_UPDATE_ESTADO"
    cmd.Parameters.Refresh
    
    cmd.Parameters(1) = Arti.idReg
    
    cmd.Parameters(2) = Arti.FechaEstado
    cmd.Parameters(3) = Arti.Estado
    
    
    
    cmd.Parameters(4) = user.USUACODIGO
    cmd.Parameters(5) = ConvertirFecha(Date)
    cmd.Parameters(6) = Format(Time, "HH:mm")
    
    cmd.Execute lRes
    If lRes = 0 Then GoTo cnError
            
    MsgBox "El estado del artículo " & Arti.Descripcion & " ha sido modificado con exito.", vbOKOnly + vbInformation, "Incorporación de artículos ..."
    cn.CommitTrans
    GrabarCambioEstado = True

    cn.Close
    Set cn = Nothing

Exit Function

errorsql:
    MsgBox "Ha ocurrido un error al intentar modificar el artículo." & vbCrLf & "Comuniquese con mantenimiento.", vbOKOnly + vbCritical, "Fallo en el proceso de ingresos..."
    GuardarErrorTransacciones cn.Errors, "al crear el artículo - [SGTRARTICULOS]"
    cn.RollbackTrans
    cn.Close
    Set cn = Nothing
    GrabarCambioEstado = False

    Exit Function

cnError:
    MostrarErrorDeTransacciones cn.Errors, "al modificar estado de un artículo"
    cn.RollbackTrans
    cn.Close
    Set cn = Nothing
    GrabarCambioEstado = False
 
 End Function

Public Function GrabarModificacion(ByVal Arti As sgDArticulos, ByVal Catalogos As sgColArtiCata, ByVal Depositos As sgColArtiDepo, Optional sNomAnterior) As Boolean
    
    On Error Resume Next

    Dim cn As ADODB.Connection
    Dim sSQL As String
    Dim lRes As Long
    Dim ih As Integer
    Dim idReg As Long

    Dim mIdRegArtiCata As Long
    Dim mIdRegArtiDepo As Long
    
    Dim rs As ADODB.Recordset
    
    Dim i As Integer
    
    On Error GoTo errorsql

    Set cn = New ADODB.Connection
    cn.Open user.sgEmpresas.Seleccionada.EMPRCONEXION
    cn.BeginTrans

    On Error GoTo cnError

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
            
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.ARTICULOS_UPDATE"
    cmd.Parameters.Refresh
    
    cmd.Parameters(1) = Arti.idReg
    cmd.Parameters(2) = Arti.Codigo
    cmd.Parameters(3) = Arti.CodigoBarras
    cmd.Parameters(4) = Arti.Etiqueta
    cmd.Parameters(5) = Arti.Nombre
    cmd.Parameters(6) = Arti.Descripcion
    cmd.Parameters(7) = Arti.Gravado
    cmd.Parameters(8) = Arti.Fabricante
    cmd.Parameters(9) = Arti.Observac
    cmd.Parameters(10) = Arti.Imagen
    
    cmd.Parameters(11) = Arti.FechaAlta
    cmd.Parameters(12) = Arti.FechaEstado
    cmd.Parameters(13) = Arti.Estado
    cmd.Parameters(14) = Arti.TipoDescri
    cmd.Parameters(15) = Arti.Relevancia
    cmd.Parameters(16) = Arti.Linea
    cmd.Parameters(17) = Arti.CodOficial
    cmd.Parameters(18) = Arti.Origen
    cmd.Parameters(19) = Arti.Tratamiento
    cmd.Parameters(20) = Arti.TipoItem
    
    cmd.Parameters(21) = Arti.Vencimiento
    cmd.Parameters(22) = Arti.Clasificacion
    cmd.Parameters(23) = Arti.Clase
    cmd.Parameters(24) = Arti.Familia
    cmd.Parameters(25) = Arti.ItemPrecio
    cmd.Parameters(26) = Arti.Utilidad
    cmd.Parameters(27) = Arti.Bonificacion
    cmd.Parameters(28) = Arti.IndUniPrecioCompra
    cmd.Parameters(29) = Arti.UniMedPreciocompra
    cmd.Parameters(30) = Arti.FactorUnidadPrecioCompra
    cmd.Parameters(31) = Arti.IndUniPrecioVenta
    
    cmd.Parameters(32) = Arti.UniMedPrecioVenta
    cmd.Parameters(33) = Arti.FactorUnidadPrecioVenta
    cmd.Parameters(34) = Arti.ModoCompra
    cmd.Parameters(35) = Arti.Proveedor
    cmd.Parameters(36) = Arti.ProvCodigoArticulo
    cmd.Parameters(37) = Arti.GrupoCompra
    cmd.Parameters(38) = Arti.IndUniDetalleCompra
    cmd.Parameters(39) = Arti.UniMedDetalleCompra
    cmd.Parameters(40) = Arti.FactorUniDetalleCompra
    cmd.Parameters(41) = Arti.FormaBultoCompra
    cmd.Parameters(42) = Arti.UnidadesxBultoCompra
    cmd.Parameters(43) = Arti.CompraMinima
    
    cmd.Parameters(44) = Arti.ItemGranel
    cmd.Parameters(45) = Arti.CantGranel
    cmd.Parameters(46) = Arti.IndUniDetalleVenta
    cmd.Parameters(47) = Arti.FactorUniDetalleVenta
    cmd.Parameters(48) = Arti.UniMedDetalleVenta
    cmd.Parameters(49) = Arti.FormaBultoVenta
    cmd.Parameters(50) = Arti.UnidadesxBultoVenta
    cmd.Parameters(51) = Arti.ItemRealStk
    cmd.Parameters(52) = Arti.ModoStock
    cmd.Parameters(53) = Arti.UniMedidaStock
    cmd.Parameters(54) = Arti.FormaBultoStock
    cmd.Parameters(55) = Arti.UnixBultoStock
    
    cmd.Parameters(56) = Arti.FactorConv
    cmd.Parameters(57) = Arti.ModoReposi
    cmd.Parameters(58) = Arti.StockMinimo
    cmd.Parameters(59) = Arti.StockMaximo
    cmd.Parameters(60) = Arti.StockReposi
    cmd.Parameters(61) = user.USUACODIGO
    cmd.Parameters(62) = ConvertirFecha(Date)
    cmd.Parameters(63) = Format(Time, "HH:mm")
    
    cmd.Execute lRes
    If lRes = 0 Then GoTo cnError
    
    If (Arti.TipoDescri = 1 Or Arti.TipoDescri = 21) And sNomAnterior <> Arti.Descripcion And Len(sNomAnterior) > 0 Then
        sSQL = "UPDATE ORDENCOMPRA_DETALLE SET OCOMPRADETA_DESCRIPCION = '" & Arti.Descripcion & "' WHERE OCOMPRADETA_IDARTICULO = " & Arti.idReg
        cn.Execute sSQL
        
        sSQL = "UPDATE PREORDEN_DETALLE SET PREORDDETA_DESCRIPCION = '" & Arti.Descripcion & "' WHERE PREORDDETA_IDARTICULO = " & Arti.idReg
        cn.Execute sSQL
        
        sSQL = "UPDATE SOLCOMPRA_DETALLE SET SCOMDETA_DESCRIPCION = '" & Arti.Descripcion & "' WHERE SCOMDETA_IDARTICULO = " & Arti.idReg
        cn.Execute sSQL
        
        sSQL = "UPDATE TRANSAC_DETALLE SET TRANDETA_DESCRIPCION = '" & Arti.Descripcion & "' WHERE TRANDETA_ARTICULO = " & Arti.idReg
        cn.Execute sSQL
    End If
    
    mIdRegArtiCata = ObtenerProximoID(user.sgEmpresas.Seleccionada.EMPRCONEXION, "ARTICULO_CATALOGOS", "ARTICATA_IDREG")
    
    sSQL = "DELETE ARTICULO_CATALOGOS WHERE ARTICATA_ARTICULO = " & Arti.idReg
    cn.Execute sSQL
    
    For i = 1 To Catalogos.Count
        If Not Catalogos.Item(i).idCatalogo = 0 Then
            sSQL = "INSERT INTO ARTICULO_CATALOGOS (ARTICATA_IDREG, ARTICATA_ARTICULO,"
            sSQL = sSQL & "ARTICATA_CATALOGO, ARTICATA_ADICIONAL, "
            sSQL = sSQL & "ARTICATA_USUARIO, ARTICATA_FECPRO, ARTICATA_HORPRO) VALUES ("
            sSQL = sSQL & mIdRegArtiCata & ","
            sSQL = sSQL & Arti.idReg & ","
            sSQL = sSQL & Catalogos.Item(i).idCatalogo & ", '"
            sSQL = sSQL & Catalogos.Item(i).Adicionales & "', "
            sSQL = sSQL & "'" & Arti.Usuario & "',"
            sSQL = sSQL & ConvertirFecha(Date) & ","
            sSQL = sSQL & "'" & Format(Now, "hh:mm:ss") & "')"
            
            cn.Execute sSQL, lRes
            If lRes = 0 Then GoTo cnError
            
            mIdRegArtiCata = mIdRegArtiCata + 1
        End If
    Next
    
    mIdRegArtiDepo = ObtenerProximoID(user.sgEmpresas.Seleccionada.EMPRCONEXION, "ARTICULO_DEPOSITOS", "ARTIDEPO_IDREG")
    
    sSQL = "UPDATE ARTICULO_DEPOSITOS SET ARTIDEPO_USUARIO = 'BORRAR' WHERE ARTIDEPO_ARTICULO = " & Arti.idReg
    cn.Execute sSQL
    
    Set rs = New ADODB.Recordset
    For i = 1 To Depositos.Count
        
        If Len(Trim(Depositos.Item(i).idDeposito)) > 0 Then
            
            sSQL = "SELECT ARTIDEPO_IDREG FROM ARTICULO_DEPOSITOS WHERE ARTIDEPO_ARTICULO = " & Arti.idReg
            sSQL = sSQL & " AND ARTIDEPO_DEPOSITO = " & Depositos.Item(i).idDeposito
            rs.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic, adCmdText
                    
            If rs.EOF Then
                
                sSQL = "INSERT INTO ARTICULO_DEPOSITOS (ARTIDEPO_IDREG, ARTIDEPO_ARTICULO,"
                sSQL = sSQL & "ARTIDEPO_DEPOSITO, ARTIDEPO_UBICACION,ARTIDEPO_MODALIDAD,ARTIDEPO_SINEGATIVO,"
                sSQL = sSQL & "ARTIDEPO_STOCKMINIMO, ARTIDEPO_STOCKMAXIMO, ARTIDEPO_STOCKREPOSI,"
                sSQL = sSQL & "ARTIDEPO_STOCKNODISPO, ARTIDEPO_STOCKFISICO, ARTIDEPO_USUARIO,"
                sSQL = sSQL & "ARTIDEPO_FECPRO, ARTIDEPO_HORPRO) VALUES ("
                sSQL = sSQL & mIdRegArtiDepo & ","
                sSQL = sSQL & Arti.idReg & ","
                sSQL = sSQL & Depositos.Item(i).idDeposito & ", '"
                sSQL = sSQL & Depositos.Item(i).Ubicacion & "',"
                sSQL = sSQL & Depositos.Item(i).Modalidad & ",'"
                sSQL = sSQL & Depositos.Item(i).SiNegativo & "', "
                sSQL = sSQL & Replace(Depositos.Item(i).StockMinimo, ",", ".") & ", "
                sSQL = sSQL & Replace(Depositos.Item(i).StockMaximo, ",", ".") & ", "
                sSQL = sSQL & Replace(Depositos.Item(i).StockReposi, ",", ".") & ", "
                sSQL = sSQL & Replace(Depositos.Item(i).StockNoDispo, ",", ".") & ", "
                sSQL = sSQL & Replace(Depositos.Item(i).StockFisico, ",", ".") & ", "
                sSQL = sSQL & "'" & Arti.Usuario & "',"
                sSQL = sSQL & ConvertirFecha(Date) & ","
                sSQL = sSQL & "'" & Format(Now, "hh:mm:ss") & "')"
                
                mIdRegArtiDepo = mIdRegArtiDepo + 1
            
            Else
                
                sSQL = "UPDATE ARTICULO_DEPOSITOS SET "
                sSQL = sSQL & "ARTIDEPO_UBICACION = '" & Depositos.Item(i).Ubicacion & "',"
                sSQL = sSQL & "ARTIDEPO_MODALIDAD = " & Depositos.Item(i).Modalidad & ","
                sSQL = sSQL & "ARTIDEPO_SINEGATIVO = '" & Depositos.Item(i).SiNegativo & "',"
                sSQL = sSQL & "ARTIDEPO_STOCKMINIMO = " & Replace(Depositos.Item(i).StockMinimo, ",", ".") & ", "
                sSQL = sSQL & "ARTIDEPO_STOCKMAXIMO = " & Replace(Depositos.Item(i).StockMaximo, ",", ".") & ", "
                sSQL = sSQL & "ARTIDEPO_STOCKREPOSI = " & Replace(Depositos.Item(i).StockReposi, ",", ".") & ", "
                sSQL = sSQL & "ARTIDEPO_USUARIO = '" & Arti.Usuario & "',"
                sSQL = sSQL & "ARTIDEPO_FECPRO = " & ConvertirFecha(Date) & ","
                sSQL = sSQL & "ARTIDEPO_HORPRO = '" & Format(Now, "hh:mm:ss") & "'"
                sSQL = sSQL & " WHERE ARTIDEPO_IDREG = " & rs!ARTIDEPO_IDREG
            
            End If
            cn.Execute sSQL, lRes
            If lRes = 0 Then GoTo cnError
        
        End If
        rs.Close
        
    Next
    
    sSQL = "DELETE FROM ARTICULO_DEPOSITOS WHERE ARTIDEPO_USUARIO = 'BORRAR' AND ARTIDEPO_ARTICULO = " & Arti.idReg
    cn.Execute sSQL
    
    MsgBox "El artículo " & Arti.Descripcion & " ha sido modificado con exito.", vbOKOnly + vbInformation, "Incorporación de artículos ..."
    cn.CommitTrans
    GrabarModificacion = True

    cn.Close
    Set cn = Nothing

Exit Function

errorsql:
    MsgBox "Ha ocurrido un error al intentar modificar el artículo." & vbCrLf & "Comuniquese con mantenimiento.", vbOKOnly + vbCritical, "Fallo en el proceso de ingresos..."
    GuardarErrorTransacciones cn.Errors, "al crear el artículo - [SGTRARTICULOS]"
    cn.RollbackTrans
    cn.Close
    Set cn = Nothing
    GrabarModificacion = False

    Exit Function

cnError:
    MostrarErrorDeTransacciones cn.Errors, "al modificar un artículo"
    cn.RollbackTrans
    cn.Close
    Set cn = Nothing
    GrabarModificacion = False

End Function
Public Function MostrarInfo(ByRef Info As SGGrid, ByVal idReg As Long) As Boolean

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim trRenglo As cTRInfoRenglon
    Dim iNdx As Integer
    Dim sSQL As String
    Dim trInformacion As cTRInfoRenglones
    Dim tmpA As SGArray

    Set trInformacion = New cTRInfoRenglones

    trInformacion.Add "Articulo: ", "ARTI_CODIGO"
    trInformacion.Add "Descripción: ", "ARTI_DESCRIPCION"
    trInformacion.Add "Unidad Medida Compra: ", "ARTI_UNIMEDIDACMP"
    trInformacion.Add "Línea: ", "LINEA_DESCRI"
    trInformacion.Add "Grupo: ", "GRUCMP_NOMBRE"
    trInformacion.Add "Proveedor: ", "PROVNOMBRE"
    trInformacion.Add "Gravado: ", "GRAV_NOMBRE"
    trInformacion.Add "Marca: ", "ARTI_MARCA"
    trInformacion.Add "Precio: ", "ARTI_PRELIS"
    trInformacion.Add "Vigencia: ", "ARTI_VIGECOST_F"

    Info.Columns.RemoveAll True
    Info.DataRowCount = trInformacion.Count
    Info.DataColCount = 2

    sSQL = "SELECT * FROM vw_TrArticulos WHERE ARTI_IDREG = " & idReg

    Set cn = New ADODB.Connection
    cn.Open user.sgEmpresas.Seleccionada.EMPRCONEXION
    Set rs = New ADODB.Recordset
    rs.Open sSQL, cn, adOpenStatic, adLockReadOnly, adCmdText

    iNdx = 0

    Set tmpA = Info.Array
    For Each trRenglo In trInformacion
        tmpA.Set iNdx, 0, trRenglo.Titulo
        If Right(trRenglo.Campo, 1) = "F" Then
            tmpA.Set iNdx, 1, IIf(IsNull(rs.Fields(Left(trRenglo.Campo, Len(trRenglo.Campo) - 2)).value), "N/D", ConvertirFecha(rs.Fields(Left(trRenglo.Campo, Len(trRenglo.Campo) - 2)).value))
        Else
            tmpA.Set iNdx, 1, IIf(IsNull(rs.Fields(trRenglo.Campo).value), "N/D", rs.Fields(trRenglo.Campo).value)
        End If
        iNdx = iNdx + 1
    Next

    Info.Height = Info.Rows.At(1).Height * (trInformacion.Count + 1)
    Info.FitLastColumn = True

    MostrarInfo = True

    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing

End Function

Public Sub Imprimir(ByVal lstProc As ListView)

'    Load frmCPImprimirArticulos
'
'    frmCPImprimirArticulos.TipoBusquedaFiltro = sgTexto
'    frmCPImprimirArticulos.KeyBusquedaFiltro = 0
'    frmCPImprimirArticulos.CargarCriterios ArmarCriterios
'    frmCPImprimirArticulos.Orden = "Ordenamiento"
'    PasarListViewChecked lstProc, frmCPImprimirArticulos.lstSeleccion
'
'    If frmCPImprimirArticulos.lstSeleccion.ListItems.Count = 0 Then
'        frmCPImprimirArticulos.chkSeleccion.Enabled = False
'        frmCPImprimirArticulos.lstSeleccion.Visible = False
'        frmCPImprimirArticulos.chkCriterios.value = True
'    Else
'        frmCPImprimirArticulos.chkSeleccion.Enabled = True
'        frmCPImprimirArticulos.chkSeleccion.value = True
'        frmCPImprimirArticulos.lstSeleccion.Visible = True
'        frmCPImprimirArticulos.chkSeleccion.value = True
'    End If
'
'    frmCPImprimirArticulos.Show vbModal

End Sub

Public Function Eliminar() As Boolean

    On Error Resume Next

    Eliminar = False
    Dim cn As ADODB.Connection
    Dim lRes As Long
    Dim sSQL As String

    On Error GoTo errorsql

    If MsgBox("Desea eliminar el artículo " & frmTratArticulos.sgGrilla.Rows.Current.Cells(2) & " - " _
        & frmTratArticulos.sgGrilla.Rows.Current.Cells(3) & "?", vbYesNo, "Atención") = vbYes Then

        Set cn = New ADODB.Connection
        cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
        cn.BeginTrans

        sSQL = "DELETE FROM ARTICULO_CATALOGOS WHERE ARTICATA_ARTICULO = " & frmTratArticulos.sgGrilla.Rows.Current.Cells(11)
        cn.Execute sSQL, lRes

        sSQL = "DELETE FROM ARTICULO_DEPOSITOS WHERE ARTIDEPO_ARTICULO = " & frmTratArticulos.sgGrilla.Rows.Current.Cells(11)
        cn.Execute sSQL, lRes
        
        sSQL = "DELETE FROM ARTICULOS WHERE ARTI_IDREG = " & frmTratArticulos.sgGrilla.Rows.Current.Cells(11)
        cn.Execute sSQL, lRes

        MsgBox "Se ha eliminado con exito el artículo ...", vbOKOnly + vbInformation, "Eliminar..."

        cn.CommitTrans
        cn.Close
        Set cn = Nothing

    End If
    Exit Function

errorsql:

    MsgBox "No se ha podido eliminar el artículo " & err.Description, vbOKOnly + vbInformation, "Eliminar..."
    Set cn = Nothing

End Function
Public Function EliminarMed() As Boolean

    On Error Resume Next

    EliminarMed = False
    Dim cn As ADODB.Connection
    Dim lRes As Long
    Dim sSQL As String

    On Error GoTo errorsql

    If MsgBox("Desea eliminar el artículo " & frmTratMedicamentos.sgGrilla.Rows.Current.Cells(2) & " - " _
        & frmTratMedicamentos.sgGrilla.Rows.Current.Cells(3) & "?", vbYesNo, "Atención") = vbYes Then

        Set cn = New ADODB.Connection
        cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
        cn.BeginTrans

        sSQL = "DELETE FROM ARTICULO_CATALOGOS WHERE ARTICATA_ARTICULO = " & frmTratMedicamentos.sgGrilla.Rows.Current.Cells(10)
        cn.Execute sSQL, lRes

        sSQL = "DELETE FROM ARTICULO_DEPOSITOS WHERE ARTIDEPO_ARTICULO = " & frmTratMedicamentos.sgGrilla.Rows.Current.Cells(10)
        cn.Execute sSQL, lRes
        
        sSQL = "DELETE FROM ARTICULOS WHERE ARTI_IDREG = " & frmTratMedicamentos.sgGrilla.Rows.Current.Cells(10)
        cn.Execute sSQL, lRes

        MsgBox "Se ha eliminado con exito el artículo ...", vbOKOnly + vbInformation, "Eliminar..."

        cn.CommitTrans
        cn.Close
        Set cn = Nothing

    End If
    Exit Function

errorsql:

    MsgBox "No se ha podido eliminar el artículo " & err.Description, vbOKOnly + vbInformation, "Eliminar..."
    Set cn = Nothing

End Function
Public Function EliminarProdMed() As Boolean

    On Error Resume Next

    EliminarProdMed = False
    Dim cn As ADODB.Connection
    Dim lRes As Long
    Dim sSQL As String

    On Error GoTo errorsql

    If MsgBox("Desea eliminar el artículo " & frmTratProdMedicos.sgGrilla.Rows.Current.Cells(2) & " - " _
        & frmTratProdMedicos.sgGrilla.Rows.Current.Cells(3) & "?", vbYesNo, "Atención") = vbYes Then

        Set cn = New ADODB.Connection
        cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
        cn.BeginTrans

        sSQL = "DELETE FROM ARTICULO_CATALOGOS WHERE ARTICATA_ARTICULO = " & frmTratProdMedicos.sgGrilla.Rows.Current.Cells(10)
        cn.Execute sSQL, lRes

        sSQL = "DELETE FROM ARTICULO_DEPOSITOS WHERE ARTIDEPO_ARTICULO = " & frmTratProdMedicos.sgGrilla.Rows.Current.Cells(10)
        cn.Execute sSQL, lRes
        
        sSQL = "DELETE FROM ARTICULOS WHERE ARTI_IDREG = " & frmTratProdMedicos.sgGrilla.Rows.Current.Cells(10)
        cn.Execute sSQL, lRes

        MsgBox "Se ha eliminado con exito el artículo ...", vbOKOnly + vbInformation, "Eliminar..."

        cn.CommitTrans
        cn.Close
        Set cn = Nothing

    End If
    Exit Function

errorsql:

    MsgBox "No se ha podido eliminar el artículo " & err.Description, vbOKOnly + vbInformation, "Eliminar..."
    Set cn = Nothing

End Function

Public Function EliminarVarios() As Boolean

    On Error Resume Next

    EliminarVarios = False
    Dim cn As ADODB.Connection
    Dim lRes As Long
    Dim sSQL As String

    On Error GoTo errorsql

    If MsgBox("Desea eliminar el artículo " & frmTratVarios.sgGrilla.Rows.Current.Cells(2).Text & " - " _
        & frmTratVarios.sgGrilla.Rows.Current.Cells(3).Text & "?", vbYesNo, "Atención") = vbYes Then

        Set cn = New ADODB.Connection
        cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
        cn.BeginTrans

        sSQL = "DELETE FROM ARTICULO_CATALOGOS WHERE ARTICATA_ARTICULO = " & frmTratVarios.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes

        sSQL = "DELETE FROM ARTICULO_DEPOSITOS WHERE ARTIDEPO_ARTICULO = " & frmTratVarios.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes
        
        sSQL = "DELETE FROM ARTICULOS WHERE ARTI_IDREG = " & frmTratVarios.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes

        MsgBox "Se ha eliminado con exito el artículo ...", vbOKOnly + vbInformation, "Eliminar..."

        cn.CommitTrans
        cn.Close
        Set cn = Nothing

    End If
    Exit Function

errorsql:

    MsgBox "No se ha podido eliminar el artículo " & err.Description, vbOKOnly + vbInformation, "Eliminar..."
    Set cn = Nothing

End Function

Public Function EliminarVariosAlimentacion() As Boolean

    On Error Resume Next

    EliminarVariosAlimentacion = False
    Dim cn As ADODB.Connection
    Dim lRes As Long
    Dim sSQL As String

    On Error GoTo errorsql

    If MsgBox("Desea eliminar el artículo " & frmTratVariosAlimentacion.sgGrilla.Rows.Current.Cells(2).Text & " - " _
        & frmTratVariosAlimentacion.sgGrilla.Rows.Current.Cells(3).Text & "?", vbYesNo, "Atención") = vbYes Then

        Set cn = New ADODB.Connection
        cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
        cn.BeginTrans

        sSQL = "DELETE FROM ARTICULO_CATALOGOS WHERE ARTICATA_ARTICULO = " & frmTratVariosAlimentacion.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes

        sSQL = "DELETE FROM ARTICULO_DEPOSITOS WHERE ARTIDEPO_ARTICULO = " & frmTratVariosAlimentacion.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes
        
        sSQL = "DELETE FROM ARTICULOS WHERE ARTI_IDREG = " & frmTratVariosAlimentacion.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes

        MsgBox "Se ha eliminado con exito el artículo ...", vbOKOnly + vbInformation, "Eliminar..."

        cn.CommitTrans
        cn.Close
        Set cn = Nothing

    End If
    Exit Function

errorsql:

    MsgBox "No se ha podido eliminar el artículo " & err.Description, vbOKOnly + vbInformation, "Eliminar..."
    Set cn = Nothing

End Function

Public Function EliminarVarios_ServGrales() As Boolean

    On Error Resume Next

    EliminarVarios_ServGrales = False
    Dim cn As ADODB.Connection
    Dim lRes As Long
    Dim sSQL As String

    On Error GoTo errorsql

    If MsgBox("Desea eliminar el artículo " & frmTratVariosServGrales.sgGrilla.Rows.Current.Cells(2).Text & " - " _
        & frmTratVariosServGrales.sgGrilla.Rows.Current.Cells(3).Text & "?", vbYesNo, "Atención") = vbYes Then

        Set cn = New ADODB.Connection
        cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
        cn.BeginTrans

        sSQL = "DELETE FROM ARTICULO_CATALOGOS WHERE ARTICATA_ARTICULO = " & frmTratVariosServGrales.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes

        sSQL = "DELETE FROM ARTICULO_DEPOSITOS WHERE ARTIDEPO_ARTICULO = " & frmTratVariosServGrales.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes
        
        sSQL = "DELETE FROM ARTICULOS WHERE ARTI_IDREG = " & frmTratVariosServGrales.sgGrilla.Rows.Current.Cells(11).value
        cn.Execute sSQL, lRes

        MsgBox "Se ha eliminado con exito el artículo ...", vbOKOnly + vbInformation, "Eliminar..."

        cn.CommitTrans
        cn.Close
        Set cn = Nothing

    End If
    Exit Function

errorsql:

    MsgBox "No se ha podido eliminar el artículo " & err.Description, vbOKOnly + vbInformation, "Eliminar..."
    Set cn = Nothing

End Function
Public Function ExisteArticulo(ByVal Codigo As String, Optional ByVal idReg As Long)
                                
    On Error Resume Next

    ExisteArticulo = False

    Dim lResultado As Long
    Dim sSQL As String
    
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If idReg > 0 Then
        sSQL = "Select ARTI_CODIGO FROM ARTICULOS WHERE "
        sSQL = sSQL & " UPPER(ARTI_CODIGO) = '" & UCase(Codigo) & "' AND "
        sSQL = sSQL & " ARTI_IDREG <> " & idReg
    Else
        sSQL = "Select ARTI_CODIGO FROM ARTICULOS WHERE "
        sSQL = sSQL & " UPPER(ARTI_CODIGO) = '" & UCase(Codigo) & "'"
    End If

    cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic, adCmdText

    If rs.EOF Then
        ExisteArticulo = False
    Else
        ExisteArticulo = True
    End If
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing

End Function

Public Function BuscarCodigoDisponible(ByVal mEspacio As Long) As String
                                
    On Error Resume Next

    Dim sSQL As String
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sSQL = "Select MAX(ARTI_CODIGO) AS CODIGO FROM ARTICULOS WHERE "
    sSQL = sSQL & " ARTI_ESPACIO = " & IIf(mEspacio > 2, 3, mEspacio)

    cn.Open user.sgEmpresas.Seleccionada.EMPRAUXILIAR
    rs.Open sSQL, cn, adOpenDynamic, adLockOptimistic, adCmdText

    Select Case mEspacio
        Case Is = 1
            BuscarCodigoDisponible = Val(rs!Codigo) + 1
        Case Is = 2
            BuscarCodigoDisponible = "S" & Val(Mid(rs!Codigo, 2)) + 1
        Case Is > 2
            BuscarCodigoDisponible = "V" & Val(Mid(rs!Codigo, 2)) + 1
    End Select

    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing

End Function



