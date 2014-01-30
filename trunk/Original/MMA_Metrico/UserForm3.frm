VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Algoritmo Map-matching Mejorado"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13890
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









'Algoritmo Mejorado

Dim pMxDoc As IMxDocument       'Este Documento
Dim pMaps As IMap               'Este Mapa
Dim pApp As IApplication        'Esta Aplicacion
Dim pLayer As IFeatureLayer     'Layer de Arcos
Dim pLayer2 As IFeatureLayer    'Layer de Puntos
Dim K1 As IFeature              'Punto K1
Dim K2 As IFeature              'Punto K2

Dim pNAExtension As INetworkAnalystExtension  'Application extension for NA
Dim pNAWindow As INAWindow                    'Network Analysis Window
Dim pNALayer As INALayer                      'Active Network Analysis Layer
Dim pNAContext As INAContext                  'Active NALayer's context
Dim pNAWindowCategory As INAWindowCategory    'Stops, barriers, incidents, etc
Dim pGPMessages As IGPMessages

Dim pNAClass As INAClass
Dim pClasses As INamedSet
Dim pQFilt As IQueryFilter
Dim strquery As String

Dim pPolygon As IPolygon
Dim pSpatialFilter As ISpatialFilter
Dim pFeatureSelection As IFeatureSelection
Dim pTopoOp As ITopologicalOperator
Dim pEnumIDs As IEnumIDs
Dim id As Long

Dim RouteLayer As INALayer
Dim pRouteLayer As IFeatureLayer
Dim rID As Long

Dim pfeaturelayer As IFeatureLayer
Dim pFeatureClass As IFeatureClass
Dim pDataset As IDataset
Dim pWorkspace As IWorkspace
Dim pWorkspaceEdit As IWorkspaceEdit
Dim ptable As ITable
Dim pNewFeature As IFeature

Dim rBuffer As Double
Dim NK As Long
Dim NA As Integer
Dim posSpeed As Integer
Dim posTime As Integer
Dim V As Double     'Velocidad Promedio Capturada
Dim S As Double     'Velocidad Promedio Calculada
Dim D As Double     'Distancia entre puntos
Dim i As Long       'Posicion origen
Dim j As Long       'Posicion destino
Dim velUmb As Double    'Velocidad Umbral Minima
Dim RutaFull As Boolean
Dim MtrxArc() As Long           'Matriz de Arcos
Dim MtrxDist() As Double        'Matriz de Distancias
Dim MtrxDescarte() As Boolean   'Matriz de Descarte
Dim ArrPrevMin() As Double      'Arreglo de distancia previa minima (para no buscar el mismo arcos candidato)
Dim MtrxPto() As IPoint         'Matriz de Puntos (Puntos Snaps en los arcos candidatos)
Dim ArrArcRuta() As Long
Dim Aceptados() As IPoint       'Arreglo de puntos aceptados
Dim Callesi() As String
Dim Callesj() As String
Dim Calles() As String

Dim ptoKi() As Long
Dim ptoKj() As Long
Dim auxID1 As Long
Dim auxID2 As Long
Dim spID1 As Long
Dim spID2 As Long
Dim ArID() As Long
Dim altk As Long
Dim Min As Double
Dim MinAlt As Boolean
Dim tpo As Long
Dim Tiempo As String
Dim timer1 As Date
Dim conttpo As String
Dim t1 As Date
Dim t2 As Date

Dim DeltaT() As Double
Dim VelocidadesV() As Double
Dim VelocidadesS() As Double
Dim DistanciaD() As Double
Dim DistanciaDA() As Double
Dim IDArcoAsociadoi() As Long
Dim IDArcoAsociadoj() As Long
Dim ArrAzimutPto() As Double
Dim ArrAzimutCalle() As Double

Dim SnapPtoAct As IPoint
Dim distMay As Double
Dim ContinuarBusqueda As Boolean
Dim DeltaBuffer As Integer
Dim multiplicador As Integer
Dim numCellHeading As Integer
Dim arcid As Long

Dim avance As Boolean       'Variables para MMA mejorado
Dim numPuntos As Integer
Dim num_i As Long
Dim num_j As Long
Dim difij As Integer
Dim rutag As Boolean
Dim rut_inviable As Boolean

    

Private Sub CheckBox1_Click()
    If CheckBox1.Value = False Then
        CheckBox1.Value = True
        TextBox13.BackColor = &H8000000F
        TextBox13.Enabled = False
    Else
        CheckBox1.Value = False
        TextBox13.BackColor = &H80000005
        TextBox13.Enabled = True
    End If
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2.Value = False Then
        CheckBox2.Value = True
        TextBox14.BackColor = &H8000000F
        TextBox14.Enabled = False
    Else
        CheckBox2.Value = False
        TextBox14.BackColor = &H80000005
        TextBox14.Enabled = True
    End If
End Sub

Private Sub CheckBox3_Click()
    If CheckBox3.Value = False Then
        CheckBox3.Value = True
        TextBox15.BackColor = &H8000000F
        TextBox15.Enabled = False
        TextBox17.BackColor = &H8000000F
        TextBox17.Enabled = False
    Else
        CheckBox3.Value = False
        TextBox15.BackColor = &H80000005
        TextBox15.Enabled = True
        TextBox17.BackColor = &H80000005
        TextBox17.Enabled = True
    End If
End Sub

Private Sub CheckBox4_Click()
    If CheckBox4.Value = False Then
        CheckBox4.Value = True
        TextBox16.BackColor = &H8000000F
        TextBox16.Enabled = False
    Else
        CheckBox4.Value = False
        TextBox16.BackColor = &H80000005
        TextBox16.Enabled = True
    End If
End Sub

Private Sub CommandButton1_Click()
    timer1 = Now
    Label25.Caption = "00:00:00:00"
    If CheckBox4.Value = False Then
        Call EjecucionAlgoritmo
    End If
    
    If CheckBox4.Value = True Then
        Call EjecucionAlgoritmoMejorado
    End If
    
End Sub

Public Sub EjecucionAlgoritmo()
Set pMxDoc = ThisDocument
Set pMaps = pMxDoc.FocusMap
rBuffer = CDbl(TextBox4.Text)
NA = 4
RutaFull = False
ContinuarBusqueda = True
rutag = True

'Se comienza la edicion para generar la Layer Resultado
Call ComienzaEdicion

If TextBox1.Text = "" Then
    MsgBox "Ingrese el numero de ID de la layer de puntos GPS"
Else
    Set pLayer2 = pMaps.Layer(CInt(TextBox1.Text)) 'se le asigna a player a la layer dentro del pmaps segun su numero de id
    Set pLayer = pMaps.Layer(CInt(TextBox3.Text))
    
    'Limpiar Layer Resultado
    If Not TypeOf pMaps.Layer(CInt(TextBox6.Text)) Is IFeatureLayer Then
        MsgBox "No existe Layer Resultado, Porfavor Carguela o Creela"
        Exit Sub
    Else
        Call LimpiarResultados
    End If
        
        


    If TextBox13.Text = "" Then  'se asigna velocidad umbral minima = 0 si no se ingresa como dato
    velUmb = 0
    Else
    velUmb = TextBox13.Text      'se registra la informacion de la velocidad umbral minima
    End If

    posSpeed = pLayer2.FeatureClass.FindField(TextBox10.Text) 'posSpeed toma la posicion de la columna SPEED (velocidad)
    If OptionButton1.Value = True Then
        posTime = pLayer2.FeatureClass.FindField(TextBox11.Text) 'posTime toma el valor de la posicion para la columna TIME (hora)
    End If
    NK = pLayer2.FeatureClass.FeatureCount(Nothing)
    
    'Definir las dimensiones de las matrices y arreglo a utilizar
    ReDim MtrxArc(NK, NA) As Long
    ReDim MtrxDist(NK, NA) As Double
    ReDim ArrPrevMin(NK) As Double
    ReDim MtrxPto(NK, NA) As IPoint
    ReDim MtrxDescarte(NK, NA) As Boolean
    ReDim Aceptados(NK) As IPoint
    ReDim Callesi(NK) As String
    ReDim Callesj(NK) As String
    ReDim Calles(NK) As String
    ReDim ArID(NK) As Long
    
    'Arreglos con Datos estadisticos
    ReDim VelocidadesV(NK) As Double
    ReDim VelocidadesS(NK) As Double
    ReDim DistanciaD(NK) As Double
    ReDim DistanciaDA(NK) As Double
    ReDim DeltaT(NK) As Double
    ReDim ptoKi(NK) As Long
    ReDim ptoKj(NK) As Long
    ReDim IDArcoAsociadoi(NK) As Long
    ReDim IDArcoAsociadoj(NK) As Long
    ReDim ArrAzimutPto(NK) As Double
    ReDim ArrAzimutCalle(NK) As Double

    'Seleccionar Arcos en el buffer y llena las matrices y arreglo con los datos necesarios
    If CheckBox2.Value = True Then
        For i = 0 To NK - 1
            Call SelectionBufferDinamicoPrev(pLayer2.FeatureClass.GetFeature(i), rBuffer, pLayer)
            Frame12.Repaint
            Cronometro
        Next
    Else
        For i = 0 To NK - 1
            Call SelectionBuffer(pLayer2.FeatureClass.GetFeature(i), rBuffer, pLayer)
            Frame12.Repaint
            Cronometro
        Next
    End If
    Dim arcid As Long
    Label17.Caption = NK
    Label19.Caption = pLayer.FeatureClass.FeatureCount(Nothing)
    Label23.Caption = NA
    DistanciaD(0) = 0
    DistanciaDA(0) = 0
    
    'Ejecucion del Algoritmo y creacion de layer de resultado
    For i = 0 To NK - 1 'hasta n-3
        Label21.Caption = i
        BusquedaAlt = True
        multiplicador = 1
        rut_inviable = False
        
        If i > 1 Then
            ArrPrevMin(i - 2) = -1
        End If
        If i > 0 Then
            ArrPrevMin(i - 1) = -1
        End If
        ArrPrevMin(i) = -1
        If i < NK - 2 Then
            ArrPrevMin(i + 1) = -1
        End If
        If i < NK - 3 Then
            ArrPrevMin(i + 2) = -1
        End If
        
        If i > 1 And i < NK - 2 Then
            For j = i - 2 To i + 2
                arcid = 0
                While MtrxDist(j, arcid) <> -1 And arcid < NA And i <> j
                    MtrxDescarte(j, arcid) = True
                    arcid = arcid + 1
                Wend
            Next
        End If
        
        If i < 2 Then
            For j = 0 To i + 2
                arcid = 0
                While MtrxDist(j, arcid) <> -1 And arcid < NA And i <> j
                    MtrxDescarte(j, arcid) = True
                    arcid = arcid + 1
                Wend
            Next
        End If
        
        If i > NK - 3 Then
            For j = i - 2 To NK - 1
                arcid = 0
                While MtrxDist(j, arcid) <> -1 And arcid < NA And i <> j
                    MtrxDescarte(j, arcid) = True
                    arcid = arcid + 1
                Wend
            Next
        End If
    
        
        If i > 1 And i < NK - 2 Then
             If Not ExisteCamino(i, i + 1) Then
                
                If ExisteCamino(i + 1, i + 2) Then
        
                    If ExisteAlt(i) Then
                        CaminoAlternativo (i)
                    Else
                        If ExisteCamino(i - 1, i + 1) Then
                            Call AceptarPuntos(i - 1, i + 1) 'DUDA ACEPTAR Los 2 PUNTOS o TODOS los Puntos entre K1 y K2
                        Else
                            If ExisteAlt(i + 1) Then
                                CaminoAlternativo (i + 1)
                            Else
                                If ExisteCamino(i, i + 2) Then
                                    Call AceptarPuntos(i, i + 2)
                                Else
                                    If ExisteAlt(i - 1) Then
                                        CaminoAlternativo (i - 1)
                                    Else
                                        If ExisteCamino(i - 2, i + 1) Then
                                            Call AceptarPuntos(i - 2, i + 1)
                                        Else
                                            'MsgBox "Ambigüedad no resuelta para ruta entre puntos " & i - 2 & " y " & i + 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If ExisteAlt(i + 1) Then
                        CaminoAlternativo (i + 1)
                    Else
                        If ExisteCamino(i, i + 2) Then
                            Call AceptarPuntos(i, i + 2)
                        Else
                            If ExisteAlt(i) Then
                                CaminoAlternativo (i)
                            Else
                                If ExisteCamino(i - 1, i + 1) Then
                                    Call AceptarPuntos(i - 1, i + 1)
                                Else
                                    If ExisteAlt(i - 1) Then
                                        CaminoAlternativo (i - 1)
                                    Else
                                        If ExisteCamino(i - 2, i + 1) Then
                                            Call AceptarPuntos(i - 2, i + 1)
                                        Else
                                            'MsgBox "Ambigüedad no resuelta para ruta entre puntos " & i - 2 & " y " & i + 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                Call AceptarPuntos(i, i + 1)
            End If
        End If
        
        'Casos bordes (i<2)
        If i < 2 Then
            If ExisteCamino(i, i + 1) Then
                Call AceptarPuntos(i, i + 1)
            Else
                If ExisteCamino(i + 1, i + 2) Then
                    If i > 0 Then
                        If ExisteAlt(i) Then
                            CaminoAlternativo (i)
                        Else
                            If ExisteCamino(i - 1, i + 1) Then
                                Call AceptarPuntos(i - 1, i + 1)
                            Else
                                If ExisteAlt(i + 1) Then
                                    CaminoAlternativo (i + 1)
                                Else
                                    If ExisteCamino(i, i + 2) Then
                                        Call AceptarPuntos(i, i + 2)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If ExisteAlt(i + 1) Then
                        CaminoAlternativo (i + 1)
                    Else
                        If ExisteCamino(i, i + 2) Then
                            Call AceptarPuntos(i, i + 2)
                        Else
                            If i > 0 Then
                                If ExisteAlt(i) Then
                                    CaminoAlternativo (i)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'Casos bordes (i>NK-3)
        If i > NK - 3 And i < NK - 1 Then
                If ExisteCamino(i, i + 1) Then
                    Call AceptarPuntos(i, i + 1)
                Else
                    If ExisteAlt(i) Then
                        CaminoAlternativo (i)
                    Else
                        If ExisteCamino(i - 1, i + 1) Then
                            Call AceptarPuntos(i - 1, i + 1)
                        Else
                            If ExisteAlt(i - 1) Then
                                CaminoAlternativo (i - 1)
                            Else
                                If ExisteCamino(i - 2, i + 1) Then
                                    Call AceptarPuntos(i - 2, i + 1)
                                End If
                            End If
                        End If
                    End If
                End If
        End If
        Frame12.Repaint
        Cronometro
    Next
    
    
    'Finaliza la edicion y se entrega los resultados del algoritmo
    For i = 0 To NK - 1
        Call CreatePointFeature(Aceptados(i), Calles(i))
    Next
    Call TerminaEdicion
    If rutag Then 'En Caso de tener tramos infactibles no genera la ruta graficamente
        RutaFull = True
        Call EfectuarRuta(pfeaturelayer, 0, 0, 0)
    End If

    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
    MsgBox "Algoritmo Terminado con Exito!"
    

    GenerarResultados
    'Reporte


End If
End Sub

Public Sub EjecucionAlgoritmoMejorado()
Set pMxDoc = ThisDocument
Set pMaps = pMxDoc.FocusMap
rBuffer = CDbl(TextBox4.Text)
NA = 4
RutaFull = False
ContinuarBusqueda = True
rutag = True

'Se comienza la edicion para generar la Layer Resultado
Call ComienzaEdicion

If TextBox1.Text = "" Then
    MsgBox "Ingrese el numero de ID de la layer de puntos GPS"
Else
    Set pLayer2 = pMaps.Layer(CInt(TextBox1.Text)) 'se le asigna a player a la layer dentro del pmaps segun su numero de id
    Set pLayer = pMaps.Layer(CInt(TextBox3.Text))
    
    'Limpiar Layer Resultado
    If Not TypeOf pMaps.Layer(CInt(TextBox6.Text)) Is IFeatureLayer Then
        MsgBox "No existe Layer Resultado, Porfavor Carguela o Creela"
        Exit Sub
    Else
        Call LimpiarResultados
    End If
        

    If TextBox13.Text = "" Then  'se asigna velocidad umbral minima = 0 si no se ingresa como dato
    velUmb = 0
    Else
    velUmb = TextBox13.Text      'se registra la informacion de la velocidad umbral minima
    End If

    posSpeed = pLayer2.FeatureClass.FindField(TextBox10.Text) 'posSpeed toma la posicion de la columna SPEED (velocidad)
    If OptionButton1.Value = True Then
        posTime = pLayer2.FeatureClass.FindField(TextBox11.Text) 'posTime toma el valor de la posicion para la columna TIME (hora)
    End If
    NK = pLayer2.FeatureClass.FeatureCount(Nothing)
    
    'Definir las dimensiones de las matrices y arreglo a utilizar
    ReDim MtrxArc(NK, NA) As Long
    ReDim MtrxDist(NK, NA) As Double
    ReDim ArrPrevMin(NK) As Double
    ReDim MtrxPto(NK, NA) As IPoint
    ReDim MtrxDescarte(NK, NA) As Boolean
    ReDim Aceptados(NK) As IPoint
    ReDim Callesi(NK) As String
    ReDim Callesj(NK) As String
    ReDim Calles(NK) As String
    ReDim ArID(NK) As Long
    
    'Arreglos con Datos estadisticos
    ReDim VelocidadesV(NK) As Double
    ReDim VelocidadesS(NK) As Double
    ReDim DistanciaD(NK) As Double
    ReDim DistanciaDA(NK) As Double
    ReDim DeltaT(NK) As Double
    ReDim ptoKi(NK) As Long
    ReDim ptoKj(NK) As Long
    ReDim IDArcoAsociadoi(NK) As Long
    ReDim IDArcoAsociadoj(NK) As Long
    ReDim ArrAzimutPto(NK) As Double
    ReDim ArrAzimutCalle(NK) As Double

    'Seleccionar Arcos en el buffer y llena las matrices y arreglo con los datos necesarios
    If CheckBox2.Value = True Then
        For i = 0 To NK - 1
            Call SelectionBufferDinamicoPrev(pLayer2.FeatureClass.GetFeature(i), rBuffer, pLayer)
            Frame12.Repaint
            Cronometro
        Next
    Else
        For i = 0 To NK - 1
            Call SelectionBuffer(pLayer2.FeatureClass.GetFeature(i), rBuffer, pLayer)
            Frame12.Repaint
            Cronometro
        Next
    End If
    
    Label17.Caption = NK
    Label19.Caption = pLayer.FeatureClass.FeatureCount(Nothing)
    Label23.Caption = NA
    

    numPuntos = CInt(TextBox16.Text)
    DistanciaD(0) = 0
    DistanciaDA(0) = 0
    
    'Ejecucion del Algoritmo y creacion de layer de resultado
    For i = 0 To NK - 1
        Label21.Caption = i
        BusquedaAlt = True
        multiplicador = 1
        avance = False
        rut_inviable = False
 
        num_i = i
        num_j = i + 1
        difij = num_j - num_i
        
       'NUEVOS PASOS
       While numPuntos >= difij And num_i >= 0 And num_j < NK
            If Not ExisteCamino(num_i, num_j) Then
                If difij < 99 Then
                    If ExisteAlt(num_j) Then
                        Frame12.Repaint
                        Cronometro
                    Else
                        If ExisteAlt(num_i) Then
                            Frame12.Repaint
                            Cronometro
                            quitaDescarte (num_j)
                            CaminoAlternativo (num_i)
                        Else
                            Cronometro
                            quitaDescarte (num_j)
                            quitaDescarte (num_i)
                            If avance = False Then
                                num_i = num_i - 1
                                difij = num_j - num_i
                                avance = True
                            Else
                                num_j = num_j + 1
                                difij = num_j - num_i
                                avance = False
                            End If
                        End If
                    End If
                End If
            Else
                AceptarPuntos num_i, num_j
                difij = difij + 100
                i = num_j - 1
                Cronometro
            End If
       Wend
        Frame12.Repaint
        Cronometro
    Next
    
    
    'Finaliza la edicion y se entrega los resultados del algoritmo
    For i = 0 To NK - 1
        Call CreatePointFeature(Aceptados(i), Calles(i))
    Next
    Call TerminaEdicion
    If rutag Then 'En caso de tramos infactibles no genera ruta graficamente
        RutaFull = True
        Call EfectuarRuta(pfeaturelayer, 0, 0, 0)
    End If
    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
    MsgBox "Algoritmo Terminado con Exito!"
    

    GenerarResultados
    'Reporte


End If
End Sub

Public Sub quitaDescarte(ByVal idptodesc As Integer)
            'Utilizada para limpiar los arcos candidatos para un punto
            arcid = 0
            ArrPrevMin(idptodesc) = -1
            While MtrxDist(idptodesc, arcid) <> -1 And arcid < NA
                MtrxDescarte(idptodesc, arcid) = True
                arcid = arcid + 1
            Wend
End Sub
Public Sub Reporte()
    Dim aciertos As Long
    Dim porcentaje As Double
    Dim ubicacion As String
    Dim nom_archivo As String
    aciertos = 0
    ubicacion = TextBox8.Text
    nom_archivo = TextBox9.Text
    
    
    'Variable de tipo Aplicación de Excel
    Dim objExcel As Excel.Application
    
    'Una variable de tipo Libro de Excel
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, Fila As Integer
    
    'creamos un nuevo objeto excel
    Set objExcel = New Excel.Application

    'Usamos el método open para abrir el archivo que está _
    en el directorio del programa llamado Result_Algoritmo.xls
    Set xLibro = objExcel.Workbooks.Open(ubicacion + nom_archivo + ".xls")
    
    'Hacemos el Excel Visible
    objExcel.Visible = True
    
         With xLibro
    
             ' Hacemos referencia a la Hoja
             With .Sheets(1)
    
                 'Recorremos las filas
                 For Fila = 1 To NK
    
                     'Se compara la segunda columna (nombre de calles)
                     If Calles(Fila - 1) = .cells(Fila + 1, 1) Then
                        aciertos = aciertos + 1
                     End If
                     .cells(Fila + 1, 2) = Calles(Fila - 1)
                 Next
                 porcentaje = aciertos / NK
                 .cells(2, 3) = porcentaje * 100
    
             End With
         End With
    
         'Eliminamos los objetos si ya no los usamos
         Set objExcel = Nothing
         Set xLibro = Nothing
         
         MsgBox "Se ha logrado un " & porcentaje * 100 & "% de Exito. Para detalles revisar el archivo Result_Algoritmo.xls"
         
         
    
End Sub

Public Sub GenerarResultados()
    Dim ubicacion As String
    ubicacion = TextBox8.Text
    
    
    'Variable de tipo Aplicación de Excel
    Dim objExcel As Excel.Application
    
    'Una variable de tipo Libro de Excel
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, Fila As Integer
    
    'creamos un nuevo objeto excel
    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Add

    'Usamos el método open para abrir el archivo que está _
    en el directorio del programa llamado Result_Algoritmo.xls
    xLibro.SaveAs ubicacion + "Tabla_de_Resultados_" & pMaps.Layer(CInt(TextBox1.Text)).Name & "_Buffini" & TextBox4.Text & "_r" & TextBox7.Text & "_BuffDin" & TextBox14.Text & "_Azm" & TextBox17.Text & "_Pts" & TextBox16.Text & ".xls"
    Set xLibro = objExcel.Workbooks.Open(ubicacion + "Tabla_de_Resultados_" & pMaps.Layer(CInt(TextBox1.Text)).Name & "_Buffini" & TextBox4.Text & "_r" & TextBox7.Text & "_BuffDin" & TextBox14.Text & "_Azm" & TextBox17.Text & "_Pts" & TextBox16.Text & ".xls")
    
    'Hacemos el Excel Visible
    objExcel.Visible = True
    
         With xLibro
    
             ' Hacemos referencia a la Hoja
             With .Sheets(1)
    
                 'Recorremos las filas
                 .cells(1, 1) = "NombreCalles"
                 .cells(1, 2) = "Pto_K"
                 .cells(1, 3) = "Velocidad_V"
                 .cells(1, 4) = "Velocidad_S"
                 .cells(1, 5) = "Distancia_ConAnterior"
                 .cells(1, 6) = "Distancia_Acumulada"
                 .cells(1, 7) = "Delta_T"
                 '.cells(1, 8) = "ArcoAsociado"
                 .cells(1, 8) = "Angulo_Azimut_Pto"
                 .cells(1, 9) = "Angulo_Azimut_Calle"
                 
                 
                 For Fila = 2 To NK + 1
                    .cells(Fila, 1) = Calles(Fila - 2)
                    .cells(Fila, 2) = pLayer2.FeatureClass.GetFeature(Fila - 2).Value(0)
                    .cells(Fila, 3) = VelocidadesV(Fila - 2)
                    .cells(Fila, 4) = VelocidadesS(Fila - 2)
                    .cells(Fila, 5) = DistanciaD(Fila - 2)
                    .cells(Fila, 6) = DistanciaDA(Fila - 2)
                    .cells(Fila, 7) = DeltaT(Fila - 2)
                    '.cells(Fila, 8) = IDArcoAsociadoi(Fila - 2)
                    .cells(Fila, 8) = ArrAzimutPto(Fila - 2)
                    .cells(Fila, 9) = ArrAzimutCalle(Fila - 2)
                 Next
    
             End With
         End With
    
         'Eliminamos los objetos si ya no los usamos
         Set objExcel = Nothing
         Set xLibro = Nothing
         
         MsgBox "Se ha generado el archivo Tabla_de_Resultados_" & pMaps.Layer(CInt(TextBox1.Text)).Name & "_Buffini" & TextBox4.Text & "_r" & TextBox7.Text & "_BuffDin" & TextBox14.Text & "_Azm" & TextBox17.Text & "_Pts" & TextBox16.Text & ".xls"
End Sub


Public Function ExisteAlt(ByVal idpunto As Long) As Boolean
    'Revisa si existe un arco candidato alternativo y de paso descarta el antiguo
    MinAlt = True
    If MinDistID(idpunto) <> -1 Then
        ExisteAlt = True
    Else
        ExisteAlt = False
    End If
End Function

Public Sub CaminoAlternativo(ByVal ptoK As Long)
    'En caso de Existir camino alternativo, esta funcion se utiliza para ver si es factible
    Dim usoBufdin As Boolean
    ContinuarBusqueda = True
    usobuffdin = False
        If ptoK = 0 Then
            While Not ExisteCamino(ptoK, ptoK + 1) And ContinuarBusqueda
                ContinuarBusqueda = ExisteAlt(ptoK)
                If CheckBox2.Value And Not ContinuarBusqueda And Not usoBufdin Then
                    Call SelectionBufferDinamicoEnMarch(pLayer2.FeatureClass.GetFeature(ptoK), rBuffer, pLayer)
                    ContinuarBusqueda = ExisteAlt(ptoK)
                    usoBufdin = True
                End If
            Wend
        End If
        
        If ptoK = NK - 1 Then
            While Not ExisteCamino(ptoK - 1, ptoK) And ContinuarBusqueda
                ContinuarBusqueda = ExisteAlt(ptoK)
                If CheckBox2.Value And Not ContinuarBusqueda And Not usoBufdin Then
                    Call SelectionBufferDinamicoEnMarch(pLayer2.FeatureClass.GetFeature(ptoK), rBuffer, pLayer)
                    ContinuarBusqueda = ExisteAlt(ptoK)
                    usoBufdin = True
                End If
            Wend
        End If
        
        If ptoK > 0 And ptoK < NK - 1 Then
            While (Not ExisteCamino(ptoK, ptoK + 1) Or Not ExisteCamino(ptoK - 1, ptoK)) And ContinuarBusqueda
                ContinuarBusqueda = ExisteAlt(ptoK)
                If CheckBox2.Value And Not ContinuarBusqueda And Not usoBufdin Then
                    Call SelectionBufferDinamicoEnMarch(pLayer2.FeatureClass.GetFeature(ptoK), rBuffer, pLayer)
                    ContinuarBusqueda = ExisteAlt(ptoK)
                    usoBufdin = True
                End If
            Wend
        End If
        
        If ContinuarBusqueda Then
            If ptoK > 0 Then
                If ExisteCamino(ptoK - 1, ptoK) Then
                    AceptarPuntos ptoK - 1, ptoK
                    difij = difij + 100
                End If
            End If
            
            If ptoK < NK - 1 Then
                If ExisteCamino(ptoK, ptoK + 1) Then
                    AceptarPuntos ptoK, ptoK + 1
                    difij = difij + 100
                End If
            End If
        End If
    
End Sub

Public Function CompAzimut(ByVal numPto As Integer, ByVal posMtrx As Double) As Boolean
    'Funcion para comparar el Azimuth del vehiculo y el camino
    Dim AngAzim As Double
    ArrAzimutPto(numPto) = CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading))
    numCellHeading = pLayer2.FeatureClass.FindField(TextBox15.Text)
    AngAzim = GotAzimut(numPto, posMtrx)
    
    'Sentido Diagonal derecha-abajo
    If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 90 And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) < 180 Then
        If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 180 - AngAzim - CDbl(TextBox17.Text) And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) <= 180 - AngAzim + CDbl(TextBox17.Text) Then
            ArrAzimutCalle(i) = 180 - AngAzim
            'ArrAzimutPto(i) = CDbl(pLayer2.FeatureClass.GetFeature(i).Value(numCellHeading))
            CompAzimut = True
        Else
            ArrAzimutCalle(i) = 180 - AngAzim
            CompAzimut = False
        End If
    End If
        
    'Sentido Diagonal izquierda-arriba
    If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 270 And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) < 360 Then
        If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 360 - AngAzim - CDbl(TextBox17.Text) And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) <= 360 - AngAzim + CDbl(TextBox17.Text) Then
            ArrAzimutCalle(i) = 360 - AngAzim
            'ArrAzimutPto(i) = CDbl(pLayer2.FeatureClass.GetFeature(i).Value(numCellHeading))
            CompAzimut = True
        Else
            ArrAzimutCalle(i) = 360 - AngAzim
            CompAzimut = False
        End If
    End If

    'Sentido Diagonal derecha-arriba
    If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 0 And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) < 90 Then
        If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= AngAzim - CDbl(TextBox17.Text) And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) <= AngAzim + CDbl(TextBox17.Text) Then
            ArrAzimutCalle(i) = AngAzim
            'ArrAzimutPto(i) = CDbl(pLayer2.FeatureClass.GetFeature(i).Value(numCellHeading))
            CompAzimut = True
        Else
            ArrAzimutCalle(i) = AngAzim
            CompAzimut = False
        End If
    End If
    
    'Sentido Diagonal izquierda-abajo
    If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 180 And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) < 270 Then
        If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 180 + AngAzim - CDbl(TextBox17.Text) And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) <= 180 + AngAzim + CDbl(TextBox17.Text) Then
            ArrAzimutCalle(i) = 180 + AngAzim
            'ArrAzimutPto(i) = CDbl(pLayer2.FeatureClass.GetFeature(i).Value(numCellHeading))
            CompAzimut = True
        Else
            ArrAzimutCalle(i) = 180 + AngAzim
            CompAzimut = False
        End If
    End If

End Function

Public Function CompAzimut2(ByVal numPto As Integer, ByVal posArc As Double) As Boolean
    'Otra funcion para comparar azumuth, (descontinuada)
    Dim AngAzim As Double
    numCellHeading = pLayer2.FeatureClass.FindField(TextBox15.Text)
    AngAzim = GotAzimut2(numPto, posArc)
    
    'Sentido Diagonal derecha-abajo
    If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 90 And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) < 180 Then
        If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 180 - AngAzim - CDbl(TextBox17.Text) And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) <= 180 - AngAzim + CDbl(TextBox17.Text) Then
            ArrAzimutCalle(i) = 180 - AngAzim
            'ArrAzimutPto(i) = CDbl(pLayer2.FeatureClass.GetFeature(i).Value(numCellHeading))
            CompAzimut2 = True
        Else
            ArrAzimutCalle(i) = 180 - AngAzim
            CompAzimut2 = False
        End If
    End If
        
    'Sentido Diagonal izquierda-arriba
    If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 270 And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) < 360 Then
        If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 360 - AngAzim - CDbl(TextBox17.Text) And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) <= 360 - AngAzim + CDbl(TextBox17.Text) Then
            ArrAzimutCalle(i) = 360 - AngAzim
            'ArrAzimutPto(i) = CDbl(pLayer2.FeatureClass.GetFeature(i).Value(numCellHeading))
            CompAzimut2 = True
        Else
            ArrAzimutCalle(i) = 360 - AngAzim
            CompAzimut2 = False
        End If
    End If

    'Sentido Diagonal derecha-arriba
    If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 0 And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) < 90 Then
        If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= AngAzim - CDbl(TextBox17.Text) And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) <= AngAzim + CDbl(TextBox17.Text) Then
            ArrAzimutCalle(i) = AngAzim
            'ArrAzimutPto(i) = CDbl(pLayer2.FeatureClass.GetFeature(i).Value(numCellHeading))
            CompAzimut2 = True
        Else
            ArrAzimutCalle(i) = AngAzim
            CompAzimut2 = False
        End If
    End If
    
    'Sentido Diagonal izquierda-abajo
    If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 180 And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) < 270 Then
        If CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) >= 180 + AngAzim - CDbl(TextBox17.Text) And CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading)) <= 180 + AngAzim + CDbl(TextBox17.Text) Then
            ArrAzimutCalle(i) = 180 + AngAzim
            'ArrAzimutPto(i) = CDbl(pLayer2.FeatureClass.GetFeature(i).Value(numCellHeading))
            CompAzimut2 = True
        Else
            ArrAzimutCalle(i) = 180 + AngAzim
            CompAzimut2 = False
        End If
    End If

End Function

Public Function GotAzimut(ByVal numPto As Integer, ByVal posMtrx As Double) As Double
    'Se encarga de calcular el angulo azimuth de un arco en un punto especifico
    Set SnapPtoAct = MtrxPto(numPto, posMtrx)
    Set pTopoOp = SnapPtoAct
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.GeometryField = "SHAPE"
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Set pPolygon = pTopoOp.buffer(1)
    Set pSpatialFilter.Geometry = pPolygon
    
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim pProxOp As IProximityOperator
    Dim pProxOp2 As IProximityOperator
    Dim pProxOp3 As IProximityOperator
    Dim pGeom As IGeometry
    Dim p2Geom As IGeometry
    Dim p3Geom As IGeometry
    Dim pGeom1 As IGeometry
    Dim p2Geom1 As IGeometry
    Dim p3Geom1 As IGeometry
    Dim punto As IGeometry
    Dim punto2 As IGeometry
    Dim punto3 As IGeometry
    Dim pto_1 As IPoint
    Dim pto_2 As IPoint
    Dim pto_3 As IPoint
    Dim angulo As Double
    Dim angPto As Double
    angPto = CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading))
    
    Set pto_1 = New Point
    pto_1.PutCoords SnapPtoAct.x, SnapPtoAct.Y + 1
    c = 1
    'MsgBox pto1.x & "   " & pto1.Y
    
    Set pGeom1 = pto_1
    Set pGeom = pLayer.FeatureClass.GetFeature(MtrxArc(numPto, posMtrx)).Shape
    Set pProxOp = pGeom
    Set punto = pProxOp.ReturnNearestPoint(pGeom1, 0)
    Set pProxOp = punto
    
    Set pto_2 = punto
    a = pProxOp.ReturnDistance(pGeom1)
    'MsgBox pto2.x & "   " & pto2.Y
    
    Set p2Geom1 = pto_2
    Set p2Geom = SnapPtoAct
    Set pProxOp2 = p2Geom
    Set punto2 = pProxOp2.ReturnNearestPoint(p2Geom1, 0)
    Set pProxOp2 = punto2
    
    Set pto_3 = punto2
    b = pProxOp2.ReturnDistance(p2Geom1)
    'MsgBox pto3.x & "   " & pto3.Y
    
    'CasoBorde: c=a=1 y b=0
    If b = 0 And a = c And angPto > 0 And angPto < 180 Then
        GotAzimut = 90
        Exit Function
    End If
    
    If b = 0 And a = c And angPto > 180 And angPto < 360 Then
        GotAzimut = 270
        Exit Function
    End If
    
    
    If b <> 0 And a <> 0 Then
        angulo = (180 / 3.14159265358979) * ArcoCoseno(((b * b) + (c * c) - (a * a)) / (2 * b * c))
        GotAzimut = angulo
    End If
End Function

Public Function GotAzimut2(ByVal numPto As Integer, ByVal posArc As Double) As Double
    ' idem (descontinuada)
    Set pTopoOp = SnapPtoAct
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.GeometryField = "SHAPE"
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Set pPolygon = pTopoOp.buffer(1)
    Set pSpatialFilter.Geometry = pPolygon
    
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim pProxOp As IProximityOperator
    Dim pProxOp2 As IProximityOperator
    Dim pProxOp3 As IProximityOperator
    Dim pGeom As IGeometry
    Dim p2Geom As IGeometry
    Dim p3Geom As IGeometry
    Dim pGeom1 As IGeometry
    Dim p2Geom1 As IGeometry
    Dim p3Geom1 As IGeometry
    Dim punto As IGeometry
    Dim punto2 As IGeometry
    Dim punto3 As IGeometry
    Dim pto_1 As IPoint
    Dim pto_2 As IPoint
    Dim pto_3 As IPoint
    Dim angulo As Double
    Dim angPto As Double
    angPto = CDbl(pLayer2.FeatureClass.GetFeature(numPto).Value(numCellHeading))
    
    Set pto_1 = New Point
    pto_1.PutCoords SnapPtoAct.x, SnapPtoAct.Y + 1
    c = 1
    'MsgBox pto1.x & "   " & pto1.Y
    
    Set pGeom1 = pto_1
    Set pGeom = pLayer.FeatureClass.GetFeature(posArc).Shape
    Set pProxOp = pGeom
    Set punto = pProxOp.ReturnNearestPoint(pGeom1, 0)
    Set pProxOp = punto
    
    Set pto_2 = punto
    a = pProxOp.ReturnDistance(pGeom1)
    'MsgBox pto2.x & "   " & pto2.Y
    
    Set p2Geom1 = pto_2
    Set p2Geom = SnapPtoAct
    Set pProxOp2 = p2Geom
    Set punto2 = pProxOp2.ReturnNearestPoint(p2Geom1, 0)
    Set pProxOp2 = punto2
    
    Set pto_3 = punto2
    b = pProxOp2.ReturnDistance(p2Geom1)
    'MsgBox pto3.x & "   " & pto3.Y
    
    'CasoBorde: c=a=1 y b=0
    If b = 0 And a = c And angPto > 0 And angPto < 180 Then
        GotAzimut2 = 90
        Exit Function
    End If
    
    If b = 0 And a = c And angPto > 180 And angPto < 360 Then
        GotAzimut2 = 270
        Exit Function
    End If
    
    
    If b <> 0 And a <> 0 Then
        angulo = (180 / 3.14159265358979) * ArcoCoseno(((b * b) + (c * c) - (a * a)) / (2 * b * c))
        GotAzimut2 = angulo
    End If
End Function


Public Function ExisteCamino(ByVal idk1 As Long, ByVal idk2 As Long) As Boolean
        'Funcion que es el corazon del algoritmo
        'Utiliza la comparacion de velocidades para verificar si un segmento de ruta es viable
        'Ademas se implementa la comparacion de los Azimuth
        Dim existRoute As Boolean
        Set K1 = pLayer2.FeatureClass.GetFeature(idk1)
        Set K2 = pLayer2.FeatureClass.GetFeature(idk2)
        
        D = -9999
        S = -9999
        V = -9999
        Call PromVelKs(K1, K2)
        ''''''''
        If V = 0 Then
                If CDbl(K1.Value(posSpeed)) = 0 Then
                    VelUmbMinima K1
                    If Calles(K1.Value(0) - 1) <> "" Then
                        Calles(K1.Value(0)) = Calles(K1.Value(0) - 1)
                    Else
                        Calles(K1.Value(0)) = Calles(K1.Value(0) - 2)
                    End If
                End If
                
                If CDbl(K2.Value(posSpeed)) = 0 Then
                    VelUmbMinima K2
                    Calles(K2.Value(0)) = Calles(K2.Value(0) - 1)
                End If
            difij = 999
            D = 0
            ExisteCamino = False
        End If
        ''''''''
        Call EvaluarDistPtos(K1, K2)
        If D > 0 And CheckBox3.Value = False Then
            Call CalcularS(K1, K2)
            
            ExisteCamino = CompararVyS()
'            If ExisteCamino Then
'                ContinuarBusqueda = False
'            End If
        End If
        
        If D > 0 And CheckBox3.Value = True Then
            Call CalcularS(K1, K2)
            If CompararVyS() Then
                If CompAzimut(K1.Value(0), spID1) And CompAzimut(K2.Value(0), spID2) Then
'                    ContinuarBusqueda = False
                    ExisteCamino = True
                End If
            Else
                ExisteCamino = False
            End If
        End If
        


End Function


Public Function ObtenerVelocidad(ByVal K As IFeature, ByVal posSpd As Double) As Double
        'el nombre lo dice todo
        ObtenerVelocidad = CDbl(K.Value(posSpd))

End Function

Public Sub VelUmbMinima(ByVal Kx As IFeature)
    'Funcion para la velocidad umbral minima
    For j = 0 To NA - 1
        MtrxDescarte(Kx.Value(0), j) = False
    Next
    If K2.Value(0) < NK - 2 Then
        'Call ExisteCamino(K1.Value(0), K2.Value(0) + 1)
        D = 0
    End If
End Sub

Public Sub PromVelKs(ByVal pto1 As IFeature, ByVal pto2 As IFeature)
            'Calcula velocidad promedio entre puntos
            If CheckBox1.Value = True And K1.Value(0) <> 0 And K2.Value(0) <> NK - 1 And (CDbl(K1.Value(posSpeed)) < velUmb Or CDbl(K2.Value(posSpeed)) < velUmb) Then
                If CDbl(K1.Value(posSpeed)) < velUmb Then
                    VelUmbMinima K1
                    Calles(K1.Value(0)) = Calles(K1.Value(0) - 1)
                End If
                
                If CDbl(K2.Value(posSpeed)) < velUmb Then
                    VelUmbMinima K2
                    Calles(K2.Value(0)) = Calles(K2.Value(0) - 1)
                End If
'
'                Callesi(K1.Value(0)) = Callesj(K1.Value(0) - 1)
'                Callesj(K2.Value(0)) = Callesj(K2.Value(0) - 1)
                difij = 999
                D = 0
'                DistanciaDA(i) = DistanciaDA(i - 1)
            Else
                V = (ObtenerVelocidad(K1, posSpeed) + ObtenerVelocidad(K2, posSpeed)) / 2
            End If
End Sub

Public Sub EvaluarDistPtos(ByVal pto1 As IFeature, ByVal pto2 As IFeature)
    'Evalua la distancia de los puntos en observacion y sus arcos candidatos
    'Para cada punto se busca el arco mas cercano NO descartado
    Dim PtoSnapID1 As Long
    Dim PtoSnapID2 As Long
    
    PtoSnapID1 = MinDistID(pto1.Value(0))
    PtoSnapID2 = MinDistID(pto2.Value(0))
    auxID1 = PtoSnapID1
    auxID2 = PtoSnapID2
    If auxID1 <> -1 And auxID2 <> -1 Then
        Call CreatePointFeature(MtrxPto(pto1.Value(0), PtoSnapID1), "NO SNAP")
        Call CreatePointFeature(MtrxPto(pto2.Value(0), PtoSnapID2), "NO SNAP")
        Call TerminaEdicion
        Call EfectuarRuta(pfeaturelayer, 0, 0, 1)
        If Not rut_inviable Then
            Call LimpiarResultados
            Call ComienzaEdicion
            D = CalculaDistRuta()
        Else
            Call LimpiarResultados
            Call ComienzaEdicion
        End If
    End If

End Sub

Public Function MinDistID(ByVal ptoid As Long) As Long
    'Busca el arco mas cercano No descartado
    'En caso de buscar un arco alternativo, descarta el antiguo (ver funcion ExisteAlt)
    Dim MinID As Long
    Min = 999999999999#
    MinID = -1
    j = 0
        While MtrxDist(ptoid, j) <> -1 And j < NA 'j debe ser menor al numero de arcos al interior del buffer
            If Min > MtrxDist(ptoid, j) And MtrxDist(ptoid, j) >= ArrPrevMin(ptoid) And MtrxDescarte(ptoid, j) Then
                Min = MtrxDist(ptoid, j)
                MinID = j
            End If
            j = j + 1
        Wend
        If MinID = -1 Then
            MinAlt = False
        End If
        If MinAlt And Min <> 999999999999# Then
            ArrPrevMin(ptoid) = Min
            MtrxDescarte(ptoid, MinID) = False
            MinAlt = False
        End If
        
    MinDistID = MinID
End Function

Public Sub CalcularS(ByVal pto1 As IFeature, ByVal pto2 As IFeature)
    'Calcula la velocidad promedio S (unidades metricas)
    If OptionButton1.Value = True Then
        'Si se escoge obtenerlo de una columna (formato hhmmss **h=hora, m=minuto, s=segundo**)
        t1 = FormatearHora(pto1.Value(posTime))
        t2 = FormatearHora(pto2.Value(posTime))
    
        S = D * 3.6 / DateDiff("s", t1, t2)
        DeltaT(i) = DateDiff("s", t1, t2)
    End If
    
    If OptionButton2.Value = True Then
        'Si se escoge ingresar la frecuencia de forma estatica
        S = D * 3.6 / (CDbl(TextBox11.Text) * (CLng(pto2.Value(0)) - CLng(pto1.Value(0))))
        DeltaT(i) = (CDbl(TextBox11.Text) * (CLng(pto2.Value(0)) - CLng(pto1.Value(0))))
    End If
    
    DistanciaD(i + 1) = D
    
    Dim h As Integer
    h = 0
    While DistanciaDA(i - h) = 0 And i - h > 0
        h = h + 1
    Wend
    DistanciaDA(i + 1) = DistanciaD(i + 1) + DistanciaDA(i - h)
End Sub


Public Function CompararVyS() As Boolean
    'Se comparan las velocidades promedios
    VelocidadesS(i) = S
    VelocidadesV(i) = V
    If V >= S - CDbl(TextBox7.Text) And V <= S + CDbl(TextBox7.Text) Then
        spID1 = auxID1
        spID2 = auxID2
        ArID(K1.Value(0)) = auxID1
        ArID(K2.Value(0)) = auxID2

        CompararVyS = True
    Else
        CompararVyS = False
    End If
End Function

Public Sub AceptarPuntos(ByVal ptoid1 As Long, ByVal ptoid2 As Long)
        'Luego de evaluar los pares de puntos y ser aceptada la ruta se guarda la informacion
        Dim contIDPtos As Long
        Dim contIDArcs As Long
        Dim NumCampoCalle As Integer
        Dim dist As Double
        Dim minimo As Double
        Dim idk As Long
        Dim idArc As Long
        Dim auxAzm As Boolean
        contIDArcs = 0
        
        Set Aceptados(ptoid1) = New Point
        Set Aceptados(ptoid2) = New Point
        NumCampoCalle = pLayer.FeatureClass.FindField(TextBox12.Text)
        Aceptados(ptoid1).PutCoords MtrxPto(ptoid1, ArID(ptoid1)).x, MtrxPto(ptoid1, ArID(ptoid1)).Y
        Calles(ptoid1) = pLayer.FeatureClass.GetFeature(MtrxArc(ptoid1, ArID(ptoid1))).Value(NumCampoCalle)
        Aceptados(ptoid2).PutCoords MtrxPto(ptoid2, ArID(ptoid2)).x, MtrxPto(ptoid2, ArID(ptoid2)).Y
        Calles(ptoid2) = pLayer.FeatureClass.GetFeature(MtrxArc(ptoid2, ArID(ptoid2))).Value(NumCampoCalle)
        
        IDArcoAsociadoi(ptoid1) = ArID(ptoid1) 'MtrxArc(ptoid1, spID1)
        IDArcoAsociadoi(ptoid2) = ArID(ptoid2) 'MtrxArc(ptoid2, spID2)
        ptoKi(i) = ptoid1
        ptoKj(i) = ptoid2

        
        If ptoid1 <> ptoid2 - 1 Then
            Call ArcosEntrePuntos
            For contIDPtos = ptoid1 + 1 To ptoid2 - 1
              If pLayer2.FeatureClass.GetFeature(contIDPtos).Value(posSpeed) > velUmb Then
                idk = -1
                minimo = 999999999999#
                    For contIDArcs = 0 To UBound(ArrArcRuta) - 1
                        dist = CalcDistPtoArcAcept(pLayer2.FeatureClass.GetFeature(contIDPtos), pLayer.FeatureClass.GetFeature(ArrArcRuta(contIDArcs)))
'                        If CheckBox3.Value = True Then
'                            If dist < minimo And CompAzimut2(contIDPtos, ArrArcRuta(contIDArcs)) Then
'                                minimo = dist
'                                idk = contIDPtos
'                                idArc = ArrArcRuta(contIDArcs)
'                            End If
'                        Else
                            If dist < minimo Then
                                minimo = dist
                                idk = contIDPtos
                                idArc = ArrArcRuta(contIDArcs)
                            End If
'                        End If
                    Next
                If idk <> -1 Then
                    Set Aceptados(idk) = New Point
                    Call GuardarPtoAcept(pLayer2.FeatureClass.GetFeature(idk), pLayer.FeatureClass.GetFeature(idArc), idk)
                    Calles(idk) = pLayer.FeatureClass.GetFeature(idArc).Value(NumCampoCalle)
                    'Callesj(idk) = pLayer.FeatureClass.GetFeature(idArc).Value(NumCampoCalle)
                    'IDArcoAsociado(idk) = idarc
                End If
              End If
            Next
        End If
End Sub


Public Sub EfectuarRuta(ByVal LayerPuntos As IFeatureLayer, ByVal buffer As Double, ByVal puntok1 As Long, ByVal puntok2 As Long)
    'Efectua rutas como su nombre lo menciona
    
    Set pGPMessages = New GPMessages
    Set pNAExtension = Application.FindExtensionByName("Network Analyst")
    Set pNAWindow = pNAExtension.NAWindow
    Set pNALayer = pNAWindow.ActiveAnalysis
    Set pNAContext = pNALayer.Context
    Set pNAWindowCategory = pNAWindow.ActiveCategory
    
    'Make sure we have a valid category selected in the NAWindow
    If pNAWindowCategory Is Nothing Then
        MsgBox "You must have an analysis layer in your map"
        Exit Sub
    End If
 
    
    LoadNANetworkLocations pNAContext, "Stops", LayerPuntos.FeatureClass, buffer, puntok1, puntok2

    On Error GoTo handler
        If pNAContext.Solver.Solve(pNAContext, pGPMessages, Nothing) Then
            'MsgBox ("Ruta Generada")
        End If
    Exit Sub
handler:
    D = -1
    difij = 100
    rutag = False
    rut_inviable = True
    Resume Next
    'pMxDoc.ActiveView.Refresh
    
End Sub



Public Sub LoadNANetworkLocations(ByRef pContext As INAContext, _
                                    ByVal strNAClassName As String, _
                                    ByVal pInputFC As IFeatureClass, _
                                    ByVal SnapTolerance As Double, _
                                    ByVal puntok1 As Long, _
                                    ByVal puntok2 As Long)
                                   
    

    ' Efectua rutas en pares de puntos (si rutafull es false) o de la layer completa (si rutafull es true)
    Set pClasses = pContext.NAClasses
    Set pNAClass = pClasses.ItemByName(strNAClassName)
    Set pQFilt = New QueryFilter
    
    strquery = """" & pInputFC.Fields.Field(0).Name & """ = " & puntok1 & " OR """ & pInputFC.Fields.Field(0).Name & """ =  " & puntok2
    
    pQFilt.WhereClause = strquery
   
    
    'Para cada ruta a calcular se debe borrar los stops previamente utilizados
    pNAClass.DeleteAllRows
       
    
    ' Create a NAClassLoader and set the snap tolerance (meters unit)
    Dim pLoader As INAClassLoader
    Set pLoader = New NAClassLoader
    Set pLoader.Locator = pContext.Locator
    If SnapTolerance > 0 Then pLoader.Locator.SnapTolerance = SnapTolerance
    Set pLoader.NAClass = pNAClass
    
    'Create field map to automatically map fields from input class to naclass
    Dim pFieldMap As INAClassFieldMap
    Set pFieldMap = New NAClassFieldMap
    
    ' Si se efectua la ruta completa, se debe utilizar todos los puntos, sino se utilizaran 2 puntos para calcular la dist entre ellos en la red.
    If RutaFull Then
        pFieldMap.CreateMapping pNAClass.ClassDefinition, pInputFC.Fields
    Else
        pFieldMap.CreateMapping pNAClass.ClassDefinition, pInputFC.Search(pQFilt, False).Fields
    End If
    Set pLoader.FieldMap = pFieldMap
    
    'Load Network Locations
    Dim rowsIn As Long
    Dim rowsLocated As Long
    
    If RutaFull Then
        pLoader.Load pInputFC.Search(Nothing, True), Nothing, rowsIn, rowsLocated
    Else
        pLoader.Load pInputFC.Search(pQFilt, True), Nothing, rowsIn, rowsLocated
    End If

End Sub

Public Sub SelectionBuffer(ByVal pFeature As IFeature, ByVal buffer As Double, ByVal pfeaturelayer As IFeatureLayer)

    'Utilizando el buffer para cada punto
    'Utilizada para guardar la informacion en las matrices
    Set pTopoOp = pFeature.Shape
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.GeometryField = "SHAPE"
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    
    
    Set pPolygon = pTopoOp.buffer(buffer)
    Set pSpatialFilter.Geometry = pPolygon
    Set pFeatureSelection = pfeaturelayer
    

    pFeatureSelection.SelectFeatures pSpatialFilter, esriSelectionResultNew, False 'esriSpatialRelIntersects 'esriSelectionResultNew
    
    'MsgBox pFeatureSelection.SelectionSet.Count
    If pFeatureSelection.SelectionSet.Count > NA Then
        NA = pFeatureSelection.SelectionSet.Count
        ReDim Preserve MtrxArc(NK, NA) As Long
        ReDim Preserve MtrxDist(NK, NA) As Double
        ReDim Preserve MtrxPto(NK, NA) As IPoint
        ReDim Preserve MtrxDescarte(NK, NA) As Boolean
    End If
    
    Set pEnumIDs = pFeatureSelection.SelectionSet.IDs
    Dim idauxiliar As Integer
    'Loop que busca los arcos seleccionados
    id = pEnumIDs.Next
    j = 0
    idauxiliar = 0
    
    While idauxiliar <> -1
        If id = -1 Then
            idauxiliar = -1
        End If
       'GUARDAR LOS IDS EN UNA MATRIZ
        MtrxArc(pFeature.Value(0), j) = id
        If id <> -1 Then
            MtrxDist(pFeature.Value(0), j) = CalcDistPtoArc(pFeature, pfeaturelayer.FeatureClass.GetFeature(id), pFeature.Value(0), j)
            MtrxDescarte(pFeature.Value(0), j) = True
        Else
            MtrxDist(pFeature.Value(0), j) = -1
            MtrxDescarte(pFeature.Value(0), j) = False
        End If
        id = pEnumIDs.Next
        j = j + 1
    Wend
    
    'pMxDoc.ActivatedView.Refresh
    pFeatureSelection.Clear
End Sub

'Public Sub SelectionBufferDinamico(ByVal pFeature As IFeature, ByVal buffer As Double, ByVal pfeaturelayer As IFeatureLayer)
'
'
'
'    Set pTopoOp = pFeature.Shape
'    Set pSpatialFilter = New SpatialFilter
'    pSpatialFilter.GeometryField = "SHAPE"
'    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
'
'
'    Set pPolygon = pTopoOp.buffer(buffer)
'    Set pSpatialFilter.Geometry = pPolygon
'    Set pFeatureSelection = pfeaturelayer
'
'
'    pFeatureSelection.SelectFeatures pSpatialFilter, esriSelectionResultNew, False 'esriSpatialRelIntersects 'esriSelectionResultNew
'
'    'MsgBox pFeatureSelection.SelectionSet.Count
'    If pFeatureSelection.SelectionSet.Count > NA Then
'        NA = pFeatureSelection.SelectionSet.Count
'        ReDim Preserve MtrxArc(NK, NA) As Long
'        ReDim Preserve MtrxDist(NK, NA) As Double
'        ReDim Preserve MtrxPto(NK, NA) As IPoint
'        ReDim Preserve MtrxDescarte(NK, NA) As Boolean
'    End If
'
'    If pFeatureSelection.SelectionSet.Count = 0 Then
'        multiplicador = 1
'        DeltaBuffer = BufferDinamico(pFeature)
'        While distMay > multiplicador * DeltaBuffer + buffer And pFeatureSelection.SelectionSet.Count = 0
'            Set pPolygon = pTopoOp.buffer(buffer + DeltaBuffer)
'            Set pSpatialFilter.Geometry = pPolygon
'            pFeatureSelection.SelectFeatures pSpatialFilter, esriSelectionResultNew, False
'            multiplicador = multiplicador + 1
'            If pFeatureSelection.SelectionSet.Count > NA Then
'                NA = pFeatureSelection.SelectionSet.Count
'                ReDim Preserve MtrxArc(NK, NA) As Long
'                ReDim Preserve MtrxDist(NK, NA) As Double
'                ReDim Preserve MtrxPto(NK, NA) As IPoint
'                ReDim Preserve MtrxDescarte(NK, NA) As Boolean
'            End If
'        Wend
'    End If
'
'
'    Set pEnumIDs = pFeatureSelection.SelectionSet.IDs
'    Dim idauxiliar As Integer
'    'Loop que busca los arcos seleccionados
'    id = pEnumIDs.Next
'    j = 0
'    idauxiliar = 0
'
'    While idauxiliar <> -1
'        If id = -1 Then
'            idauxiliar = -1
'        End If
'       'GUARDAR LOS IDS EN UNA MATRIZ
'        MtrxArc(pFeature.Value(0), j) = id
'        If id <> -1 Then
'            MtrxDist(pFeature.Value(0), j) = CalcDistPtoArc(pFeature, pfeaturelayer.FeatureClass.GetFeature(id), pFeature.Value(0), j)
'            MtrxDescarte(pFeature.Value(0), j) = True
'        Else
'            MtrxDist(pFeature.Value(0), j) = -1
'            MtrxDescarte(pFeature.Value(0), j) = False
'        End If
'        id = pEnumIDs.Next
'        j = j + 1
'    Wend
'
'    'pMxDoc.ActivatedView.Refresh
'    pFeatureSelection.Clear
'End Sub
'
'Public Function BufferDinamico(ByVal PuntoK As IFeature) As Double
'
'
'    If i > 0 And i < NK - 1 Then
'        If CalcDistPtoArcAcept(PuntoK, pLayer2.FeatureClass.GetFeature(PuntoK.Value(0) + 1)) > CalcDistPtoArcAcept(PuntoK, pLayer2.FeatureClass.GetFeature(PuntoK.Value(0) - 1)) Then
'            distMay = 2 * CalcDistPtoArcAcept(PuntoK, pLayer2.FeatureClass.GetFeature(PuntoK.Value(0) + 1))
'        Else
'            distMay = 2 * CalcDistPtoArcAcept(PuntoK, pLayer2.FeatureClass.GetFeature(PuntoK.Value(0) - 1))
'        End If
'    End If
'
'    If i = 0 Then
'        distMay = 2 * CalcDistPtoArcAcept(PuntoK, pLayer2.FeatureClass.GetFeature(PuntoK.Value(0) + 1))
'    End If
'
'    If i = NK - 1 Then
'        distMay = 2 * CalcDistPtoArcAcept(PuntoK, pLayer2.FeatureClass.GetFeature(PuntoK.Value(0) - 1))
'    End If
'    BufferDinamico = (distMay - rBuffer) * CDbl(TextBox14.Text) / 100
'
'
'End Function

Public Sub SelectionBufferDinamicoPrev(ByVal pFeature As IFeature, ByVal buffer As Double, ByVal pfeaturelayer As IFeatureLayer)
    
    'Funcion del buffer dimanico (previo al algoritmo)
    
    Set pTopoOp = pFeature.Shape
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.GeometryField = "SHAPE"
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    
    
    Set pPolygon = pTopoOp.buffer(buffer)
    Set pSpatialFilter.Geometry = pPolygon
    Set pFeatureSelection = pfeaturelayer
    

    pFeatureSelection.SelectFeatures pSpatialFilter, esriSelectionResultNew, False 'esriSpatialRelIntersects 'esriSelectionResultNew
    
    'MsgBox pFeatureSelection.SelectionSet.Count
    If pFeatureSelection.SelectionSet.Count > NA Then
        NA = pFeatureSelection.SelectionSet.Count
        ReDim Preserve MtrxArc(NK, NA) As Long
        ReDim Preserve MtrxDist(NK, NA) As Double
        ReDim Preserve MtrxPto(NK, NA) As IPoint
        ReDim Preserve MtrxDescarte(NK, NA) As Boolean
    End If
    
    If pFeatureSelection.SelectionSet.Count = 0 Then
        multiplicador = 1
        DeltaBuffer = CDbl(TextBox14.Text)
        While buffer * DeltaBuffer >= multiplicador * buffer And pFeatureSelection.SelectionSet.Count = 0
            Set pPolygon = pTopoOp.buffer(buffer * multiplicador)
            Set pSpatialFilter.Geometry = pPolygon
            pFeatureSelection.SelectFeatures pSpatialFilter, esriSelectionResultNew, False
            multiplicador = multiplicador + 1
            If pFeatureSelection.SelectionSet.Count > NA Then
                NA = pFeatureSelection.SelectionSet.Count
                ReDim Preserve MtrxArc(NK, NA) As Long
                ReDim Preserve MtrxDist(NK, NA) As Double
                ReDim Preserve MtrxPto(NK, NA) As IPoint
                ReDim Preserve MtrxDescarte(NK, NA) As Boolean
            End If
        Wend
    End If
    
    
    Set pEnumIDs = pFeatureSelection.SelectionSet.IDs
    Dim idauxiliar As Integer
    'Loop que busca los arcos seleccionados
    id = pEnumIDs.Next
    j = 0
    idauxiliar = 0
    
    While idauxiliar <> -1
        If id = -1 Then
            idauxiliar = -1
        End If
       'GUARDAR LOS IDS EN UNA MATRIZ
        MtrxArc(pFeature.Value(0), j) = id
        If id <> -1 Then
            MtrxDist(pFeature.Value(0), j) = CalcDistPtoArc(pFeature, pfeaturelayer.FeatureClass.GetFeature(id), pFeature.Value(0), j)
            MtrxDescarte(pFeature.Value(0), j) = True
        Else
            MtrxDist(pFeature.Value(0), j) = -1
            MtrxDescarte(pFeature.Value(0), j) = False
        End If
        id = pEnumIDs.Next
        j = j + 1
    Wend
    
    'pMxDoc.ActivatedView.Refresh
    pFeatureSelection.Clear
End Sub
Public Sub SelectionBufferDinamicoEnMarch(ByVal pFeature As IFeature, ByVal buffer As Double, ByVal pfeaturelayer As IFeatureLayer)
    
    'Buffer dinamico en la marcha
    
    Set pTopoOp = pFeature.Shape
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.GeometryField = "SHAPE"
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    DeltaBuffer = CDbl(TextBox14.Text)
    
    Set pPolygon = pTopoOp.buffer(buffer * DeltaBuffer)
    Set pSpatialFilter.Geometry = pPolygon
    Set pFeatureSelection = pfeaturelayer
    

    pFeatureSelection.SelectFeatures pSpatialFilter, esriSelectionResultNew, False 'esriSpatialRelIntersects 'esriSelectionResultNew
    
    'MsgBox pFeatureSelection.SelectionSet.Count
    If pFeatureSelection.SelectionSet.Count > NA Then
        NA = pFeatureSelection.SelectionSet.Count
        ReDim Preserve MtrxArc(NK, NA) As Long
        ReDim Preserve MtrxDist(NK, NA) As Double
        ReDim Preserve MtrxPto(NK, NA) As IPoint
        ReDim Preserve MtrxDescarte(NK, NA) As Boolean
    End If
    

    Set pEnumIDs = pFeatureSelection.SelectionSet.IDs
    Dim idauxiliar As Integer
    'Loop que busca los arcos seleccionados
    id = pEnumIDs.Next
    j = 0
    idauxiliar = 0
    
    While idauxiliar <> -1
        If id = -1 Then
            idauxiliar = -1
        End If
       'GUARDAR LOS IDS EN UNA MATRIZ
        MtrxArc(pFeature.Value(0), j) = id
        If id <> -1 Then
            MtrxDist(pFeature.Value(0), j) = CalcDistPtoArc(pFeature, pfeaturelayer.FeatureClass.GetFeature(id), pFeature.Value(0), j)
        Else
            MtrxDist(pFeature.Value(0), j) = -1
            MtrxDescarte(pFeature.Value(0), j) = False
        End If
        id = pEnumIDs.Next
        j = j + 1
    Wend
    
    'pMxDoc.ActivatedView.Refresh
    pFeatureSelection.Clear
End Sub
Public Function CalculaDistRuta() As Double
    'Calcula la distancia de un tramo de ruta

    Set pMxDoc = ThisDocument
    Set pMaps = pMxDoc.FocusMap
    Set RouteLayer = pMaps.Layer(CInt(TextBox5.Text))
    
    Set pRouteLayer = RouteLayer.LayerByNAClassName("Routes")
    rID = pRouteLayer.FeatureClass.Search(Nothing, True).NextFeature.Value(0)
    
    ' El rID es el id de la ruta generada, el value(10) corresponde a la columna Total_costo
    'MsgBox "Nombre de la layer de rutas es: " & pRouteLayer.Name
    'MsgBox "Los costos de la ruta son: " & pRouteLayer.FeatureClass.GetFeature(rID).Value(10)
    CalculaDistRuta = pRouteLayer.FeatureClass.GetFeature(rID).Value(10)
End Function





Private Function CalcDistPtoArc(ByVal pPunto As IFeature, ByVal pArco As IFeature, ByVal idPto As Long, ByVal idsnap As Long) As Double
    'Calcula distancia entre un punto y un arco
    'tambien guarda el punto snap (punto mas cercano a dicho arco)
    Dim pProxOp As IProximityOperator
    Dim pGeom As IGeometry
    Dim pGeom1 As IGeometry
    Dim punto As IGeometry
    Dim pto As IPoint
    Set pGeom1 = pPunto.Shape
    Set pGeom = pArco.Shape
    Set pProxOp = pGeom
    
    Set punto = pProxOp.ReturnNearestPoint(pGeom1, esriNoExtension)
    Set pProxOp = punto
    

    Set pto = punto
     
    Set MtrxPto(idPto, idsnap) = New Point
    MtrxPto(idPto, idsnap).PutCoords pto.x, pto.Y

    
    CalcDistPtoArc = pProxOp.ReturnDistance(pGeom1)
    
    
    
    

End Function

Public Function CalcDistPtoArcAcept(ByVal pPunto As IFeature, ByVal pArco As IFeature) As Double
    'Calcula distancia entre un punto y un arco
    'se utiliza cuando existe una ruta aceptada entre un par de puntos, pero los puntos intermedios no tienen snap y se debe buscar los arcos mas cercanos en la ruta
    Dim pProxOp As IProximityOperator
    Dim pGeom As IGeometry
    Dim pGeom1 As IGeometry
    Dim punto As IGeometry
    Set pGeom1 = pPunto.Shape
    Set pGeom = pArco.Shape
    Set pProxOp = pGeom
    
    Set punto = pProxOp.ReturnNearestPoint(pGeom1, esriNoExtension)
    Set pProxOp = punto
    
    CalcDistPtoArcAcept = pProxOp.ReturnDistance(pGeom1)
End Function

Public Sub GuardarPtoAcept(ByVal pPunto As IFeature, ByVal pArco As IFeature, ByVal idPto As Long)
    'Despues de aceptar los puntos se guardan con esta funcion
    Dim pProxOp As IProximityOperator
    Dim pGeom As IGeometry
    Dim pGeom1 As IGeometry
    Dim punto As IGeometry
    Dim pto As IPoint
    Set pGeom1 = pPunto.Shape
    Set pGeom = pArco.Shape
    Set pProxOp = pGeom
    
    Set punto = pProxOp.ReturnNearestPoint(pGeom1, esriNoExtension)
    Set pProxOp = punto
    

    Set pto = punto
     
    Set Aceptados(idPto) = New Point
    Aceptados(idPto).PutCoords pto.x, pto.Y

End Sub

Private Sub CreatePointFeature(pPoint As IPoint, calle As String)
        'Crea puntos para guardarlos en la layer de resultados
      Set pNewFeature = pFeatureClass.CreateFeature
      With pNewFeature
           Set .Shape = pPoint
      End With
'      If calle = "" Then
'        pNewFeature.Value(2) = Calles(i)
'      Else
        pNewFeature.Value(2) = calle
'      End If
      pNewFeature.Value(3) = pLayer2.FeatureClass.GetFeature(i).Value(posSpeed)
      pNewFeature.Store
End Sub

Private Sub LimpiarResultados()

      ptable.DeleteSearchedRows Nothing

End Sub


Private Sub ComienzaEdicion()
      If Not TypeOf pMaps.Layer(CInt(TextBox6.Text)) Is IFeatureLayer Then
            MsgBox "First layer must be a Featurelayer"
            Exit Sub
      End If
      Set pfeaturelayer = pMaps.Layer(CInt(TextBox6.Text))
      Set pFeatureClass = pfeaturelayer.FeatureClass
      If Not pFeatureClass.ShapeType = esriGeometryPoint Then
            MsgBox "First Layer must be a Point Layer"
            Exit Sub
      End If

      Set pDataset = pFeatureClass
      Set pWorkspace = pDataset.Workspace
      Set pWorkspaceEdit = pWorkspace
      Set ptable = pfeaturelayer.FeatureClass
      pWorkspaceEdit.StartEditing False
      pWorkspaceEdit.StartEditOperation
End Sub

Private Sub TerminaEdicion()
      pWorkspaceEdit.StopEditOperation
      pWorkspaceEdit.StopEditing True
      
      'pMxDoc.UpdateContents
      'pMxDoc.ActiveView.Refresh
End Sub

Public Sub ArcosEntrePuntos()
    'Funcion que guarda en un arreglo todos los arcos que existen en una ruta
  Dim pMxDoc As IMxDocument
  Dim pNetworkAnalystExtension As INetworkAnalystExtension
  Dim pNALayer As INALayer
  Dim pTraversalResultQuery As INATraversalResultQuery
  
  Set pMxDoc = ThisDocument
  Set pNetworkAnalystExtension = Application.FindExtensionByName("Network Analyst")
  Set pNALayer = pNetworkAnalystExtension.NAWindow.ActiveAnalysis
  Set pTraversalResultQuery = pNALayer.Context.Result
  

  ReDim ArrArcRuta(pTraversalResultQuery.FeatureClass(esriNETEdge).FeatureCount(Nothing)) As Long
  For j = 1 To pTraversalResultQuery.FeatureClass(esriNETEdge).FeatureCount(Nothing)
    ArrArcRuta(j - 1) = pTraversalResultQuery.FeatureClass(esriNETEdge).GetFeature(j).Value(3)
  Next
  
    
End Sub

Public Sub DistanciasEntrePuntos()
  Dim pMxDoc As IMxDocument
  Dim pNetworkAnalystExtension As INetworkAnalystExtension
  Dim pNALayer As INALayer
  Dim pTraversalResultQuery As INATraversalResultQuery
Dim sadaf As Integer
  
  Set pMxDoc = ThisDocument
  Set pNetworkAnalystExtension = Application.FindExtensionByName("Network Analyst")
  Set pNALayer = pNetworkAnalystExtension.NAWindow.ActiveAnalysis
  Set pTraversalResultQuery = pNALayer.Context.Result
  sadaf = pTraversalResultQuery.FeatureClass(esriNETEdge).FeatureCount(Nothing)
    For j = 1 To NK 'pTraversalResultQuery.FeatureClass(esriNETEdge).FeatureCount(Nothing)
        DistanciaD(j - 1) = pTraversalResultQuery.FeatureClass(esriNETEdge).GetFeature(j).Value(9)
        DistanciaDA(j - 1) = pTraversalResultQuery.FeatureClass(esriNETEdge).GetFeature(j).Value(10)
    Next

  
    
End Sub

Public Sub Cronometro()
    'Se utiliza para llevar el tiempo de ejecucion del algoritmo
    tpo = DateDiff("s", timer1, Now)
    conttpo = Format(Int(tpo / 86400) Mod 24, "00") & ":" & _
              Format(Int(tpo / 3600) Mod 60, "00") & ":" & _
              Format(Int(tpo / 60) Mod 60, "00") & ":" & _
              Format(tpo Mod 60, "00")
              
    Label25.Caption = conttpo
End Sub

Public Function FormatearHora(ByVal StrHora As String) As Date
    Dim minuto As String
    Dim hora As String
    Dim segundo As String
    Dim horaToma As Date
    
    hora = Mid(StrHora, 1, 2)
    minuto = Mid(StrHora, 3, 2)
    segundo = Mid(StrHora, 5, 2)
    
    horaToma = hora & ":" & minuto & ":" & segundo

    
    FormatearHora = horaToma
End Function

Public Function ArcoCoseno(ByVal x As Double) As Double
    ArcoCoseno = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Private Sub OptionButton1_Click()
    OptionButton2.Value = False
    OptionButton1.Value = True
    TextBox11.Text = ""
End Sub

Private Sub OptionButton2_Click()
    OptionButton1.Value = False
    OptionButton2.Value = True
    TextBox11.Text = ""
End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub UserForm_Click()

End Sub


