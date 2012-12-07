Imports System.Threading
Imports System.io
Imports System.Xml
Imports System.Text

Public Class Form1
    Private HiloEjecutor As Thread
    'Delegado para poder acceder al txt del form desde el sub ejecutado por el hilo
    Delegate Sub ImprimirText(ByVal txt As String)
    Delegate Sub FileText(ByVal path As String, ByVal txt As String)
    Private strBuscar1, strReemplazar1 As String
    Private strBuscar2, strReemplazar2 As String
    Private strBuscar3, strReemplazar3 As String

    Private flagExiste1 As Boolean = False
    Private flagExiste2 As Boolean = False

    Private flagSelected As Boolean = False
    Delegate Sub AddItemCallBack(ByVal Item As Object)
    Delegate Function getSelectedCallBack()
    Public dsGlobal As DataSet


    Private Sub ImprimirTexto(ByVal txt As String)
        If Me.InvokeRequired Then
            'Si es necesario utilizar Invoke, llamo al delegado
            Me.Invoke(New ImprimirText(AddressOf ImprimirTexto), New Object() {txt})
        Else
            'Aquí puedo modificar el texto del txt tranquilamente
            Me.txtInformacion.Text = Me.txtInformacion.Text & txt
        End If
    End Sub


    Public Function getSelectedItem()
        If Me.lstCampos.InvokeRequired Then
            Dim d As New getSelectedCallBack(AddressOf getSelectedItem)
            Me.lstCampos.Invoke(d)
        Else
            Dim ItemsSeleccionados As System.Windows.Forms.ListBox.SelectedIndexCollection
            ItemsSeleccionados = Me.lstCampos.SelectedIndices
            ImprimirTexto("Bloqueando / Desbloqueando campos" & vbCrLf)
            For i As Integer = 0 To ItemsSeleccionados.Count - 1
                'MsgBox(lstCampos.Items.Item(ItemsSeleccionados(i)).ToString())
                If Me.chkBloquear.Checked = True Then
                    dsGlobal = modificarCampo(dsGlobal, lstCampos.Items.Item(ItemsSeleccionados(i)).ToString(), "", "locked")
                Else
                    dsGlobal = modificarCampo(dsGlobal, lstCampos.Items.Item(ItemsSeleccionados(i)).ToString(), "", "Descripcion " & lstCampos.Items.Item(ItemsSeleccionados(i)).ToString())
                End If

            Next i
        End If
    End Function

    Private Sub EscribeArchivo(ByVal path As String, ByVal text As String)
        If Me.InvokeRequired Then
            'Si es necesario utilizar Invoke, llamo al delegado
            Me.Invoke(New FileText(AddressOf EscribeArchivo), New Object() {path, text})
        Else
            'Aquí puedo modificar el texto del archivo tranquilamente
            Dim fs As New FileStream(path, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
            Dim sw As New StreamWriter(fs)
            sw.BaseStream().Seek(0, SeekOrigin.End)
            'Escribo el contenido en el objeto Stream
            sw.Write("{0}", text)
            sw.Flush()
            sw.Close()
            fs.Close()
        End If

    End Sub

    Private Sub btnExaminar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExaminar.Click
        Me.ExploradorCarpetas.ShowDialog()
        Me.txtDirectorio.Text = Me.ExploradorCarpetas.SelectedPath
    End Sub


    Private Sub btnIniciar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIniciar.Click
        If Me.txtDirectorio.Text.Trim <> "" Then
            'If Me.txtBuscar.Text.Trim <> "" Then
            Me.btnParar.Enabled = True
            Me.btnIniciar.Enabled = False
            Me.strBuscar1 = Me.txtBuscar.Text.Trim
            Me.strReemplazar1 = Me.txtRemplazar.Text.Trim
            Me.strBuscar2 = Me.txtBuscar2.Text.Trim
            Me.strReemplazar2 = Me.txtRemplazar2.Text.Trim
            Me.strBuscar3 = Me.txtBuscar3.Text.Trim
            Me.strReemplazar3 = Me.txtRemplazar3.Text.Trim

            HiloEjecutor = New Thread(AddressOf IniciaHilo)
            'HiloEjecutor.Priority = ThreadPriority.BelowNormal
            HiloEjecutor.Priority = ThreadPriority.Lowest
            HiloEjecutor.Start()

            Me.btnParar.Enabled = False
            Me.btnIniciar.Enabled = True
            'Else
            '    MsgBox("No ha indicado por que se va a remplazar la busqueda")
            'End If
        Else
        MsgBox("No ha indicado la carpeta raiz")
        End If
    End Sub

    Private Sub IniciaHilo()
        Try
            If Directory.Exists(Me.txtDirectorio.Text.Trim) Then
                Dim Directorio As String()
                ImprimirTexto("Comienza la búsqueda en " & Me.txtDirectorio.Text.Trim & "..." & vbCrLf)
                ImprimirTexto("Hora: " & Now.ToString & vbCrLf)
                Directorio = System.IO.Directory.GetDirectories(Me.txtDirectorio.Text.Trim)

                For i As Integer = 0 To UBound(Directorio, 1)
                    RecorreCarpetas(Directorio(i))
                Next i
                ImprimirTexto("Fin de la búsqueda en " & Me.txtDirectorio.Text.Trim & "..." & vbCrLf)
                ImprimirTexto("Hora: " & Now.ToString & vbCrLf)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
            ImprimirTexto("******************************************************" & vbCrLf)
            ImprimirTexto("************************ ERROR ***********************" & vbCrLf)
            ImprimirTexto(ex.ToString & vbCrLf)
            ImprimirTexto("******************************************************" & vbCrLf)

        End Try
        
    End Sub


    Private Sub obtenerCampos()
        Try

            If Directory.Exists(Me.txtDirectorio.Text.Trim) Then
                Dim Directorio As String()
                ImprimirTexto("Obteniendo campos..." & vbCrLf)
                ImprimirTexto("Hora: " & Now.ToString & vbCrLf)
                Directorio = System.IO.Directory.GetDirectories(Me.txtDirectorio.Text.Trim)

                For i As Integer = 0 To UBound(Directorio, 1)
                    RecorreCarpetas(Directorio(i), True)
                Next i
                ImprimirTexto("Obtención terminada..." & vbCrLf)
                ImprimirTexto("Hora: " & Now.ToString & vbCrLf)

            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
            ImprimirTexto("******************************************************" & vbCrLf)
            ImprimirTexto("************************ ERROR ***********************" & vbCrLf)
            ImprimirTexto(ex.ToString & vbCrLf)
            ImprimirTexto("******************************************************" & vbCrLf)

        End Try
    End Sub

    Private Sub RecorreCarpetas(ByVal path As String, Optional ByVal modoObtencion As Boolean = False)
        'Dim SubDirectorio As String()
        'Dim subSubDirectorio As String()
        'Dim leng As Integer

        If modoObtencion = True Then
            If Me.txtNombreWeb.Text.Trim = "" Then
                If File.Exists(path & "\xml\configuracion.xml") Then
                    ImprimirTexto(" ······ Obteniendo campos del XML" & vbCrLf)
                    obtenerCampos(path & "\xml\configuracion.xml")

                End If
            Else
                If path.ToLower.Substring(path.Replace("/", "\").LastIndexOf("\") + 1) = Me.txtNombreWeb.Text.Trim.ToLower Then
                    If File.Exists(path & "\xml\configuracion.xml") Then
                        ImprimirTexto(" ······ Obteniendo campos del XML" & vbCrLf)
                        obtenerCampos(path & "\xml\configuracion.xml")

                    End If
                End If
            End If
            Return
        End If

        If Me.txtNombreWeb.Text.Trim = "" Then
            ImprimirTexto("Buscando " & path & "\configuracion.xml..." & vbCrLf)
            If File.Exists(path & "\xml\configuracion.xml") Then
                ImprimirTexto(" ······ Encontrado configuracion.xml" & vbCrLf)
                ModificaXML(path & "\xml\configuracion.xml")
                'HiloEjecutor.Abort()
            End If

            ''---------------------- SUBDIRECTORIO (NO LOS NECESITAMOS PARA LOS XML) ---------------------
            'SubDirectorio = Directory.GetDirectories(path)

            'For i As Integer = 0 To UBound(SubDirectorio)
            '    ImprimirTexto("Entrada en " & SubDirectorio(i) & "..." & vbCrLf)

            '    If SubDirectorio(i).ToLower.Substring(SubDirectorio(i).Replace("/", "\").LastIndexOf("\") + 1).Length >= 4 Then
            '        leng = 4
            '    Else
            '        leng = SubDirectorio(i).ToLower.Substring(SubDirectorio(i).Replace("/", "\").LastIndexOf("\") + 1).Length
            '    End If

            '    If SubDirectorio(i).ToLower.Substring(SubDirectorio(i).Replace("/", "\").LastIndexOf("\") + 1, leng) = "foro" OrElse SubDirectorio(i).ToLower.Substring(SubDirectorio(i).Replace("/", "\").LastIndexOf("\") + 1) = "adm" _
            '    OrElse SubDirectorio(i).ToLower.Substring(SubDirectorio(i).Replace("/", "\").LastIndexOf("\") + 1) = "admprimas" OrElse SubDirectorio(i).ToLower.Substring(SubDirectorio(i).Replace("/", "\").LastIndexOf("\") + 1) = "primas" Then
            '        ImprimirTexto("Buscando " & SubDirectorio(i) & "\configuracion.xml..." & vbCrLf)
            '        'If File.Exists(SubDirectorio(i) & "\editBody.xml") Then
            '        '    File.Delete(SubDirectorio(i) & "\editBody.xml")
            '        'End If

            '        If File.Exists(SubDirectorio(i) & "\xml\configuracion.xml") Then
            '            ImprimirTexto(" ······ Encontrado configuracion.xml" & vbCrLf)
            '            ModificaXML(SubDirectorio(i) & "\xml\configuracion.xml")

            '        End If
            '    End If
            'Next i
        Else
            'ImprimirTexto(path.ToLower.Substring(path.Replace("/", "\").LastIndexOf("\") + 1))
            If path.ToLower.Substring(path.Replace("/", "\").LastIndexOf("\") + 1) = Me.txtNombreWeb.Text.Trim.ToLower Then
                ImprimirTexto("Buscando " & path & "\configuracion.xml..." & vbCrLf)
                If File.Exists(path & "\xml\configuracion.xml") Then
                    ImprimirTexto(" ······ Encontrado configuracion.xml" & vbCrLf)
                    ModificaXML(path & "\xml\configuracion.xml")
                    'HiloEjecutor.Abort()
                End If
            End If

            '' --------------------- SUBDIRECTORIOS (NO LO NECESITAMOS) ----------------------
            ''SubDirectorio = Directory.GetDirectories(path)
            ''For i As Integer = 0 To UBound(SubDirectorio)
            ''    ImprimirTexto("Entrada en " & SubDirectorio(i) & "..." & vbCrLf)
            ''    ImprimirTexto("Buscando " & SubDirectorio(i) & "\Web.config..." & vbCrLf)
            ''    If File.Exists(SubDirectorio(i) & "\xml\configuracion.xml") Then
            ''        ImprimirTexto(" ······ Encontrado configuracion.xml" & vbCrLf)
            ''        ModificaXML(SubDirectorio(i) & "\xml\configuracion.xml")
            ''    End If

            ''    'If File.Exists(SubDirectorio(i) & "\log.xml") Then
            ''    '    File.Delete(SubDirectorio(i) & "\log.xml")
            ''    '    ImprimirTexto("Archivo " & SubDirectorio(i) & "\log.xml eliminado" & vbCrLf)
            ''    'End If

            ''    subSubDirectorio = Directory.GetDirectories(SubDirectorio(i))

            ''    For j As Integer = 0 To UBound(subSubDirectorio)
            ''        ImprimirTexto("Entrada en " & subSubDirectorio(j) & "..." & vbCrLf)

            ''        If subSubDirectorio(j).ToLower.Substring(subSubDirectorio(j).Replace("/", "\").LastIndexOf("\") + 1, leng) = "foro" OrElse subSubDirectorio(j).ToLower.Substring(subSubDirectorio(j).Replace("/", "\").LastIndexOf("\") + 1) = "adm" Then
            ''            ImprimirTexto("Buscando " & subSubDirectorio(j) & "\Web.config..." & vbCrLf)
            ''            'If File.Exists(subSubDirectorio(j) & "\editBody.xml") Then
            ''            '    File.Delete(subSubDirectorio(j) & "\editBody.xml")
            ''            '    ImprimirTexto("Archivo " & subSubDirectorio(j) & "\editBody.xml eliminado" & vbCrLf)
            ''            'End If

            ''            If File.Exists(subSubDirectorio(j) & "\xml\configuracion.xml") Then
            ''                ImprimirTexto(" ······ Encontrado configuracion.xml" & vbCrLf)
            ''                ModificaXML(subSubDirectorio(j) & "\xml\configuracion.xml")
            ''            End If
            ''        End If
            ''    Next j
            ''Next i
        End If
    End Sub



    Private Sub ModificaXML(ByVal path As String)
        remplazarXML(path) 'llama a la funcion que realiza los 3 remplazos de texto

        'Dim ds As DataSet = obtenerXML(path)
        dsGlobal = obtenerXML(path)

        'añadimos el primer campo opcional
        If Me.txtNombreCampo1.Text.Trim() <> "" And txtNuevoValor1.Text.Trim() <> "" And txtDescripcion1.Text.Trim() <> "" Then
            ImprimirTexto("Modificando campo 1")
            dsGlobal = modificarCampo(dsGlobal, txtNombreCampo1.Text.Trim(), txtNuevoValor1.Text.Trim(), txtDescripcion1.Text.Trim())
        End If
        'añadimos el segundo campo opcional
        If Me.txtNombreCampo2.Text.Trim() <> "" And txtNuevoValor2.Text.Trim() <> "" And txtDescripcion2.Text.Trim() <> "" Then
            ImprimirTexto("Modificando campo 2")
            dsGlobal = modificarCampo(dsGlobal, txtNombreCampo2.Text.Trim(), txtNuevoValor2.Text.Trim(), txtDescripcion2.Text.Trim())
        End If


        'añadimos el primer campo opcional
        If Me.txtNombreCampo3.Text.Trim() <> "" And txtNuevoValor3.Text.Trim() <> "" And txtDescripcion3.Text.Trim() <> "" Then
            ImprimirTexto("Añadiendo campo 1")
            dsGlobal = añadirCampo(dsGlobal, txtNombreCampo3.Text.Trim(), txtNuevoValor3.Text.Trim(), txtDescripcion3.Text.Trim())
        End If
        'añadimos el segundo campo opcional
        If Me.txtNombreCampo4.Text.Trim() <> "" And txtNuevoValor4.Text.Trim() <> "" And txtDescripcion4.Text.Trim() <> "" Then
            ImprimirTexto("Añadiendo campo 2")
            dsGlobal = añadirCampo(dsGlobal, txtNombreCampo4.Text.Trim(), txtNuevoValor4.Text.Trim(), txtDescripcion4.Text.Trim())
        End If
        If chkBackup.Checked = True Then
            backupXML(path)
        End If


        If Me.lstCampos.Items.Count > 0 Then
            ImprimirTexto("Existen campos" & vbCrLf)

            'ImprimirTexto("Bloqueando campos" & vbCrLf)

            getSelectedItem()

            'ItemsSeleccionados = getSelectedItem()
            'MsgBox(ItemsSeleccionados.Count)
            'For i As Integer = 0 To ItemsSeleccionados.Count
            '    MsgBox(ItemsSeleccionados(i))
            '    ds = modificarCampo(ds, lstCampos.Items.Item(i).ToString(), "", "locked")
            'Next i

            'For i As Integer = 0 To lstCampos.Items.Count - 1
            '    MsgBox(getSelectedItem(i))
            '    'ImprimirTexto("index: " & i & " Selected: " & getSelectedItem(i).ToString())
            '    'ImprimirTexto("Campo: " & lstCampos.Items.Item(i).ToString() & "Selected: " & getSelectedItem(i).ToString() & vbCrLf)
            '    If getSelectedItem(i) = True Then

            '    End If
            'Next
            'ImprimirTexto("Campos bloqueados" & vbCrLf)

            'ImprimirTexto("Desbloqueando campos" & vbCrLf)
            'For i As Integer = 0 To lstCampos.Items.Count - 1
            '    If getSelectedItem(i) Then
            '        ds = modificarCampo(ds, lstCampos.Items.Item(i).ToString(), "", "Descripción" & lstCampos.Items.Item(i).ToString())
            '    End If

            'Next
            'ImprimirTexto("Campos desbloqueados" & vbCrLf)

        End If

        guardarCambios(path, dsGlobal)

    End Sub

    Private Sub backupXML(ByVal path As String)
        If File.Exists(path) Then
            File.Copy(path, path.Replace("configuracion", "configuracion_bak_" & Now.Hour.ToString() & Now.Minute.ToString() & Now.Second.ToString() & Now.DayOfYear().ToString() & Now.Year.ToString()))
            ImprimirTexto("Copia de seguridad realizada" & vbCrLf)
        End If
    End Sub

    Private Sub guardarCambios(ByVal path As String, ByVal ds As DataSet)
        Dim writer As New XmlTextWriter(path, Encoding.UTF8)
        writer.Formatting = Formatting.Indented
        writer.WriteStartElement("config")
        writer.WriteStartElement("valores")

        For i As Integer = 0 To ds.Tables(0).Columns.Count - 1
            writer.WriteElementString(ds.Tables(0).Columns(i).ColumnName, ds.Tables(0).Rows(0)(ds.Tables(0).Columns(i).ColumnName))
        Next
        writer.WriteEndElement()
        writer.WriteStartElement("valores")

        For i As Integer = 0 To ds.Tables(0).Columns.Count - 1
            If ds.Tables(0).Rows.Count > 1 Then
                writer.WriteElementString(ds.Tables(0).Columns(i).ColumnName, ds.Tables(0).Rows(1)(ds.Tables(0).Columns(i).ColumnName))
            End If
        Next
        writer.WriteEndElement()
        writer.WriteEndElement()

        writer.Flush()
        writer.Close()
    End Sub

    Private Function añadirCampo(ByVal ds As DataSet, ByVal nombreCampo As String, ByVal valorCampo As String, ByVal descripcion As String) As DataSet

        If checkExisteCampo(ds, nombreCampo) Then 'Si existe la columna que queremos añadir obviamos el resto de instrucciones
            If Me.chkSobreescribir.Checked = True Then
                ds = modificarCampo(ds, nombreCampo, valorCampo, descripcion)
                Return ds
            Else
                Return ds
            End If

        End If
        ds.Tables(0).Columns.Add(New DataColumn(nombreCampo))
        ds.Tables(0).Rows(0)(nombreCampo) = valorCampo
        If ds.Tables(0).Rows.Count > 1 Then 'Si solo tenemos una fila significa que el XML no tiene descripciones
            ds.Tables(0).Rows(1)(nombreCampo) = descripcion
        End If

        Return ds
    End Function

    Private Function checkExisteCampo(ByVal ds As DataSet, ByVal nombreCampo As String) As Boolean
        For i As Integer = 0 To ds.Tables(0).Columns.Count - 1
            If ds.Tables(0).Columns(i).ColumnName.Equals(nombreCampo) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function modificarCampo(ByVal ds As DataSet, ByVal nombreCampo As String, ByVal valorCampo As String, ByVal descripcion As String) As DataSet
        'MsgBox(ds.Tables(0).Rows(0)(nombreCampo).ToString())

        If Not ds.Tables(0).Columns(nombreCampo) Is Nothing Then
            If valorCampo <> "" Then
                ds.Tables(0).Rows(0)(nombreCampo) = valorCampo
            End If

            If ds.Tables(0).Rows.Count > 1 Then
                'ImprimirTexto(ds.Tables(0).Rows(1)(nombreCampo).ToString())
                'MsgBox(ds.Tables(0).Rows(1)(nombreCampo).ToString())

                ds.Tables(0).Rows(1)(nombreCampo) = descripcion




            End If
        End If
        Return ds
    End Function

    Private Sub remplazarXML(ByVal path As String)
        Dim strFile As String = ""
        Dim rFile As New StreamReader(path)

        strFile = rFile.ReadToEnd
        rFile.Close()
        rFile = Nothing
        If Me.strBuscar1 <> "" Then
            ImprimirTexto("Remplazando valor 1")
            strFile = Replace(strFile, Me.strBuscar1, Me.strReemplazar1)
        End If
        If Me.strBuscar2 <> "" Then
            ImprimirTexto("Remplazando valor 2")
            strFile = Replace(strFile, Me.strBuscar2, Me.strReemplazar2)
        End If
        If Me.strBuscar3 <> "" Then
            ImprimirTexto("Remplazando valor 3")
            strFile = Replace(strFile, Me.strBuscar3, Me.strReemplazar3)
        End If

        EscribeArchivo(path, strFile)
    End Sub

    Public Function obtenerXML(ByVal path As String) As DataSet
        'Cadena con la ruta al archivo XML 
        Dim sourceXML As String
        'Generar una instancia de DataSet 
        Dim dsDatos As New DataSet()
        Try
            'Rellenar el DataSet desde el archivo XML 
            sourceXML = path
            dsDatos.ReadXml(sourceXML)
        Catch
            sourceXML = path
            'Rellenar el DataSet desde el archivo XML 
            dsDatos.ReadXml(sourceXML)
        End Try
        'Escribir el contenido del DataSet en el archivo XML 
        'dsDatos.WriteXml(sourceXML); 
        'Leemos el DataSet 
        Return dsDatos
    End Function

    Private Sub obtenerCampos(ByVal path As String)
        Dim ds As DataSet = obtenerXML(path)

        For i As Integer = 0 To ds.Tables(0).Columns.Count - 1
            If Not existeLista(ds.Tables(0).Columns(i).ColumnName) Then
                AddItemToList(ds.Tables(0).Columns(i).ColumnName.ToString())

            End If
        Next
    End Sub

    Public Sub AddItemToList(ByVal Item As Object)
        If Me.lstCampos.InvokeRequired Then
            Dim d As New AddItemCallBack(AddressOf AddItemToList)
            Me.Invoke(d, New Object() {Item})
        Else
            lstCampos.Items.Add(Item)
        End If

    End Sub



    Private Function existeLista(ByVal valor As String) As Boolean
        For i As Integer = 0 To Me.lstCampos.Items.Count - 1
            If Me.lstCampos.Items.Item(i).Equals(valor) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub btnParar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParar.Click
        HiloEjecutor.Abort()
        Me.btnParar.Enabled = False
        Me.btnIniciar.Enabled = True
    End Sub

    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        End
    End Sub




    Private Sub chkBloqueo1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBloqueo1.CheckedChanged
        If chkBloqueo1.Checked = True Then
            txtDescripcion1.Enabled = False
            txtDescripcion1.Text = "locked"
        Else
            txtDescripcion1.Enabled = True
            txtDescripcion1.Text = ""
        End If
    End Sub


    Private Sub chkBloqueo2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBloqueo2.CheckedChanged
        If chkBloqueo2.Checked = True Then
            txtDescripcion2.Enabled = False
            txtDescripcion2.Text = "locked"
        Else
            txtDescripcion2.Enabled = True
            txtDescripcion2.Text = ""
        End If
    End Sub

    Private Sub chkBloqueo3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBloqueo3.CheckedChanged
        If chkBloqueo3.Checked = True Then
            txtDescripcion3.Enabled = False
            txtDescripcion3.Text = "locked"
        Else
            txtDescripcion3.Enabled = True
            txtDescripcion3.Text = ""
        End If
    End Sub

    Private Sub chkBloqueo4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBloqueo4.CheckedChanged
        If chkBloqueo4.Checked = True Then
            txtDescripcion4.Enabled = False
            txtDescripcion4.Text = "locked"
        Else
            txtDescripcion4.Enabled = True
            txtDescripcion4.Text = ""
        End If
    End Sub



    Private Sub btnObtener_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnObtener.Click
        If Me.txtDirectorio.Text.Trim <> "" Then
            'If Me.txtBuscar.Text.Trim <> "" Then
            Me.btnParar.Enabled = True
            Me.btnIniciar.Enabled = False


            HiloEjecutor = New Thread(AddressOf obtenerCampos)

            HiloEjecutor.Priority = ThreadPriority.BelowNormal
            HiloEjecutor.Start()

            Me.btnParar.Enabled = False
            Me.btnIniciar.Enabled = True

        Else
            MsgBox("No ha indicado la carpeta raiz")
        End If
    End Sub

    Private Sub chkBloquear_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBloquear.CheckedChanged
        If chkBloquear.Checked = True Then
            chkBloquear.Text = "Bloquear"
        Else
            chkBloquear.Text = "Desbloquear"
        End If
    End Sub

    Private Sub btnTodos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTodos.Click

        If lstCampos.Items.Count > 0 Then
            If flagSelected = False Then
                For i As Integer = 0 To lstCampos.Items.Count - 1
                    lstCampos.SelectedIndex = i
                Next
                Me.flagSelected = True
            Else

                lstCampos.ClearSelected()

                Me.flagSelected = False
            End If

        End If
    End Sub
End Class
