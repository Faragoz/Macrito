Attribute VB_Name = "Module1"
'Déclaration des variables globales
Dim remb_source As Worksheet, march_source, rap_source, tcd_source, tmp_src As Worksheet
Dim nbRemb, nbMarch As Integer

Sub openGUI()
' https://www.delftstack.com/fr/howto/vba/try-catch-in-vba/#:~:text=La%20m%C3%A9thode%20Try%2DCatch%20emp%C3%AAche,de%20l'ex%C3%A9cution%20du%20code.
' Hacer try catch
' https://www.automateexcel.com/fr/vba/gestion-erreurs/
    
    GUI.Show
    GUI.ComboBox1.Clear
    
    For Each am In getAccountManagers() 'SortDictionaryByKey(getAccountManagers())
        GUI.ComboBox1.AddItem am
    Next am
    
    GUI.ComboBox1.AddItem "Tous"
    GUI.ComboBox1.ListIndex = 0
    
    GUI.ComboBox2.Clear
    GUI.ComboBox2.AddItem "Envoyer mails"
    GUI.ComboBox2.AddItem "Afficher mails"
'    GUI.ComboBox2.AddItem "Afficher et envoyer mails"
    GUI.ComboBox2.ListIndex = 0
    
    GUI.ComboBox3.Clear
    GUI.ComboBox3.AddItem "Exclure"
    GUI.ComboBox3.AddItem "Inclure"
    GUI.ComboBox3.ListIndex = 0
End Sub

'Sub loadDB()
'    'Crear variable acceso a la requete
'    'Set remb_source = Workbooks("requete_sql_987.xlsx").Sheets("requete_sql_987")
'    'Set march_source = Workbooks("requete_sql_490.xlsx").Sheets("requete_sql_490")
'    Set remb_source = Sheets("Remboursements")
'    Set march_source = Sheets("Marchands")
'    Set rap_source = Sheets("Rapport")
'    Set tcd_source = Sheets("TCD")
'    Set tmp_src = Sheets("Templates")
'
'    'Número de líneas para devoluciones (987)
'    nbRemb = remb_source.Cells(Rows.Count, "D").End(xlUp).Row
'    'Número de líneas para marchands (490)
'    'nbMarch = march_source.UsedRange.Rows.Count
'    nbMarch = march_source.Cells(Rows.Count, "A").End(xlUp).Row
'End Sub

Sub filterMails()
    ' Asignamos qué valores a buscar
    Set rgData = march_source.Range("A1").CurrentRegion
    ' Asignamos dónde lo vamos a buscar
    Set rgCriteriaRange = remb_source.Range("D1:D" & nbRemb)
    ' Copiamos y pegamos la información encontrada
    Set rgCopyToRange = rap_source.Range("A1").CurrentRegion
    
    ' Run AdvancedFilter
    rgData.AdvancedFilter xlFilterCopy, rgCriteriaRange, rgCopyToRange
End Sub

Sub checkEmpty(Optional delete As Boolean = 0)
    ' Cargamos la información
    Set remb_source = ActiveWorkbook.Sheets("Remboursements")
    
    'Número de líneas para devoluciones (987)
    nbRemb = remb_source.Cells(Rows.Count, "D").End(xlUp).Row
    ' Si true entonces se eliminan las filas sin retail website id, si no, se rellenan con la palabra vide
    If delete Then
    ' Revisar para cuando ya no haya valores vacíos
        On Error Resume Next ' https://excel-downloads.com/threads/erreur-1004-pas-de-cellules-correspondantes.129116/
        remb_source.Range("D1:D" & nbRemb).SpecialCells(xlCellTypeBlanks).EntireRow.delete
    Else
        For Each ID In remb_source.Range("D1:D" & nbRemb)
            If IsEmpty(ID) Then
                ID.Value = "vide"
            End If
        Next
    End If
End Sub

Sub formatRemb()
    Application.ScreenUpdating = False
    ' Cargamos la información
    Set remb_source = ActiveWorkbook.Sheets("Remboursements")
    
    'Número de líneas para devoluciones (987)
    nbRemb = remb_source.Cells(Rows.Count, "D").End(xlUp).Row
    ' Cambiar punto por coma
    ' https://forum.excel-pratique.com/excel/remplacer-points-par-virgules-t24504.html
    'remb_source.Range("P2").Replace What:=Separator, Replacement:=Application.DecimalSeparator
    
    ' Revisamos el valor a buscar
    Separator = "."
    If InStr(1, remb_source.Range("P2"), ",") Then Separator = ","
    
    ' Cambiamos el formato de la columna texto para poder identificar los valores
    remb_source.Range("P1:P" & nbRemb).NumberFormat = "@"
    For i = 2 To nbRemb
        ' Reemplazamos el separador decimal por el del sistema
        ' Nota: Utilizamos cInt en lugar de cDec por optimización
        remb_source.Range("P" & i).Value = CDec(Replace(remb_source.Range("P" & i).Value, Separator, Application.DecimalSeparator))
    Next i
    
    'remb_source.Range("P1:P" & nbRemb).NumberFormat = "General"
    
    ' Eliminar entradas con order id y montant duplicadas
    remb_source.Range("A1:X" & nbRemb).RemoveDuplicates Columns:=Array(10, 16), header:=xlYes
    
    ' Recargar nbRemb en caso de que cambie para evitar error en filterMails() (Deprecated?)
    nbRemb = remb_source.Cells(Rows.Count, "D").End(xlUp).Row
    
    ' Agregar columnas
    ' https://www.automateexcel.com/vba/vlookup-xlookup/
    remb_source.Range("Y1").Value = "Récupérable"
    remb_source.Range("Y2:Y" & nbRemb).Formula = "=VLookup(W2,Status!A:B,2,FALSE)"
    
    remb_source.Range("Z1").Value = "Récupéré"
    remb_source.Range("Z2:Z" & nbRemb).Formula = "=VLookup(M2,Status!A:B,2,FALSE)"
    
    remb_source.Range("AA1").Value = "order_id_long"
    remb_source.Range("AA2:AA" & nbRemb).Formula = "=REPLACE(J2,16,1,8)"
    
    ' No se actualiza correctamente, deprecated para esta aplicación
    'remb_source.Range("Z2:Z" & nbRemb) = Application.WorksheetFunction.VLookup(remb_source.Range("W2:W" & nbRemb).Value, Sheets("Status").Range("A:B"), 2, False)
    'remb_source.Range("Z2:Z" & nbRemb) = Application.WorksheetFunction.VLookup(remb_source.Range("M2:M" & nbRemb).Value, Sheets("Status").Range("A:B"), 2, False)
    
    ' Filtrar retour en raison
    Set toFilter = remb_source.Range("W1:W" & nbRemb)
    toFilter.AutoFilter
    toFilter.AutoFilter Field:=1, Criteria1:="=Retour*", Operator:=xlFilterValues, VisibleDropDown:=True
    
    ' Remplazamos columna Récupéré en función de
    For y = 2 To nbRemb
        If toFilter.Cells(y, 1).EntireRow.Hidden = False Then
            remb_source.Range("Z" & y).Formula = "=IF(S" & y & "="""",""non"",""oui"")"
        End If
    Next
    
    ' Eliminar filtro
    remb_source.AutoFilterMode = False
    Application.ScreenUpdating = True
End Sub

Sub createTCD(Optional accountManager As String = Empty, Optional mode As String = "Exclude", Optional debugM As Boolean = False)
    Application.ScreenUpdating = False
    ' Cargamos la información
    Set remb_source = ActiveWorkbook.Sheets("Remboursements")
    Set tcd_source = ActiveWorkbook.Sheets("TCD")
    
    ' TUTO 1 : https://xlbusinesstools.com/modifier-tableaux-croises-dynamiques-excel-avec-vba/
    ' TUTO 2 : https://www.automateexcel.com/fr/vba/guide-tableaux-croises-dynamiques/
    ' ITEMS : https://docs.microsoft.com/fr-fr/office/vba/api/excel.pivotitem
    'https://docs.microsoft.com/fr-fr/office/vba/api/excel.pivottable.addfields
    
    Dim tcd, tcdCache As PivotCache
    Dim tPivot As PivotTable
    
    ' Recuperamos la lista de excepciones
    Dim ex As New Dictionary
    Set ex = getExceptions()

    ' Eliminar TCD si existe
    If tcd_source.PivotTables.Count > 0 Then tcd_source.PivotTables("TCD Remboursements").TableRange2.Clear

    Set tcdCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=remb_source.Cells(1).CurrentRegion)

    ' Créer le tcd à partir du cache
    Set tcd = tcdCache.CreatePivotTable( _
        TableDestination:=tcd_source.Cells(1), _
        TableName:="TCD Remboursements")
        
    ' Variable solo aesthetic lol
    Set tPivot = tcd_source.PivotTables("TCD Remboursements")
    
    tPivot.AddFields _
        RowFields:=Array("pays_origine", "nom_marchand", "raison", "pays_vente", "order_id_long"), _
        PageFields:=Array("Récupérable", "Récupéré", "account_manager")
    
    ' Filtro account_manager
    If accountManager <> Empty Then tPivot.PivotFields("account_manager").CurrentPage = accountManager
    
    ' Filtro oui et flou
    tPivot.PivotFields("Récupérable").EnableMultiplePageItems = True
    tPivot.PivotFields("Récupérable").PivotItems("non").Visible = False
    
    ' Filtro sólo non
    tPivot.PivotFields("Récupéré").CurrentPage = "non"
    
    ' Filtro para raison
    tPivot.PivotFields("raison").EnableMultiplePageItems = True
    With tPivot.PivotFields("raison")
        .ClearAllFilters
        
        ' Desactivamos todas las opciones excepto la última (causa error)
        ' https://stackoverflow.com/questions/29356631/pivottable-how-to-set-all-items-in-filter-to-false
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(.PivotItems(i).Name).Visible = False
        Next i
        
        ' Activamos las raison que nos interesan
        .PivotItems("Echec de livraison").Visible = True
        .PivotItems("Retour reçu par le marchand mais non traité").Visible = True
        .PivotItems("FCL - le marchand n'a pas fourni preuve signé").Visible = True
        
        ' Desactivamos la última opción (mejora para comprobar si la última opción nos interesa)
        lastN = .PivotItems(.PivotItems.Count).Name
        If lastN = "Echec de livraison" Or lastN = "Retour reçu par le marchand mais non traité" Or lastN = "FCL - le marchand n'a pas fourni preuve signé" Then
            .PivotItems(lastN).Visible = True
        Else
            .PivotItems(lastN).Visible = False
        End If
    End With
    
    ' Filtro para excepciones nom_marchand
    tPivot.PivotFields("nom_marchand").EnableMultiplePageItems = True
    With tPivot.PivotFields("nom_marchand")
        If mode = "Include" Then
            ' Pasamos todos los nom_marchand a false (exclude)
            For i = 1 To .PivotItems.Count - 1
                .PivotItems(.PivotItems(i).Name).Visible = False
            Next i
            
            'Variable control último elemento
            Status = False
            
            ' Incluímos todos aquellos que nos interesen
            For Each nm In ex.Keys
                If marchandExists(CStr(nm)) Then
                    .PivotItems(nm).Visible = True
                    If nm = .PivotItems(.PivotItems.Count).Name Then Status = True
                    If debugM Then Debug.Print "Include: "; .PivotItems(nm).Name
                Else
                    If debugM Then Debug.Print "Le marchand "; nm; " n'existe pas"
                End If
            Next nm
            
            ' Actualizamos el último nom_marchand
            .PivotItems(.PivotItems(.PivotItems.Count).Name).Visible = Status
        Else
            For Each nm In ex.Keys
                If marchandExists(CStr(nm)) Then
                    .PivotItems(nm).Visible = False
                    If debugM Then Debug.Print "Exclude: "; .PivotItems(nm).Name
                Else
                    If debugM Then Debug.Print "Le marchand "; nm; " n'existe pas"
                End If
            Next nm
        End If
    End With
    
    ' Valeur somme
    tPivot.AddDataField tcd.PivotFields("montant euro"), "Somme de remboursements", xlSum
    
    ' Trier de plus grand au plus petit
    tPivot.PivotFields("pays_origine").AutoSort xlDescending, "Somme de remboursements"
    tPivot.PivotFields("nom_marchand").AutoSort xlDescending, "Somme de remboursements"
    Application.ScreenUpdating = True
End Sub

Sub reloadTCD()
    ActiveWorkbook.Sheets("TCD").PivotTables("TCD Remboursements").PivotCache.Refresh
End Sub

Sub createReport()
    Dim tcd_source As Worksheet
    Set tcd_source = Sheets("TCD")
    
    Dim tPivot As PivotTable
    Set tPivot = tcd_source.PivotTables("TCD Remboursements")

    Dim Rapport As Range
    Set Rapport = tPivot.RowRange
    ceT = Rapport.Rows.Count

    Application.DisplayAlerts = False
    For Each sht In ThisWorkbook.Sheets
        If sht.Name = "Rapport" Then Sheets("Rapport").delete
    Next sht
    Application.DisplayAlerts = True

    With Rapport
        .Range("B" & ceT).ShowDetail = True
        ActiveSheet.Name = "Rapport"
    End With

    Dim rap_source As Worksheet
    Set rap_source = Sheets("Rapport")
    raT = rap_source.Range("A1").CurrentRegion.Rows.Count
    
    With rap_source
        .Columns("A:C").delete
        .Columns("F:L").delete
        .Columns("G:K").delete
        .Columns("I:J").delete

        .Columns("D").Cut
        .Columns("B").Insert

        .Columns("H").Cut
        .Columns("D").Insert

        .Columns("J").Cut
        .Columns("F").Insert
        
        .Range("K1").Value = "Montant récuperé"
        .Range("K2:K" & raT).Formula = "=IF([@Récupéré]=""oui"", [@[montant euro]],0)"

        .Range("L1").Value = "Commentaire suivi"
        
        .Range("A1").CurrentRegion.Sort key1:=.Columns("C"), Order1:=xlAscending, header:=xlYes, Key2:=Columns("H"), Order2:=xlDescending, header:=xlYes

    End With

End Sub

Public Function getExceptions() As Dictionary
    Dim exc_src As Worksheet
    Set exc_src = ActiveWorkbook.Sheets("Exceptions")
    
    exT = exc_src.Range("A1").CurrentRegion.Rows.Count
    
    Dim ex As New Dictionary
    
    For i = 1 To exT
        If Not ex.Exists(exc_src.Range("A" & i).Value) Then ex.Add exc_src.Range("A" & i).Value, "test"
    Next i
    
    Set getExceptions = ex
End Function

Public Function getData(Optional debugM As Boolean = 0) As Dictionary
    ' Cargamos la información
    Dim tcd_source As Worksheet
    Set tcd_source = ActiveWorkbook.Sheets("TCD")
    ' Nos posicionamos en la hoja para evitar errores
    tcd_source.Activate
    
    ' Cargamos la información de la tabla Pivot
    Dim tPivot As PivotTable
    Set tPivot = tcd_source.PivotTables("TCD Remboursements")
    tPivot.RowAxisLayout 2 ' xlOutlineRow

    ' Seleccionamos la información que nos interesa de la tabla Pivot
    Dim Rapport As Range
    Set Rapport = tPivot.RowRange
    
    ceT = Rapport.Rows.Count
    
    ' On utilise Dictionary en lieu des collections car il est plus rapide, pour utiliser les array il faut savoir les dimensions par avance
    ' https://excelmacromastery.com/vba-dictionary/ LEER diferencia entre late y early binding
'    Dim pod As Object
'    Set pod = CreateObject("Scripting.Dictionary")
'
'    Dim nmd As Object
'    Set nmd = CreateObject("Scripting.Dictionary")
'
'    Dim rad As Object
'    Set rad = CreateObject("Scripting.Dictionary")
'
'    Dim pad As Object
'    Set pad = CreateObject("Scripting.Dictionary")
'
'    Dim oid As Object
'    Set oid = CreateObject("Scripting.Dictionary")

    Dim pod As New Scripting.Dictionary
    Dim nmd As New Scripting.Dictionary
    Dim nmde As New Scripting.Dictionary
    Dim rad As New Scripting.Dictionary
    Dim pad As New Scripting.Dictionary
    Dim oid As New Scripting.Dictionary
    
'    Dim pos As Range
'    Dim po As Variant
'
'    Dim nms As Range
'    Dim nm As Variant
    ' Desactivamos la animación de la pantalla para optimizar
    Application.ScreenUpdating = False
    
    ' Seleccionamos todos los países de origen
    cad = "pays_origine[all]"
    tPivot.PivotSelect cad, xlLabelOnly
    Set pos = Selection.SpecialCells(xlCellTypeConstants)
    
'    tPos = pos.Count
'    tNms = Rapport.Range("B2:B" & ceT).SpecialCells(xlCellTypeConstants).Count
'    tRas = Rapport.Range("C2:C" & ceT).SpecialCells(xlCellTypeConstants).Count
'    tPas = Rapport.Range("D2:D" & ceT).SpecialCells(xlCellTypeConstants).Count
'    tOis = Rapport.Range("E2:E" & ceT).SpecialCells(xlCellTypeConstants).Count
    
    'Debug.Print ceT, tPos, tNms, tRas, tPas, tOis
    
    ' Boucle para buscar en cada país de origen (po)
    For Each po In pos
        ' Seleccionamos todos los nom_marchand donde pays_origine = po
        cad = "nom_marchand[all] pays_origine['" & po & "']"
        tPivot.PivotSelect cad, xlLabelOnly
        Set nms = Selection.SpecialCells(xlCellTypeConstants)
        
        ' Boucle...
        For Each nm In nms
            ' Si existe un ' en el nom_marchand lo reemplazamos por '' para evitar error sintaxis
            If InStr(1, nm, "'") Then nm = Replace(nm, "'", "''")
            
            ' Seleccionamos...
            cad = "raison[all] nom_marchand['" & nm & "'] pays_origine['" & po & "']"
            tPivot.PivotSelect cad, xlLabelOnly
            
            Set ras = Selection.SpecialCells(xlCellTypeConstants)
            
            For Each ra In ras
                If InStr(1, ra, "'") Then ra = Replace(ra, "'", "''")
                cad = "pays_vente[All] raison['" & ra & "'] nom_marchand['" & nm & "'] pays_origine['" & po & "']"
                tPivot.PivotSelect cad, xlLabelOnly
                Set pas = Selection.SpecialCells(xlCellTypeConstants)
                
                For Each pa In pas
                    cad = "order_id_long[All] pays_origine['" & po & "'] nom_marchand['" & nm & "'] raison['" & ra & "'] pays_vente['" & pa & "']"
                    tPivot.PivotSelect cad, xlLabelOnly, True
                    Set ois = Selection
                    
                    For Each oi In ois
                        ' Anidamos todas las oi correspondientes. Como es un dictionary debemos de indicar key e item, entonces lo dejamos empty
                        oid.Add oi, Empty
                        ' Si el debugD es true imprimimos los resultados
                        If debugM Then Debug.Print po; Left(nm, 10), Left(ra, 15); pa; oi
                    Next oi
                    ' Anidamos los valores del dictionary oid en el dictionary pad
                    pad.Add pa, oid
                    ' Creamos nuevo dictionary: DIFERENTE a método .RemoveAll (revisar early / late binding; http://net-informations.com/faq/oops/binding.htm)
                    Set oid = New Dictionary

                Next pa
                rad.Add ra, pad
                Set pad = New Dictionary
                
            Next ra
            nmd.Add nm, rad
            Set rad = New Dictionary
            
        Next nm
        
        Set nmde = New Dictionary
        ' Para anidar los pays_origine debemos de revisar si existía previamente y eliminarlo
        If pod.Exists(po) Then pod.Remove po
        
        pod.Add po, nmd
        Set nmd = New Dictionary
    Next po
    
    ' Reactivamos animación para uso normal
    Application.ScreenUpdating = True
    
    ' Set return value
    Set getData = pod
End Function

Sub resetFormat()
    ActiveWorkbook.Sheets("TCD").PivotTables("TCD Remboursements").RowAxisLayout 0  ' xlOutlineRow
End Sub

Public Function marchandExists(nm As String) As Boolean
    Dim rgData As Range
    
    Dim remb_src As Worksheet
    Set remb_src = ActiveWorkbook.Sheets("Remboursements")
    
    ' Asignamos qué valores a buscar
    Set rgData = remb_src.Range("A1").CurrentRegion
    ceM = rgData.Rows.Count
    
    ' Asignamos dónde lo vamos a buscar
    rgData.AutoFilter 5, nm
   
    'Return mail found
    On Error GoTo notExist:
    getMarch = rgData.Range("E2:E" & ceM).SpecialCells(xlVisible).Value
    
    If getMarch <> Empty Or getMarch <> "" Then
        marchandExists = True
    Else
        GoTo notExist:
    End If
    
    ' Reset filter
    remb_src.ShowAllData
    Exit Function

notExist:
    marchandExists = False
    remb_src.ShowAllData
    Exit Function
End Function

Public Function getMail(nm As String, Optional testMail As String) As String
    ' Si test mail return y finalizamos funcion
    If testMail <> Empty Then
        getMail = testMail
        Exit Function
    End If
    
    Dim rgData As Range
    
    Dim march_src As Worksheet
    Set march_src = ActiveWorkbook.Sheets("Marchands")
    
    ' Asignamos qué valores a buscar
    Set rgData = march_src.Range("A1").CurrentRegion
    ceM = rgData.Rows.Count
    
    ' Asignamos dónde lo vamos a buscar
    rgData.AutoFilter 2, nm

    'Return mail found
    On Error GoTo Alert:
    getMail = rgData.Range("K2:K" & ceM).SpecialCells(xlVisible).Value
    
    If getMail = Empty Or getMail = "" Then
        getMail = rgData.Range("G2:G" & ceM).SpecialCells(xlVisible).Value
        If getMail = Empty Or getMail = "" Then GoTo Alert
    End If
    
    ' Reset filter
    march_src.ShowAllData
    Exit Function
    
Alert:
    getMail = "NOT_FOUND"
    MsgBox "Mail not found: " & nm & vbNewLine & vbNewLine & "Please, provide the address mail in the Marchands sheet.", vbCritical, "Mail not found"
    Exit Function
End Function

Function getAccountManagers() As Dictionary
    Set getAccountManagers = CreateObject("Scripting.Dictionary")
    ' Cargamos la información
    Set remb_source = ActiveWorkbook.Sheets("Remboursements")
    
    'Número de líneas para devoluciones (987)
    nbRemb = remb_source.Cells(Rows.Count, "D").End(xlUp).Row

    Dim oid As Object
    Set oid = CreateObject("Scripting.Dictionary")
    
    Total = 1000
    If CInt(0.2 * nbRemb) > Total Then Total = CInt(0.2 * nbRemb)
    

    For i = 2 To nbRemb
            If Not getAccountManagers.Exists(remb_source.Range("H" & i).Value) Then
                'Debug.Print i, j, getAccountManagers.Count, Not getAccountManagers.Exists(remb_source.Range("H" & i).Value), getAccountManagers.Item(remb_source.Range("H" & i).Value)
                getAccountManagers.Add remb_source.Range("H" & i).Value, i
            End If
    Next i
    'Debug.Print getAccountManagers.Count
End Function

Public Function getAccountManager() As String
    Set tcd_source = ActiveWorkbook.Sheets("TCD")
    Dim Name As String
    Dim tcd, tcdCache As PivotCache
    Dim tPivot As PivotTable
    
    ' Variable solo aesthetic lol
    Set tPivot = tcd_source.PivotTables("TCD Remboursements")
    
    For Each ac In tPivot.PivotFields("account_manager").PivotItems
        If ac.Visible Then Name = ac.Name
    Next ac
    
    If Name = Empty Then Name = "Tous"
    getAccountManager = Name
End Function


