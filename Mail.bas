Attribute VB_Name = "Module2"
'https://www.automateexcel.com/fr/vba/envoyer-courriels-depuis-excel-via-outlook/

Sub sendMail(toa As String, CC As String, title As String, body As String, Optional test As Boolean = False)
    Dim appOutlook As Object
    Dim mItem As Object
    
    If toa = "" Or title = "" Or body = "" Then 'Or CC = ""
        MsgBox "Mail vacio", vbCritical ' + vbMsgBoxHelpButton
        Exit Sub
    End If
    
' Créez une nouvelle instance d'Outlook
    Set appOutlook = CreateObject("Outlook.Application")
    Set mItem = appOutlook.CreateItem(0)
    
    With mItem
     .To = toa
     .CC = CC
     .Subject = title
     '.BodyFormat = olFormatRichText
     .body = body
     '.Attachments.Add ActiveWorkbook.FullName
' Utilisez send pour envoyer immédiatement ou display pour afficher à l'écran
    End With
    
    If test = True Then
        mItem.display
    Else
        mItem.Send
    End If
     
' Nettoyage des objets
    Set mItem = Nothing
    Set appOutlook = Nothing

End Sub

Sub composeMail(dict As Dictionary, Optional testMail As String, Optional debugM As Boolean = False, Optional mode As String = "Envoyer mails", Optional simple As Boolean = False)
    Dim tmp_src As Worksheet
    Set tmp_src = ActiveWorkbook.Sheets("Templates")
    
    Dim title As String
    Dim body As String
    
    For Each po In dict.Keys
    
        langue = po
        If tmp_src.Cells(po, 1) = Empty Then langue = 2
        For Each nm In dict.Item(po)
            ' Como nm está parseado para la búsqueda de información en getData no encontrará el mail en getMail si nm contiene '', entonces lo cambiamos por '
            nmt = nm
            If InStr(1, nmt, "''") Then nmt = Replace(nmt, "''", "'")
            
            
            title = tmp_src.Range("AE" & langue) & nmt
            header = tmp_src.Range("AF" & langue)
            hello = tmp_src.Range("AG" & langue)
            first = tmp_src.Range("AH" & langue)
            body = ""
            last = tmp_src.Range("AO" & langue)
            mail = getMail(CStr(nmt), testMail)
            CC = tmp_src.Range("AQ" & langue)
            
            If testMail <> Empty Then
                mail = testMail
                CC = Empty
            End If
                
            If mail = "NOT_FOUND" Then Exit For
            
            For Each ra In dict.Item(po)(nm)
                If ra = "Echec de livraison" Then
                    body = body & composeRaison(tmp_src.Range("AI" & langue)) & composeExpli(tmp_src.Range("AJ" & langue))
                ElseIf ra = "Retour reçu par le marchand mais non traité" Then
                    body = body & composeRaison(tmp_src.Range("AM" & langue)) & composeExpli(tmp_src.Range("AN" & langue))
                ElseIf ra = "FCL - le marchand n''a pas fourni preuve signé" Then
                    body = body & composeRaison(tmp_src.Range("AK" & langue)) & composeExpli(tmp_src.Range("AL" & langue))
                End If
                
                'Debug.Print tmp_src.Range("K" & langue)
                For Each pa In dict.Item(po)(nm)(ra)
                    body = body & composePa(CStr(tmp_src.Cells(langue, pa)))
                    For Each oi In dict.Item(po)(nm)(ra)(pa)
                        body = body & composeOi(CStr(oi))
                        ' Si el debugM es true imprimimos los resultados
                        If debugM Then Debug.Print po; Left(nmt, 10), Left(ra, 15); pa, oi
                    Next oi
                Next pa
            Next ra

        'If debugM Then
        Debug.Print "Mail envoyé à "; nmt; " [" & mail & "] CC: [" & CC & "]"
        sendMailTemplate CStr(mail), CStr(CC), title, CStr(header), CStr(hello), CStr(nmt), CStr(first), CStr(body), CStr(last), getAccountManager(), mode, simple
        Next nm
    Next po
End Sub

' https://www.ablebits.com/office-addins-blog/outlook-email-templates-fillable-fields-dropdown/
' https://my.stripo.email/
' https://answers.microsoft.com/en-us/outlook_com/forum/all/i-want-to-insert-variable-data-into-an-outlook/d89d9b84-b9ef-4103-acf4-b689bf9baa8c
Sub sendMailTemplate(toa As String, CC As String, title As String, header As String, hello As String, nom_marchand As String, first As String, body As String, last As String, nomac As String, Optional mode As String = "Envoyer mails", Optional simple As Boolean = False)
    Dim appOutlook As Object
    Dim mItem As Object
 
    Set appOutlook = CreateObject("Outlook.Application")
    Set mItem = appOutlook.CreateItemFromTemplate(Application.ActiveWorkbook.Path & "\template.oft")

    If simple Then
        mItem.HTMLBody = "<html><head><meta charset=""UTF-8""><title>Spartoo</title></head><body>[hello] [nomm], <br><br>[first]<br><br><table><tbody>[body]</tbody></table> <br>[last]<br>[nomac] <br>Account Manager, Spartoo</body></html>"
    End If
    
    With mItem
    
        If InStr(1, toa, ",") Then
            ' https://forums.commentcamarche.net/forum/affich-12437010-vb-excel-outlook-liste-adresses-mails
            For Each mail In Split(toa, ",")
                If mail Like "?*@?*.?*" Then
                    mItem.Recipients.Add mail
                Else
                    MsgBox "Le format de l'addresse mail " & mail & " n'est pas valide (marchand: " & nom_marchand & "), elle a été supprimée de la liste de contacts", vbCritical, "Invalid email"
                End If
            Next mail
        Else
            If toa Like "?*@?*.?*" Then
                .To = toa
            Else
                MsgBox "Le format de l'addresse mail " & mail & " n'est pas valide (marchand: " & nom_marchand & "), elle a été supprimée de la liste de contacts", vbCritical, "Invalid email"
            End If
        End If
        
        
        .CC = CC
        .Subject = title
        .HTMLBody = Replace(mItem.HTMLBody, "[header]", header)
        .HTMLBody = Replace(mItem.HTMLBody, "[hello]", hello)
        .HTMLBody = Replace(mItem.HTMLBody, "[nomm]", nom_marchand)
        .HTMLBody = Replace(mItem.HTMLBody, "[first]", first)
        .HTMLBody = Replace(mItem.HTMLBody, "[body]", body)
        .HTMLBody = Replace(mItem.HTMLBody, "[last]", last)
        .HTMLBody = Replace(mItem.HTMLBody, "[nomac]", nomac)
        
        Select Case mode
        Case "Envoyer mails"
            .Send
        Case "Afficher mails"
            .display
        Case "Afficher et envoyer mails"
            .display
            .Send
        End Select
    End With
    
    Set mItem = Nothing
    Set appOutlook = Nothing
End Sub

Function composeRaison(raison As String) As String
    composeRaison = "<br> <tr style=""border-collapse:collapse""><td bgcolor=""#282626"" align=""left"" style=""Margin:0;padding-top:10px;padding-bottom:10px;padding-left:10px;padding-right:10px""><table style=""mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;width:500px"" class=""cke_show_border"" cellspacing=""1"" cellpadding=""1"" border=""0"" align=""left"" role=""presentation""><tr style=""border-collapse:collapse""><td width=""80%"" style=""padding:10px 0 0 0;Margin:0""><h4 style=""Margin:0;line-height:120%;mso-line-height-rule:exactly;font-family:'open sans', 'helvetica neue', helvetica, arial, sans-serif;color:#fbf5f5"">" & raison & "</h4></td></tr></table></td></tr>"
End Function

Function composeExpli(expli As String) As String
    composeExpli = "<tr style=""border-collapse:collapse""><td bgcolor=""#eeeeee"" align=""left"" style=""Margin:0;padding-top:10px;padding-bottom:10px;padding-left:10px;padding-right:10px""><table style=""mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;width:500px"" class=""cke_show_border"" cellspacing=""1"" cellpadding=""1"" border=""0"" align=""left"" role=""presentation""><tr style=""border-collapse:collapse""><td width=""80%"" style=""padding:0;Margin:0""><p style=""Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:'open sans', 'helvetica neue', helvetica, arial, sans-serif;line-height:23px;color:#333333;font-size:15px"">" & expli & "</p></td></tr></table></td></tr>"
End Function

Function composePa(pa As String) As String
    composePa = "<tr style=""border-collapse:collapse""><td align=""left"" style=""Margin:0;padding-top:10px;padding-bottom:10px;padding-left:10px;padding-right:10px""><table style=""mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;width:500px"" cellspacing=""1"" cellpadding=""1"" border=""0"" align=""left"" class=""cke_show_border"" role=""presentation""><thead><tr style=""border-collapse:collapse""><th style=""padding:5px 10px 5px 0px"" width=""80%"" align=""left"" scope=""row""><p style=""Margin:0;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:'open sans', 'helvetica neue', helvetica, arial, sans-serif;line-height:23px;color:#333333;font-size:15px"">" & pa & "</p></th></tr></thead></table></td></tr>"
End Function

Function composeOi(oi As String) As String
    composeOi = "<tr style=""border-collapse:collapse""><td align=""left"" style=""Margin:0;padding-top:0px;padding-bottom:0px;padding-left:10px;padding-right:0px""><table style=""mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;width:500px"" class=""cke_show_border"" cellspacing=""1"" cellpadding=""1"" border=""0"" align=""left"" role=""presentation""><tr style=""border-collapse:collapse""><td width=""80%"" style=""padding:0;Margin:0""><p style=""Margin:0 0 0 20px;-webkit-text-size-adjust:none;-ms-text-size-adjust:none;mso-line-height-rule:exactly;font-family:'open sans', 'helvetica neue', helvetica, arial, sans-serif;line-height:23px;color:#333333;font-size:15px"">" & oi & "</p></td></tr></table></td></tr>"
End Function



