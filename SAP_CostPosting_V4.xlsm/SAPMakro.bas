Attribute VB_Name = "SAPMakro"

Sub SAP_RepstPrimCosts_post()
    SAP_RepstPrimCosts_exec ("post")
End Sub

Sub SAP_RepstPrimCosts_check()
    SAP_RepstPrimCosts_exec ("check")
End Sub

Sub SAP_RepstPrimCosts_exec(p_mode As String)
    Dim aSAPAcctngRepstPrimCosts As New SAPAcctngRepstPrimCosts
    Dim aDateFormatString As New DateFormatString
    Dim aSAPDocItem As New SAPDocItem
    Dim aData As New Collection
    Dim aRetStr As String

    Dim aSAPUser As New SAPUser
    Dim bRetStr As String

    Dim aKOKRS As String
    Dim aEB As String
    Dim aFromLine As Integer
    Dim aToLine As Integer

    Dim aBLDAT As String
    Dim aBUDAT As String
    Dim aNextBUDAT As String
    Dim aMENGE As String
    Dim aEPSP As String
    Dim aSKOSTL As String
    Dim aLEART As String

    Worksheets("Parameter").Activate
    aKOKRS = Format(Cells(2, 2), "0000")
    aEB = Cells(3, 2)
    If IsNull(aKOKRS) Or aKOKRS = "" Then
        MsgBox "Bitte alle Mussfelder der Parameter füllen!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connection to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If

    Worksheets("Data").Activate
    i = 3
    Do
        If InStr(Cells(i, 21), "Beleg wird unter der Nummer") = 0 And InStr(Cells(i, 21), "Document is posted under number") = 0 Then
            aBUDAT = Format(Cells(i, 1), aDateFormatString.getString)
            aBLDAT = Format(Cells(i, 2), aDateFormatString.getString)
            aNextBUDAT = Format(Cells(i + 1, 1), aDateFormatString.getString)
            Set aSAPDocItem = New SAPDocItem
            aSAPDocItem.create Cells(i, 3).Value, Cells(i, 4).Value, Cells(i, 5).Value, Cells(i, 6).Value, _
            Cells(i, 7).Value, Cells(i, 8).Value, Cells(i, 9).Value, _
            Cells(i, 10).Value, Cells(i, 11).Value, CDbl(Cells(i, 12).Value), _
            Cells(i, 13).Value, Cells(i, 14).Value, Cells(i, 15).Value, Cells(i, 16).Value, _
            Cells(i, 17).Value, Cells(i, 18).Value, Cells(i, 19).Value, Cells(i, 20).Value
            aData.Add aSAPDocItem
            If aEB = "J" Or aEB = "Y" Or aBUDAT <> aNextBUDAT Then
                If p_mode = "post" Then
                    aRetStr = aSAPAcctngRepstPrimCosts.post(aKOKRS, aBUDAT, aBLDAT, aData)
                Else
                    aRetStr = aSAPAcctngRepstPrimCosts.check(aKOKRS, aBUDAT, aBLDAT, aData)
                End If
                Cells(i, 21) = aRetStr
                Set aData = New Collection
            End If
        End If
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
End Sub

