Attribute VB_Name = "SQL_Calls"
Public Sub init_db()
    'THIS ROUTINE OPENS THE CONNECTION TO THE SQL SERVER AND INITIALIZES
    'ALL RECORDSETS, ALSO SETS THE GLOBAL MESSAGE
    Set M2MConn = New ADODB.Connection
    M2MConn.Open "Provider=sqloledb;" & _
           "Data Source=" + M2MSERVER + ";" & _
           "Initial Catalog=" + M2MDB + ";" & _
           "User Id=" + M2MUSER + ";" & _
           "Password=" + M2MPASS
    Set M2MEmp = New ADODB.Recordset
    Set M2MJobs = New ADODB.Recordset
    Set M2MDesc = New ADODB.Recordset
    Set M2MJobs = New ADODB.Recordset
    Set M2MPart = New ADODB.Recordset
    Set M2MEff = New ADODB.Recordset
    Set M2MMsg = New ADODB.Recordset
    Set M2MLocations = New ADODB.Recordset
    'Set viewMsg = New ADODB.Command
End Sub

Function emp_login(empID As String) As String
    'CHECKS IF EMPLOYEE IS VALID
    'CHECKS FOR NEW MESSAGES
    If M2MEmp.State = adStateOpen Then
        M2MEmp.Close
    End If
    M2MEmp.Open "SELECT FNAME,FFNAME FROM PREMPL WHERE FEMPNO = '" + empID + "'", M2MConn, adOpenKeyset, adLockReadOnly
    empNumber = empID
    'CHECK FOR VALID EMPLOYEE
    If M2MEmp.RecordCount = 1 Then
        emp_login = Trim(M2MEmp.Fields(0)) + ", " + Trim(M2MEmp.Fields(1))
    Else
        emp_login = "INVALID EMPLOYEE"
    End If
End Function
Public Sub Populate_Groups()
    If M2MJobs.State = adStateOpen Then
        M2MJobs.Close
    End If
    M2MJobs.Open "SELECT DISTINCT JOMAST.FJOB_NAME FROM JOMAST " & _
                  "INNER JOIN JODRTG ON JOMAST.FJOBNO = JODRTG.FJOBNO " & _
                  "INNER JOIN INWORK ON JODRTG.FPRO_ID = INWORK.FCPRO_ID " & _
                  "WHERE (JOMAST.FSTATUS = 'RELEASED') " & _
                  "AND (JOMAST.FJOBNO <> 'I3104-0000') " & _
                  "AND (INWORK.FDEPT = '" + PLANT + "') " & _
                  "AND JOMAST.FITYPE = 1 " & _
                  "ORDER BY JOMAST.FJOB_NAME", M2MConn, adOpenKeyset, adLockReadOnly
                  
                  
    Dim itmx As ListItem
    While Not M2MJobs.EOF
        Set itmx = JobGroup.List1.ListItems.Add(, , M2MJobs.Fields(0))
        M2MJobs.MoveNext
    Wend
End Sub
Public Sub Populate_Jobs()
    If M2MJobs.State = adStateOpen Then
        M2MJobs.Close
    End If
    M2MJobs.Open "SELECT DISTINCT JOMAST.FJOBNO, JOMAST.FPARTNO " & _
                  "FROM JOMAST " & _
                  "INNER JOIN JODRTG ON JOMAST.FJOBNO = JODRTG.FJOBNO " & _
                  "INNER JOIN INWORK ON JODRTG.FPRO_ID = INWORK.FCPRO_ID " & _
                  "WHERE (JOMAST.FSTATUS = 'RELEASED') " & _
                  "AND (JOMAST.FJOB_NAME = '" + JobGrp + "') " & _
                  "AND (INWORK.FDEPT = '" + PLANT + "') " & _
                  "AND JOMAST.FITYPE = 1 " & _
                  "ORDER BY JOMAST.FPARTNO", M2MConn, adOpenKeyset, adLockReadOnly
    If M2MJobs.RecordCount = 1 Then
        global_jobnumber = M2MJobs.Fields(0)
        Verify.JobNum = global_jobnumber
        Verify.PartNum = GetPartNumber(global_jobnumber)
        Populate_Ops (global_jobnumber)
        Exit Sub
    End If

    Dim itmx As ListItem
    While Not M2MJobs.EOF
        Set itmx = JobList.List1.ListItems.Add(, , M2MJobs.Fields(0))
        itmx.SubItems(1) = Trim(M2MJobs.Fields(1))
        M2MJobs.MoveNext
    Wend
    JobList.Show
End Sub
' EnterCount()
' Uses JobGrp and PartRev selected from PartGroup.ListView object to retrieve the JOMAST M2M database
' record which will be updated with the scrap count
Public Sub Enter_Count()
    If M2MJobs.State = adStateOpen Then
        M2MJobs.Close
    End If
    ' REPLACE ONE SINGLE QUOTES WITH TWO SINGLE QUOTES OR QUERY WILL FAIL
    Dim strJobName As String
    Dim strQuote As String
    strQuote = "'"
    Dim strDoubleQuote As String
    strDoubleQuote = "''"
    strJobName = Replace(global_jobnumber, strQuote, strDoubleQuote)
    M2MJobs.Open "SELECT DISTINCT JOMAST.FJOBNO, JOMAST.FPARTNO, JOMAST.FPARTREV " & _
                  "FROM JOMAST " & _
                  "INNER JOIN JODRTG ON JOMAST.FJOBNO = JODRTG.FJOBNO " & _
                  "INNER JOIN INWORK ON JODRTG.FPRO_ID = INWORK.FCPRO_ID " & _
                  "WHERE (JOMAST.FSTATUS = 'RELEASED') " & _
                  "AND (JOMAST.FJOB_NAME = '" + strJobName + "') " & _
                  "AND (JOMAST.FPARTNO = '" + gPartNumber + "') " & _
                  "AND (JOMAST.FPARTREV = '" + gPartRev + "') " & _
                  "AND (INWORK.FDEPT = '" + PLANT + "') " & _
                  "ORDER BY JOMAST.FPARTNO", M2MConn, adOpenKeyset, adLockReadOnly
    If M2MJobs.RecordCount = 1 Then
        global_jobnumber = M2MJobs.Fields(0)
        Verify.JobNum = global_jobnumber
        'Verify.PartNum = Trim(PartNumber)
        Verify.PartNum = M2MJobs.Fields(1)
        Populate_Ops (global_jobnumber)
    End If
End Sub
Public Sub Populate_Parts()
    If M2MJobs.State = adStateOpen Then
        M2MJobs.Close
    End If
    M2MJobs.Open "SELECT DISTINCT JOMAST.FPARTNO, JOMAST.FPARTREV, JOMAST.FJOB_NAME FROM JOMAST " & _
                  "INNER JOIN JODRTG ON JOMAST.FJOBNO = JODRTG.FJOBNO " & _
                  "INNER JOIN INWORK ON JODRTG.FPRO_ID = INWORK.FCPRO_ID " & _
                  "WHERE (JOMAST.FSTATUS = 'RELEASED') " & _
                  "AND (JOMAST.FJOBNO <> 'I3104-0000') " & _
                  "AND (INWORK.FDEPT = '" + PLANT + "') " & _
                  "AND JOMAST.FITYPE = 1 " & _
                  "ORDER BY JOMAST.FPARTNO, JOMAST.FPARTREV", M2MConn, adOpenKeyset, adLockReadOnly
    PartGroup.Show
    
    Dim itmx As ListItem
    While Not M2MJobs.EOF
        Set itmx = PartGroup.List1.ListItems.Add(, , M2MJobs.Fields(0))
        itmx.SubItems(1) = Trim(M2MJobs.Fields(1))
        itmx.SubItems(2) = Trim(M2MJobs.Fields(2))
        M2MJobs.MoveNext
    Wend
    PartGroup.Show
End Sub
Public Sub Populate_Emps()
    If M2MJobs.State = adStateOpen Then
        M2MJobs.Close
    End If
    M2MJobs.Open "SELECT PREMPL.FEMPNO, PREMPL.FNAME, PREMPL.FFNAME FROM PREMPL " & _
                  "WHERE (PREMPL.FENDATE = '1/1/1900') " & _
                  "ORDER BY PREMPL.fempno", M2MConn, adOpenKeyset, adLockReadOnly

    Dim itmx As ListItem
    While Not M2MJobs.EOF
        Set itmx = EmpList.List1.ListItems.Add(, , M2MJobs.Fields(0))
        itmx.SubItems(1) = Trim(M2MJobs.Fields(1)) + ", " + Trim(M2MJobs.Fields(2))
        M2MJobs.MoveNext
    Wend

End Sub
Public Sub Populate_ScrapCodeList()
    If M2MJobs.State = adStateOpen Then
        M2MJobs.Close
    End If
    M2MJobs.Open "SELECT QAINSP.FCODE, QAINSP.FDESC FROM QAINSP " & _
                  "WHERE (SUBSTRING(QAINSP.FCODE,2,2) <> '-') OR " & _
                  "(QAINSP.FCODE <> 'PASS') " & _
                  "ORDER BY QAINSP.FCODE", M2MConn, adOpenKeyset, adLockReadOnly

    Dim itmx As ListItem
    While Not M2MJobs.EOF
        Set itmx = ScrapCodeList.List1.ListItems.Add(, , M2MJobs.Fields(0))
        itmx.SubItems(1) = Trim(M2MJobs.Fields(1))
        M2MJobs.MoveNext
    Wend

End Sub

Public Sub Populate_Locations()
    If M2MLocations.State = adStateOpen Then
        M2MLocations.Close
    End If
    
    M2MLocations.Open "SELECT fLocation, fLocDesc " & _
                            "FROM Location " & _
                            "ORDER BY fLocation", M2MConn, adOpenKeyset, adLockReadOnly

    Dim itmx As ListItem
    While Not M2MLocations.EOF
        Set itmx = Locations.lstAvailable.ListItems.Add(, , M2MLocations.Fields(0))
        itmx.SubItems(1) = Trim(M2MLocations.Fields(1))
        M2MLocations.MoveNext
    Wend

End Sub

Function GetDescription(joNum As String) As String
    If M2MDesc.State = adStateOpen Then
        M2MDesc.Close
    End If
    M2MDesc.Open "SELECT JOITEM.FDESC FROM JOITEM WHERE JOITEM.FJOBNO = '" + joNum + "'", M2MConn, adOpenKeyset, adLockReadOnly
    GetDescription = M2MDesc.Fields(0)
End Function
Function GetDescription2(joNum As String) As String
    If M2MDesc.State = adStateOpen Then
        M2MDesc.Close
    End If
    M2MDesc.Open "SELECT JOITEM.FDESCMEMO FROM JOITEM WHERE JOITEM.FJOBNO = '" + joNum + "'", M2MConn, adOpenKeyset, adLockReadOnly
    GetDescription2 = M2MDesc.Fields(0)
End Function

Public Sub Populate_Ops(jobnumber As String)
    If M2MJobs.State = adStateOpen Then
        M2MJobs.Close
    End If
    M2MJobs.Open "SELECT JODRTG.FOPERNO, INWORK.FCPRO_NAME, JODRTG.FOPERMEMO " & _
                  "FROM JODRTG " & _
                  "INNER JOIN INWORK ON JODRTG.FPRO_ID = INWORK.FCPRO_ID " & _
                  "WHERE (JODRTG.FJOBNO = '" + jobnumber + "') " & _
                  "ORDER BY JODRTG.FOPERNO", M2MConn, adOpenKeyset, adLockReadOnly
    If M2MJobs.RecordCount = 1 Then
        Verify.OperationNumber = Trim(M2MJobs.Fields(0))
        opNumber = Trim(M2MJobs.Fields(0))
        Pieces.Show
        Exit Sub
    End If
    Dim itmx2 As ListItem
    
    
    While Not M2MJobs.EOF
        Set itmx2 = OpList.List1.ListItems.Add(, , M2MJobs.Fields(0))
        itmx2.SubItems(1) = Trim(M2MJobs.Fields(1))
        itmx2.SubItems(2) = Trim(M2MJobs.Fields(2))
        M2MJobs.MoveNext
    Wend
    OpList.Show
End Sub


Public Sub Populate_BOMMtl()
    If M2MJobs.State = adStateOpen Then
        M2MJobs.Close
    End If
    M2MJobs.Open "SELECT JODBOM.fBOMPart, JODBOM.fBOMRev " & _
                  "FROM JODBOM " & _
                  "WHERE (JODBOM.FJOBNO = '" + global_jobnumber + "') " & _
                  "ORDER BY JODBOM.fBOMINUM", M2MConn, adOpenKeyset, adLockReadOnly
                  
    Dim itmx2 As ListItem
    If M2MJobs.RecordCount >= 1 Then
    
        While Not M2MJobs.EOF
            Set itmx2 = BOMList.List1.ListItems.Add(, , M2MJobs.Fields(0))
            itmx2.SubItems(1) = Trim(M2MJobs.Fields(1))
            'itmx2.SubItems(2) = Trim(M2MJobs.Fields(2))
            M2MJobs.MoveNext
            Wend
        End If
        
    'BOMList.Show
End Sub


Public Sub EmailErrorMessage(ErrorMessage As String)
    Dim EmailSession As OSSMTP.SMTPSession
    Dim env As String
    env = "COMPUTERNAME"
    computer = Environ(env)
    Set EmailSession = New OSSMTP.SMTPSession

    EmailSession.MessageSubject = "ScrapScan Error: " + computer
    EmailSession.MessageText = ErrorMessage
    EmailSession.AuthenticationType = AuthNone
    EmailSession.Server = "10.1.2.13"
    EmailSession.SendTo = "wrex@busche-cnc.com"
    EmailSession.MailFrom = computer + "@busche-cnc.com"
    EmailSession.SendEmail
End Sub


Function GetPartNumber(jobnumber As String) As String
    strFunction = "GetPartNumber"
    If M2MPart.State = adStateOpen Then
        M2MPart.Close
    End If
    M2MPart.Open "SELECT JOMAST.FPARTNO FROM JOMAST WHERE JOMAST.FJOBNO = '" + jobnumber + "'", M2MConn, adOpenKeyset, adLockReadOnly
    If M2MPart.RecordCount <> 0 Then
        M2MPart.MoveFirst
        GetPartNumber = Trim(M2MPart.Fields(0))
    Else
        ErrorMessage = "In Module: " + "SQL_Calls" + Chr(13) & Chr(10) + "Function: " + strFunction + Chr(13) & Chr(10) + ReportToMessage
        EmailErrorMessage (ErrorMessage)
        GetPartNumber = "Error"
    End If
    
End Function
Function GetPartNumberRev(jobnumber As String) As String
    strFunction = "GetPartNumberRev"
    If M2MPart.State = adStateOpen Then
        M2MPart.Close
    End If
    M2MPart.Open "SELECT JOMAST.FPARTREV FROM JOMAST WHERE JOMAST.FJOBNO = '" + jobnumber + "'", M2MConn, adOpenKeyset, adLockReadOnly
    If M2MPart.RecordCount <> 0 Then
        M2MPart.MoveFirst
        GetPartNumberRev = Trim(M2MPart.Fields(0))
    Else
        ErrorMessage = "In Module: " + "SQL_Calls" + Chr(13) & Chr(10) + "Function: " + strFunction + Chr(13) & Chr(10) + ReportToMessage
        EmailErrorMessage (ErrorMessage)
        GetPartNumberRev = "Error"
    End If
    
End Function


Function WriteRecords()
   On Error GoTo WriteRecords_Error

    Set VFPConn = New ADODB.Connection
    VFPConn.Open "Provider=VFPOLEDB;" + _
            "DATA SOURCE=" + App.Path + ";"
    Set VFRS = New ADODB.Recordset
    VFRS.ActiveConnection = VFPConn
    VFRS.Source = "bcctemp"
    VFRS.CursorType = adOpenKeyset
    VFRS.LockType = adLockOptimistic
    VFRS.CursorLocation = adUseClient
    VFRS.Open , , , , adCmdTable
    Dim strScrapCode As String
    Dim i As Integer
    Dim strPartNo As String
    Dim strPartRev As String

    For i = Menu.ListView1.ListItems.Count To 1 Step -1
        global_jobnumber = Menu.ListView1.ListItems.Item(i).SubItems(1)
        strScrapCode = Trim(Menu.ListView1.ListItems.Item(i).SubItems(3))
'*************************************
    If Left(strScrapCode, 1) = "3" Then
            strPartNo = Menu.ListView1.ListItems.Item(i).SubItems(6)
        Else
            strPartNo = ""
        End If
            
    If Left(strScrapCode, 1) = "3" Then
            strPartRev = Menu.ListView1.ListItems.Item(i).SubItems(7)
        Else
            strPartRev = ""
        End If
    
    If strPartNo <> "" Then
        Dim strPartId As String
        strPartId = GetPartRevIdentityColumn(strPartNo, strPartRev)
        lPN = Len(strPartId)
        For X = 1 To 10 - lPN
            strPartId = " " + strPartId
        Next X
        strPartId = "Y" + strPartId

        End If
            
    
    '******************


        If UCase(Left(Trim(Menu.ListView1.ListItems.Item(i).SubItems(5)), 1)) = "H" Then
                If Len(Menu.ListView1.ListItems.Item(i).SubItems(8)) > 0 Then
                        Call WriteF11Move(i, strScrapCode, True)
                        Menu.ListView1.ListItems.Remove (i)
                    Else
                        MsgBox ("Unable to save HOLD record without Location being assigned.")
                        Exit Function
                    End If
                        
            Else
                VFRS.AddNew
                VFRS.Fields(0) = "1"
                VFRS.Fields(1) = "104"
                VFRS.Fields(2) = ""
                VFRS.Fields(3) = False
                VFRS.Fields(4) = "F25"      '*****   Value Used to be F18, Before M2M Changed it.
                VFRS.Fields(5) = Menu.ListView1.ListItems.Item(i).Text
                VFRS.Fields(6) = 0
                VFRS.Fields(7) = Menu.ListView1.ListItems.Item(i).SubItems(3)
                VFRS.Fields(8) = ""
                VFRS.Fields(9) = "C" + Menu.ListView1.ListItems.Item(i).SubItems(1)
                VFRS.Fields(10) = "." + Menu.ListView1.ListItems.Item(i).SubItems(2)
                VFRS.Fields(11) = ""
                VFRS.Fields(12) = ""
                VFRS.Fields(13) = ""
                VFRS.Fields(14) = ""
                VFRS.Fields(15) = ""
                VFRS.Fields(16) = Val(Trim(Menu.ListView1.ListItems.Item(i).SubItems(4)))
                VFRS.Fields(17) = 0
                VFRS.Fields(18) = Menu.ListView1.ListItems.Item(i).SubItems(5)
                VFRS.Fields(19) = strPartId
                VFRS.Fields(20) = ""
                VFRS.Fields(21) = Format(Date, "yyyymmdd")
                VFRS.Fields(22) = Format(clocktime, "hh:mm")
                VFRS.Fields(23) = ""
                VFRS.Fields(24) = ""
                VFRS.Fields(25) = 0
                VFRS.Fields(26) = 0
                VFRS.Fields(27) = ""
                VFRS.Fields(28) = ""
                VFRS.Fields(29) = 0
                VFRS.Fields(30) = ""
                VFRS.Fields(31) = ""
                VFRS.Fields(32) = 0
                VFRS.Update
                
                If (strScrapCode <> "000") Then
                    Call WriteF11Move(i, strScrapCode, False)
                End If
                Menu.ListView1.ListItems.Remove (i)
        End If
        
    Next
    
    If Not (VFRS Is Nothing) Then
       If VFRS.State = adStateOpen Then
         VFRS.Close
        End If
        Set VFRS = Nothing
    End If
    If Not (VFPConn Is Nothing) Then
       If VFPConn.State = adStateOpen Then
          VFPConn.Close
       End If
       Set VFPConn = Nothing
    End If

   On Error GoTo 0
   Exit Function

WriteRecords_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteRecords of Module SQL_Calls"
  
End Function

Sub WriteF11Move(ByVal ii As Integer, strScrapCode As String, HoldFlag As Boolean)
' Database Connection must already be open
    Dim strJobNo As String
    Dim strPartNo As String
    Dim strPartRev As String
    Dim iCompQty As Integer
     
 '   strJobNo = Menu.ListView1.ListItems.Item(i).SubItems(1)
   On Error GoTo WriteF11Move_Error

    strJobNo = global_jobnumber
    If Left(strScrapCode, 1) = "3" Then
            strPartNo = Menu.ListView1.ListItems.Item(ii).SubItems(6)
        Else
            strPartNo = GetPartNumber(strJobNo)
        End If
        
 '   strPartNo = gPartNumber ' = List1.SelectedItem.Text
    If Left(strScrapCode, 1) = "3" Then
            strPartRev = Menu.ListView1.ListItems.Item(ii).SubItems(7)
        Else
            strPartRev = GetPartNumberRev(strJobNo)
        End If
     ' = List1.SelectedItem.ListSubItems.Item(1).Text
    Dim strPartId As String
    strPartId = GetPartRevIdentityColumn(strPartNo, strPartRev)
    iCompQty = Val(Trim(Menu.ListView1.ListItems.Item(ii).SubItems(4)))

    Dim strBinLocationId As String
    If (HoldFlag) And (Len(Menu.ListView1.ListItems.Item(ii).SubItems(9)) > 0) Then
            strBinLocationId = GetLocationIDCol(Menu.ListView1.ListItems.Item(ii).SubItems(9))
        Else
'            strScrapBinLocationId = GetScrapBinLocationIdentityColumn(strScrapCode)
            strBinLocationId = GetBinLocationIdentityColumn(strJobNo, strScrapCode, strPartNo, strPartRev)
        End If
    lId = Len(strBinLocationId)
    For i = 1 To 11 - lId
        strBinLocationId = " " + strBinLocationId
    Next i
    Dim strScrapBinLocationId As String
    If HoldFlag Then
            strScrapBinLocationId = GetLocationIDCol(Menu.ListView1.ListItems.Item(ii).SubItems(8))
        Else
            strScrapBinLocationId = GetScrapBinLocationIdentityColumn(strScrapCode)
        End If
            
    lId = Len(strScrapBinLocationId)
    For j = 1 To 11 - lId
        strScrapBinLocationId = " " + strScrapBinLocationId
    Next j
    
    VFRS.AddNew
'The following is specific information that needs to be entered for the "F9" Move to finished goods
'transaction
    VFRS.Fields(4) = "F11"                          'Fnction - Function Code
    VFRS.Fields(11) = strBinLocationId           'Fpro_id - From Location - BCLOCBIN.identity_column - ^^^^^^^^132 (11 digits)
    VFRS.Fields(16) = iCompQty                   'Fcompqty - Quantity - 124.0
    lPN = Len(strPartId)
    For X = 1 To 10 - lPN
        strPartId = " " + strPartId
    Next X
    strPartId = "Y" + strPartId
    VFRS.Fields(18) = strPartId                  'Fpartno  Y^^^^^^^132 (11 digits)
    VFRS.Fields(19) = strScrapBinLocationId      'Ftojob  - To Location - BCLOCBIN.identity_column - ^^^^^^^^132 (11 digits)
    VFRS.Fields(21) = Format(Date, "yyyymmdd")   'Fdate - 20020114
    VFRS.Fields(22) = Format(clocktime, "hh:mm") 'Ftime - 10:19

'The following is stub data that is generic for all "F9" transactions.
    VFRS.Fields(0) = "1"    'Fbcnum
    VFRS.Fields(1) = "104"  'Fnetaddr
    VFRS.Fields(2) = ""     'Frecstat
    VFRS.Fields(3) = 0      'Fpost_strt
    VFRS.Fields(5) = ""     'Fempno
    VFRS.Fields(7) = ""     'Fjobno - Job Number - I3634-0000
    VFRS.Fields(6) = False  'Ferrbc
    VFRS.Fields(8) = ""     'Foperno
    VFRS.Fields(9) = ""     'Fnjobno
    VFRS.Fields(10) = ""    'Fnoperno
    VFRS.Fields(12) = "Y"   'Flead
    VFRS.Fields(13) = "N"   'Fsetup
    VFRS.Fields(14) = "N"   'Frework
    VFRS.Fields(15) = "N"   'Fcmpl
    VFRS.Fields(17) = 0     'Fscrpqty
    VFRS.Fields(20) = ""    'Fseriesend
    VFRS.Fields(23) = ""    'Forg_date
    VFRS.Fields(24) = ""    'Forg_time
    VFRS.Fields(25) = 0     'Fshftdt
    VFRS.Fields(26) = 0     'Fshfttm
    VFRS.Fields(27) = ""    'Fsub
    VFRS.Fields(28) = ""    'Fship
    VFRS.Fields(29) = 0     'Flast
    VFRS.Fields(30) = ""    'Fclockout
    VFRS.Fields(31) = ""    'Fclot
    VFRS.Fields(32) = 0     'Fdlotexp
    VFRS.Update
    

   On Error GoTo 0
   Exit Sub

WriteF11Move_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteF11Move of Module SQL_Calls"

End Sub


Function GetBinLocationIdentityColumn(jobnumber As String, strScrapCode As String, _
                                        strPartNo As String, strPartRev As String) As String
    strFunction = "GetBinLocationIdentityColumn"
    If M2MPart.State = adStateOpen Then
        M2MPart.Close
    End If
    
    If Left(strScrapCode, 1) = "3" Then
            M2MPart.Open "select BCLOCBIN.identity_column " & _
                "From INMAST inner join " & _
                "BCLOCBIN on INMAST.FLOCATE1 = BCLOCBIN.FCLOCATION AND INMAST.FBIN1 = BCLOCBIN.FCBINNO " & _
                "where (INMAST.FPARTNO = '" + strPartNo + "' and INMAST.FREV = '" + strPartRev + "') ", _
                M2MConn, adOpenKeyset, adLockReadOnly
        Else
            M2MPart.Open "select BCLOCBIN.identity_column " & _
                "From JOMAST inner join " & _
                "INMAST on JOMAST.FPARTNO = INMAST.FPARTNO and jomast.fpartrev = inmast.frev inner join " & _
                "BCLOCBIN on INMAST.FLOCATE1 = BCLOCBIN.FCLOCATION AND INMAST.FBIN1 = BCLOCBIN.FCBINNO " & _
                "where (JOMAST.FJOBNO = '" + jobnumber + "')", M2MConn, adOpenKeyset, adLockReadOnly
        End If
        
    If M2MPart.RecordCount <> 0 Then
        M2MPart.MoveFirst
        GetBinLocationIdentityColumn = Trim(M2MPart.Fields(0))
    Else
        ErrorMessage = "In Module: " + "SQL_Calls" + Chr(13) & Chr(10) + "Function: " + strFunction + Chr(13) & Chr(10) + ReportToMessage
        EmailErrorMessage (ErrorMessage)
        GetBinLocationIdentityColumn = "Error"
        
    End If

End Function


Function GetPartRevIdentityColumn(strPartNumber As String, strPartRev As String) As String
    strFunction = "GetPartRevIdentityColumn"
    If M2MPart.State = adStateOpen Then
        M2MPart.Close
    End If
    'select inmast.fpartno, identity_column, *
    '    From INMAST where (fpartno = '177899')
    
    M2MPart.Open "select identity_column " & _
        "From INMAST " & _
        "where (fpartno = '" + strPartNumber + "') and " & _
        "      (frev = '" + strPartRev + "')", M2MConn, adOpenKeyset, adLockReadOnly
    
    If M2MPart.RecordCount <> 0 Then
        M2MPart.MoveFirst
        GetPartRevIdentityColumn = Trim(M2MPart.Fields(0))
    Else
        ErrorMessage = "In Module: " + "SQL_Calls" + Chr(13) & Chr(10) + "Function: " + strFunction + Chr(13) & Chr(10) + ReportToMessage
        EmailErrorMessage (ErrorMessage)
        GetPartRevIdentityColumn = "Error"
    End If


End Function




Function GetScrapBinLocationIdentityColumn(strScrapCode As String) As String
    Dim strScrapBinLocation As String
    strFunction = "GetScrapBinLocationIdentityColumn"
    Select Case Left(strScrapCode, 1)
    
        Case "2"
           If (MDIForm1.P2.Checked) Then strScrapBinLocation = "222-2"
           If (MDIForm1.P3.Checked) Then strScrapBinLocation = "333-2"
           If (MDIForm1.P5.Checked) Then strScrapBinLocation = "555-2"
           If (MDIForm1.P6.Checked) Then strScrapBinLocation = "666-2"
           If (MDIForm1.P7.Checked) Then strScrapBinLocation = "777-2"
           If (MDIForm1.P8.Checked) Then strScrapBinLocation = "999-2"
           If (MDIForm1.P9.Checked) Then strScrapBinLocation = "990-2"
           If (MDIForm1.P11.Checked) Then strScrapBinLocation = "110-2"
           If (MDIForm1.P8A.Checked) Then strScrapBinLocation = "888-2"
        
        Case "3"
           If (MDIForm1.P2.Checked) Then strScrapBinLocation = "222-3"
           If (MDIForm1.P3.Checked) Then strScrapBinLocation = "333-3"
           If (MDIForm1.P5.Checked) Then strScrapBinLocation = "555-3"
           If (MDIForm1.P6.Checked) Then strScrapBinLocation = "666-3"
           If (MDIForm1.P7.Checked) Then strScrapBinLocation = "777-3"
           If (MDIForm1.P8.Checked) Then strScrapBinLocation = "999-3"
           If (MDIForm1.P9.Checked) Then strScrapBinLocation = "990-3"
           If (MDIForm1.P11.Checked) Then strScrapBinLocation = "110-3"
           If (MDIForm1.P8A.Checked) Then strScrapBinLocation = "888-3"
        
        Case Else
        '  If (MDIForm1.P2.Checked) Then GetScrapBinLocationIdentityColumn = "12345678901"
           If (MDIForm1.P2.Checked) Then strScrapBinLocation = "222"
           If (MDIForm1.P3.Checked) Then strScrapBinLocation = "333"
           If (MDIForm1.P5.Checked) Then strScrapBinLocation = "555"
           If (MDIForm1.P6.Checked) Then strScrapBinLocation = "666"
           If (MDIForm1.P7.Checked) Then strScrapBinLocation = "777"
           If (MDIForm1.P8.Checked) Then strScrapBinLocation = "999"
           If (MDIForm1.P9.Checked) Then strScrapBinLocation = "990"
           If (MDIForm1.P11.Checked) Then strScrapBinLocation = "110-1"
           If (MDIForm1.P8A.Checked) Then strScrapBinLocation = "888"
       End Select
    
    GetScrapBinLocationIdentityColumn = GetLocationIDCol(strScrapBinLocation)
    
    End Function

Public Function GetLocationIDCol(LocCode As String)

    If M2MPart.State = adStateOpen Then
        M2MPart.Close
    End If
    
    M2MPart.Open "select identity_column " & _
        "From BCLOCBIN WHERE fclocation = '" + LocCode + "'", M2MConn, adOpenKeyset, adLockReadOnly
 
    If M2MPart.RecordCount <> 0 Then
        M2MPart.MoveFirst
        GetLocationIDCol = Trim(M2MPart.Fields(0))
    Else
        ErrorMessage = "In Module: " + "SQL_Calls" + Chr(13) & Chr(10) + "Function: " + strFunction + Chr(13) & Chr(10) + ReportToMessage
        EmailErrorMessage (ErrorMessage)
        GetLocationIDCol = "Error"
    End If

End Function

Public Sub closedb()
    M2MConn.Close
    Set M2MEmp = Nothing
    Set M2MJobs = Nothing
    Set M2MDesc = Nothing
    Set M2MJobs = Nothing
    Set M2MPart = Nothing
    Set M2MEff = Nothing
    Set M2MGMsg = Nothing
    Set M2MMsg = Nothing

End Sub

Public Sub InsertRecords()
   On Error GoTo InsertRecords_Error

    Set VFPConn = New ADODB.Connection
    VFPConn.Open "Provider=VFPOLEDB;" + _
            "DATA SOURCE=" + App.Path + ";"
    Set VFRS = New ADODB.Recordset
    VFRS.ActiveConnection = VFPConn
    VFRS.Source = "bcctemp"
    VFRS.CursorType = adOpenKeyset
    VFRS.LockType = adLockOptimistic
    VFRS.CursorLocation = adUseClient
    VFRS.Open , , , , adCmdTable
    If VFRS.RecordCount > 0 Then
        Set VFPConn2 = New ADODB.Connection
        VFPConn2.Open "Provider=VFPOLEDB;" + _
            "DATA SOURCE=" + VFPTABLE + ";"
        Set VFRS2 = New ADODB.Recordset
        VFRS2.ActiveConnection = VFPConn2
        VFRS2.Source = "bcshared"
        VFRS2.CursorType = adOpenKeyset
        VFRS2.LockType = adLockOptimistic
        VFRS2.CursorLocation = adUseClient
        VFRS2.Open , , , , adCmdTable
        VFRS.MoveFirst
        For i = 0 To VFRS.RecordCount - 1
            VFRS2.AddNew
            For j = 0 To 32
                VFRS2.Fields(j) = VFRS.Fields(j)
            Next
            VFRS.MoveNext
            VFRS2.Update
        Next
        Dim fs
        Set fs = CreateObject("scripting.filesystemobject")
        fs.CopyFile App.Path + "\BCMaster\bcmast.dbf", App.Path + "\bcctemp.dbf"
        MDIForm1.StatusBar1.Panels(3).Text = "RECORDS TRANSFERRED"
    End If
If VFRS.State = adStateOpen Then
    VFRS.Close
End If
If VFRS2.State = adStateOpen Then
    VFRS2.Close
End If
If VFPConn.State = adStateOpen Then
    VFPConn.Close
End If
While VFPConn2.State = adStateOpen
    VFPConn2.Close
Wend
Set VFPConn2 = Nothing
Set VFRS = Nothing
Set VFPConn = Nothing
Set VFRS2 = Nothing

   On Error GoTo 0
   Exit Sub

InsertRecords_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InsertRecords of Module SQL_Calls"
End Sub
