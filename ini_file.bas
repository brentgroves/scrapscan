Attribute VB_Name = "ini_file"
' Retrieves info from the settings.ini file to initialize global variables
' TODO add a plant# ini file option.  Currently the plant# is hardcoded to 3.
Public Sub getSettings()
    Dim inifile As String
    Dim fs, a, unknown
    inifile = Dir(App.Path + "\settings.ini")
    If Len(inifile) > 0 Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.OpenTextFile(App.Path + "\settings.ini")
        Settings.M2MDB.Text = a.ReadLine
        Settings.M2MSERVER_TEXT.Text = a.ReadLine
        Settings.M2MUSER.Text = a.ReadLine
        Settings.M2MPASS.Text = a.ReadLine
'        Settings.VFPTABLE.Text = "T:\"
        Settings.VFPTABLE.Text = a.ReadLine
        Settings.txtListType.Text = a.ReadLine
'       uncomment for plant revision addition
'       Settings.PLANT.Text = a.ReadLine
        
        
        VFPTABLE = Trim(Settings.VFPTABLE.Text)
        M2MDB = Trim(Settings.M2MDB.Text)
        M2MSERVER = Trim(Settings.M2MSERVER_TEXT.Text)
        M2MUSER = Trim(Settings.M2MUSER.Text)
        M2MPASS = Trim(Settings.M2MPASS.Text)
        LISTTYPE = Trim(Settings.txtListType.Text)
'       uncomment for plant revision addition
'       PLANT = Trim(Settings.PLANT.Text)
        a.Close
    Else
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(App.Path + "\settings.ini", True)
        a.WriteLine ("M2MDATA01")
        a.WriteLine ("MOBIL\AARON")
        a.WriteLine ("sa")
        a.WriteLine ("buschecnc1")
        a.WriteLine ("c:\")
        a.WriteLine ("PartList")
'       uncomment for plant revision addition
'       a.WriteLine ("10")
        a.Close
    End If
End Sub

Public Sub writesettings()
    Dim inifile As String
    Dim fs, a, t
    Set fs = CreateObject("scripting.filesystemobject")
    fs.DeleteFile (App.Path + "\settings.ini")
    Set a = fs.CreateTextFile(App.Path + "\settings.ini", True)
    a.WriteLine Trim(Settings.M2MDB.Text)
    a.WriteLine Trim(Settings.M2MSERVER_TEXT.Text)
    a.WriteLine Trim(Settings.M2MUSER.Text)
    a.WriteLine Trim(Settings.M2MPASS.Text)
    a.WriteLine Trim(Settings.VFPTABLE.Text)
    a.WriteLine Trim(Settings.txtListType.Text)
'   uncomment for plant revision addition
'   a.WriteLine Trim(Settings.PLANT.Text)
    a.Close
    getSettings
End Sub

' Retrieves info from the Locations.ini file to initialize Locations List

Public Sub getLocationSettings()
    Dim inifile As String
    Dim fs, a, unknown
    Dim itmx As ListItem
    
    inifile = Dir(App.Path + "\Locations.ini")
    If Len(inifile) > 0 Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.OpenTextFile(App.Path + "\Locations.ini")
       
        Locations.lstSelected.ListItems.Clear
        
        Do Until a.AtEndOfStream
            Code = a.ReadLine
            Set itmx = Locations.lstSelected.ListItems.Add(, , Code)
            Desc = a.ReadLine
            itmx.SubItems(1) = Desc
            Loop
            
        a.Close
    Else
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(App.Path + "\Locations.ini", True)
        a.WriteLine ("")
        a.Close
    End If
    
    
End Sub

Public Sub writeLocationSettings()
    Dim inifile As String
    Dim fs, a, t
    Dim itmx As ListItem
    
    Set fs = CreateObject("scripting.filesystemobject")
    fs.DeleteFile (App.Path + "\Locations.ini")
    Set a = fs.CreateTextFile(App.Path + "\Locations.ini", True)
    
    For i = 1 To Locations.lstSelected.ListItems.Count
        a.WriteLine Trim(Locations.lstSelected.ListItems(i).Text)
        a.WriteLine Trim(Locations.lstSelected.ListItems(i).ListSubItems(1).Text)
        Next i

    a.Close
End Sub


Public Sub GetBCMast()
        Set VFPConn = New ADODB.Connection
        VFPConn.Open "Driver=Microsoft Visual Foxpro Driver; " + _
            "UID=;SourceType=DBf;sourceDB=" + App.Path + "; Exclusive=No;"
        Set VFRS = New ADODB.Recordset
        VFRS.ActiveConnection = VFPConn
        VFRS.Source = "bcctemp.dbf"
        VFRS.CursorType = adOpenKeyset
        VFRS.LockType = adLockOptimistic
        VFRS.CursorLocation = adUseClient
        VFRS.Open , , , , adCmdTable
        If VFRS.RecordCount > 0 Then
            InsertRecords
        End If
End Sub

