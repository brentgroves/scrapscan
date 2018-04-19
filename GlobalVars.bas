Attribute VB_Name = "GlobalVars"
'*************GENERAL VARIABLES***********
Public M2MConn As ADODB.Connection
Public VFPConn As ADODB.Connection
Public VFRS As ADODB.Recordset
Public VFPConn2 As ADODB.Connection
Public VFRS2 As ADODB.Recordset
Public M2MEmp As ADODB.Recordset
Public M2MJobs As ADODB.Recordset
Public M2MDesc As ADODB.Recordset
Public M2MPart As ADODB.Recordset
Public fnction As String
Public empName As String
Public global_jobnumber As String
Public opNumber As String
Public empNumber As String
Public clocktime As String
Public Piece_Count As Integer
Public PLANT As String
Public M2MDB As String
Public VFPTABLE As String
Public M2MSERVER As String
Public M2MUSER As String
Public M2MPASS As String
Public LISTTYPE As String
Public JobGrp As String
Public gPartNumber As String
Public gPartRev As String
Public LastActivity As Date
Public RefreshDB As Date
Public SCRAPCODE As String
Public TagID As String
Public ErrorMessage As String
Public ReportToMessage As String
Public PartRev As String
Public M2MLocations As ADODB.Recordset


