VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Principal"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   7215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8070
      _Version        =   393216
   End
   Begin VB.CommandButton Importar 
      Caption         =   "Importar"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   600
      Width           =   1635
   End
   Begin VB.CommandButton Exportar 
      Caption         =   "Exportar"
      Height          =   500
      Left            =   9240
      TabIndex        =   2
      Top             =   5940
      Width           =   2000
   End
   Begin VB.CommandButton Documentacao 
      Caption         =   "Documentação"
      Height          =   500
      Left            =   120
      TabIndex        =   1
      Top             =   5940
      Width           =   2000
   End
   Begin VB.CommandButton DownloadLayout 
      Caption         =   "Download Layout"
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Caminho:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Desafio Paschoalotto - Roberto de Jesus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   11415
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoConectBank As New ADODB.Connection
Dim ConectBank As String

Private Sub Command1_Click()
    txtPath.Text = ""
End Sub

Private Sub Documentacao_Click()
    Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://github.com/0roberto0/"
End Sub

Private Sub Exportar_Click()
    Dim xl As New excel.Application
    Dim xlwb As excel.Workbook
    Dim xlws As excel.Worksheet
    Dim Folder As Object
    Dim FolderFullPath, newCell, dateNow As String
    Dim objshell As New Shell32.Shell
    Dim i, p As Integer

    Set Folder = objshell.BrowseForFolder(Me.hWnd, "Selecione a pasta", _
    BIF_RETURNONLYFSDIRS)
    If Not (Folder Is Nothing) Then
        FolderFullPath = Folder.Items.Item.Path
        
        Set xl = New excel.Application
        xl.Visible = False
        
        Set xlwb = xl.Workbooks.Add
        Set xlws = xlwb.Worksheets(1)
        
        For i = 1 To Grid.Rows - 1
            For p = 2 To Grid.Cols - 1
                Grid.col = p
                Grid.Row = i
                newCell = Chr(p - 1 + 64) & i
                xl.Worksheets(1).Range(newCell).Value = Grid.Text
            Next
        Next
            
        dateNow = Format(Now, "ddMMyyyy_hhmmss")
        xlws.SaveAs FolderFullPath + "\" + "Export" + "_" + dateNow + ".xlsx"
        MsgBox "Arquivo exportado com sucesso!"
        
        xlwb.Close
        xl.Visible = True
        xl.Workbooks.Close
        xl.Quit
        reload_Grid
    Else
       MsgBox "Diretório não selecionado!"
    End If
End Sub

Private Function reload_Grid()
    Dim rs As New ADODB.Recordset
    
    ConectBank = "DSN=PostgreSQL30;Database=postgres;Server=localhost;Uid=postgres;Port=5432;pwd=1234"
    AdoConectBank.Open ConectBank
    
    rs.Open "select * from pokedex", AdoConectBank, adOpenKeyset, adLockOptimistic
    FillGrid Grid, rs

    rs.Close
    AdoConectBank.Close
End Function

Private Sub Form_Load()
    reload_Grid
End Sub

Private Sub DownloadLayout_Click()
    Dim xl As New excel.Application
    Dim xlwb As excel.Workbook
    Dim xlws As excel.Worksheet
    Dim objshell As New Shell32.Shell
    Dim Folder As Object
    Dim FolderFullPath As String
     
    Set Folder = objshell.BrowseForFolder(Me.hWnd, "Select a Folder", _
    BIF_RETURNONLYFSDIRS)
    If Not (Folder Is Nothing) Then
        FolderFullPath = Folder.Items.Item.Path
        
        Set xl = New excel.Application
        xl.Visible = False
        
        Set xlwb = xl.Workbooks.Add
        Set xlws = xlwb.Worksheets(1)
        
        xlws.Cells(1, 1).Value = "Name"
        xlws.Cells(1, 2).Value = "Type 1"
        xlws.Cells(1, 3).Value = "Type 2"
        xlws.Cells(1, 4).Value = "Total"
        xlws.Cells(1, 5).Value = "HP"
        xlws.Cells(1, 6).Value = "Attack"
        xlws.Cells(1, 7).Value = "Defense"
        
        xlws.SaveAs FolderFullPath + "\" + "teste2.xls"
        MsgBox "Arquivo salvo com sucesso!"
        
        xlwb.Close
        xl.Visible = True
        xl.Workbooks.Close
        xl.Quit
    Else
       MsgBox "Diretório não selecionado!"
    End If
End Sub

Private Sub Importar_Click()
    If txtPath.Text <> "" Then
        ImportCSV
    Else
        ImportArchive.Show vbModal
        If Not txtPath.Text = "" Then
            Importar_Click
        End If
    End If
End Sub

Private Function ImportCSV()
    Dim xl As New excel.Application
    Dim xlwb As excel.Workbook
    Dim xlws As excel.Worksheet
    Dim col As Range
    Dim Interval As Range
    Dim i, j, lines, columns As Integer
    Dim data(1 To 7) As String
    Dim clocal, newCell As String

    i = 1
    Set xl = New excel.Application
    xl.Visible = False
    Set xlwb = xl.Workbooks.Open(txtPath.Text)
    Set xlws = xlwb.Worksheets(1)
    lines = xlws.Range("A2").End(xlDown).Row
    columns = xlws.Range("A1").End(xlToRight).Column
    For i = 2 To lines
        For j = 1 To columns
            newCell = Chr(j + 64) & i
            data(j) = xlws.Range(newCell).Value
        Next
        InsertValues data(1), data(2), data(3), CInt(data(4)), CInt(data(5)), CInt(data(6)), CInt(data(7))
    Next
    MsgBox "Importação realizada com sucesso!"

    xlwb.Close
    xl.Visible = True
    xl.Workbooks.Close
    xl.Quit
    reload_Grid
End Function

Private Function InsertValues(Nome As String, Type_1 As String, Type_2 As String, Total As Integer, HP As Integer, Attack As Integer, Defense As Integer)
On Error GoTo Err
    Dim adoConn As New ADODB.Connection
    Dim statement As String
    
    ConectBank = "DSN=PostgreSQL30;Database=postgres;Server=localhost;Uid=postgres;Port=5432;pwd=1234"
    adoConn.Open ConectBank
    
    statement = "INSERT INTO public.pokedex(name_pokemon, Type_1, Type_2, Total, HP, Attack, Defense) VALUES (" & "'" & Nome & "', " & "'" & Type_1 & "', " & "'" & Type_2 & "', " & "'" & Total & "', " & "'" & HP & "', " & "'" & Attack & "', " & "'" & Defense & "')"
    adoConn.Execute statement, , adCmdText
    adoConn.Close
    Exit Function
Err:
    MsgBox Err.Description
End Function

Public Function FillGrid(FlexGrid As Object, rs As Object)
On Error GoTo Err

    If Not TypeOf FlexGrid Is MSFlexGrid Then Exit Function
    If Not TypeOf rs Is ADODB.Recordset Then Exit Function
    
    Dim i As Integer
    Dim j As Integer

    FlexGrid.FixedRows = 1
    FlexGrid.FixedCols = 0

    If Not rs.EOF Then
        FlexGrid.Rows = rs.RecordCount + 1
        FlexGrid.Cols = rs.Fields.Count

        For i = 0 To rs.Fields.Count - 1
            FlexGrid.TextMatrix(0, i) = rs.Fields(i).Name
        Next
        i = 1
        
        Do While Not rs.EOF
            For j = 0 To rs.Fields.Count - 1
                If Not IsNull(rs.Fields(j).Value) Then
                    FlexGrid.TextMatrix(i, j) = rs.Fields(j).Value
                End If
            Next
        i = i + 1
        rs.MoveNext
    Loop
    End If
    FillGrid = True
Err:
   FillGrid = False
   Exit Function
End Function
