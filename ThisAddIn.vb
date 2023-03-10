Imports System.ComponentModel
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Word

Public Class ThisAddIn
    Public Shared AddinCustomTaskPanes As New Dictionary(
        Of Document, Microsoft.Office.Tools.CustomTaskPane)


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            Dim NewCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane
            NewCustomTaskPane = Me.CustomTaskPanes.Add(New MyPanel, "VPCC LHA")
            NewCustomTaskPane.Visible = False

            CustomTaskPanes.Remove(NewCustomTaskPane)

        Catch ex As Exception
        End Try


    End Sub
    Private Sub Application_DocumentBeforeClose(ByVal Doc As Document, ByRef Cancel As Boolean) Handles Application.DocumentBeforeClose

        DeleteCustomTaskPane(Doc)
    End Sub

    Private Sub DeleteCustomTaskPane(ByVal Doc As Document)
        Try
            Me.CustomTaskPanes.Remove(AddinCustomTaskPanes.Item(Doc))
            AddinCustomTaskPanes.Remove(Doc)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        'https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2022
        'https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce?view=vs-2022&tabs=vb
    End Sub

End Class

Module SharedFunction
    Public Function NumtoStr(strIn As String) As String

        Dim sc() As String = {"không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín", "mười", "lăm"}
        Dim a As Integer, b As Integer, C As Integer, k As Integer
        Dim s1 As String, s2 As String, s3 As String, mns As String

        Dim iResult As String = ""

        Dim instrNum As Double = Double.Parse(strIn)
        Dim strInFormated As String

        a = Len(strIn)
        If (a Mod 3) <> 0 Then
            C = (a \ 3 + 1) * 3
            strInFormated = instrNum.ToString(StrDup(C, "0"))
        Else
            strInFormated = instrNum.ToString(StrDup(a, "0"))
        End If

        C = Len(strInFormated) / 3
        k = 0

        For i = C To 1 Step -1

            b = i * 3 - 2
            k += 1

            mns = Mid(strInFormated, b, 3)
            s1 = Mid(mns, 1, 1) : s2 = Mid(mns, 2, 1) : s3 = Mid(mns, 3, 1)

            Select Case k
                Case 1 : If (i <> C) Then iResult = "tỷ " & Trim(iResult)
                Case 2 : If mns <> "000" Then iResult = "nghìn " & Trim(iResult)
                Case 3 : If mns <> "000" Then iResult = "triệu " & Trim(iResult)
            End Select

            Select Case Integer.Parse(s3)
                Case 0
                Case 5
                    Select Case Integer.Parse(s2)
                        Case 0 : iResult = "năm " & Trim(iResult)
                        Case Else : iResult = "lăm " & Trim(iResult)
                    End Select
                Case 1
                    Select Case Integer.Parse(s2)
                        Case 0, 1 : iResult = "một " & Trim(iResult)
                        Case Else : iResult = "mốt " & Trim(iResult)
                    End Select
                Case Else
                    iResult = sc(s3) & Space(1) & Trim(iResult)
            End Select

            Select Case Integer.Parse(s2)
                Case 0
                    Select Case Integer.Parse(s3)
                        Case Is <> 0
                            Select Case i
                                Case 1 : Select Case Integer.Parse(s1) : Case 0 : Case Else : iResult = "lẻ " & Trim(iResult) : End Select
                                Case Else : iResult = "lẻ " & Trim(iResult)
                            End Select
                    End Select
                Case 1 : iResult = "mười " & Trim(iResult)
                Case Else : iResult = sc(s2) & " mươi " & Trim(iResult)
            End Select

            Select Case Integer.Parse(s1)
                Case 0
                    If (Mid(strInFormated, 1, b)) <> 0 And (s2 <> 0 Or s3 <> 0) Then iResult = "không trăm " & Trim(iResult)
                Case Else
                    iResult = sc(s1) & " trăm " & Trim(iResult)
            End Select

            If k = 3 Then k = 0

        Next i

        Return iResult
    End Function
    Public Function RemoveSpace(strIn As String) As String
        RemoveSpace = Regex.Replace(strIn, "\s+", "")
    End Function
    Public Function SearchDocForPattern(inPattern As String, Optional IgnoreCase As Boolean = True) As Integer
        Dim findcount As Integer = 0
        Dim oRx As New Regex(inPattern, RegexOptions.IgnoreCase = IgnoreCase)
        Dim ActiveDocument As Word.Document = Globals.ThisAddIn.Application.ActiveDocument

        Dim temp_para As Word.Paragraph
        For Each temp_para In ActiveDocument.Paragraphs
            If oRx.IsMatch(temp_para.Range.Text) Then
                temp_para.Range.Select()
                findcount += 1
            End If
        Next

        Return findcount
    End Function
    Public Function CalculateExactDateDifference(startDate As DateTime, endDate As DateTime) As String
        Dim years As Integer = endDate.Year - startDate.Year
        Dim months As Integer = endDate.Month - startDate.Month
        Dim days As Integer = endDate.Day - startDate.Day

        ' Adjust the year, month, and day values based on the actual number of days in each month and the leap year.
        If days < 0 Then
            days += DateTime.DaysInMonth(endDate.Year, endDate.Month)
            months -= 1
        End If
        If months < 0 Then
            months += 12
            years -= 1
        End If
        Return years & " năm " & months & " tháng " & days & " ngày"
        ' Display the results.
    End Function

End Module
Public Class NotaryCustomer
#Region "KhProperties"
    Private iID As String
    Public Property ID() As String
        Get
            Return iID
        End Get
        Set(ByVal value As String)
            iID = value
        End Set
    End Property
    Private iGt As String
    Public Property Gt() As String
        Get
            Return iGt
        End Get
        Set(ByVal value As String)
            iGt = value
        End Set
    End Property
    Private iTen As String
    Public Property Ten() As String
        Get
            Return iTen
        End Get
        Set(ByVal value As String)
            iTen = value
        End Set
    End Property
    Private iSn As String
    Public Property Sn() As String
        Get
            Return iSn
        End Get
        Set(ByVal value As String)
            iSn = value
        End Set
    End Property
    Private iCCCD As String
    Public Property CCCD() As String
        Get
            Return iCCCD
        End Get
        Set(ByVal value As String)
            iCCCD = value
        End Set
    End Property
    Private iCMND As String
    Public Property CMND() As String
        Get
            Return iCMND
        End Get
        Set(ByVal value As String)
            iCMND = value
        End Set
    End Property
    Private iHC As String
    Public Property HC() As String
        Get
            Return iHC
        End Get
        Set(ByVal value As String)
            iHC = value
        End Set
    End Property
    Private iCMSQ As String
    Public Property CMSQ() As String
        Get
            Return iCMSQ
        End Get
        Set(ByVal value As String)
            iCMSQ = value
        End Set
    End Property
    Private iSDDCN As String
    Public Property SDDCN() As String
        Get
            Return iSDDCN
        End Get
        Set(ByVal value As String)
            iSDDCN = value
        End Set
    End Property
    Private iTT As String
    Public Property TT() As String
        Get
            Return iTT
        End Get
        Set(ByVal value As String)
            iTT = value
        End Set
    End Property
    Private iLocation As String = "\\192.168.1.30\ho so chung\z.kh\KH.accdb"
    Public Property Location() As String
        Get
            Return iLocation
        End Get
        Set(ByVal value As String)
            iLocation = value
        End Set
    End Property
    Public ReadOnly Property CMT As String
        Get
            If Trim(CCCD) <> "" Then
                Return CCCD
            ElseIf Trim(CMND) <> "" Then
                Return CMND
            ElseIf Trim(HC) <> "" Then
                Return HC
            ElseIf Trim(CMSQ) <> "" Then
                Return CMSQ
            ElseIf Trim(SDDCN) <> "" Then
                Return SDDCN
            End If
        End Get
    End Property
    Private Function ReturnSizeOfProperty(input As String) As Integer
        If input = "" Then Return 1
        Return Len(input)
    End Function
#End Region
#Region "PrivateFunction"
    Public Enum ExportType
        Form
        OneLine
    End Enum

    Private Function AddSpaceCCCD(strIn As String) As String
        Dim iResult As String = ""
        If Regex.IsMatch(strIn, "\b\d{12}\b") Then
            For Each match In Regex.Matches(strIn, "(\d{3})")
                iResult &= match.ToString & " "
            Next
            Return iResult
        End If
        If Regex.IsMatch(strIn, "\b\d{9}\b") Then
            For Each match In Regex.Matches(strIn, "(\d{3})")
                iResult &= match.ToString & " "
            Next
            Return iResult
        End If
        Return strIn
    End Function
#End Region
    Public Sub GetInfoFromRS(ByVal rs As ADODB.Recordset)
        If rs.State = ADODB.ObjectStateEnum.adStateClosed Then Exit Sub
        If rs.EOF And rs.BOF Then Exit Sub

        ID = rs("ID").Value.ToString
        Gt = rs("Gt").Value.ToString
        Ten = rs("Ten").Value.ToString
        Sn = rs("Sn").Value.ToString
        CCCD = rs("CCCD").Value.ToString
        CMND = rs("CMND").Value.ToString
        HC = rs("HC").Value.ToString
        CMSQ = rs("CMSQ").Value.ToString
        SDDCN = rs("SDDCN").Value.ToString
        TT = rs("TT").Value.ToString
    End Sub
    Public Function IsLegitCustomer() As Boolean
        If Trim(Ten) = "" Then Return False
        If Trim(CCCD) = "" And Trim(CMND) = "" And Trim(HC) = "" And Trim(CMSQ) = "" And Trim(SDDCN) = "" Then
            Return False
        End If
        Return True
    End Function
    Public Function IsCustomerAvailabled() As Boolean
        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Location
        Dim connection As New ADODB.Connection
        connection.Open(connectionString)

        Dim sql As String = "SELECT * FROM [KH (3)] WHERE Ten = @Ten"

        Dim parameters As New List(Of String)

        Dim command As New ADODB.Command
        command.ActiveConnection = connection

        command.Parameters.Append(command.CreateParameter(Name:="@Ten", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(Ten), Value:=Ten.ToString))
        If CCCD <> "" Then
            parameters.Add("CCCD = @CCCD")
            command.Parameters.Append(command.CreateParameter(Name:="@CCCD", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(CCCD), Value:=CCCD.ToString))
        End If

        If CMND <> "" Then
            parameters.Add("CMND = @CMND")
            command.Parameters.Append(command.CreateParameter(Name:="@CMND", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(CMND), Value:=CMND.ToString))
        End If

        If HC <> "" Then
            parameters.Add("HC = @HC")
            command.Parameters.Append(command.CreateParameter(Name:="@HC", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(HC), Value:=HC.ToString))
        End If

        If CMSQ <> "" Then
            parameters.Add("CMSQ = @CMSQ")
            command.Parameters.Append(command.CreateParameter(Name:="@CMSQ", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(CMSQ), Value:=CMSQ.ToString))
        End If

        If SDDCN <> "" Then
            parameters.Add("SDDCN = @SDDCN")
            command.Parameters.Append(command.CreateParameter(Name:="@SDDCN", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(SDDCN), Value:=SDDCN.ToString))
        End If

        If parameters.Count > 0 Then
            sql = sql & " AND " & String.Join(" OR ", parameters)
        End If
        command.CommandText = sql

        Dim rs As New ADODB.Recordset
        rs.Open(command)

        If Not rs.EOF And Not rs.BOF Then
            rs.Close()
            command.ActiveConnection = Nothing
            connection.Close()
            Return True
        End If

        rs.Close()
        command.ActiveConnection = Nothing
        connection.Close()
        Return False

    End Function
    Public Sub SaveCustomer()
        If Trim(ID) <> "" Then Exit Sub

        Dim connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Location
        Dim connection As New ADODB.Connection
        connection.Open(connectionString)

        Dim sqlString As String = "INSERT INTO [KH (3)] (Gt, Ten, CCCD, CMND, HC, CMSQ, SDDCN, TT) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
        Dim command As New ADODB.Command
        command.ActiveConnection = connection
        command.CommandText = sqlString

        command.Parameters.Append(command.CreateParameter(Name:="@Gt", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(Gt), Value:=Gt.ToString))
        command.Parameters.Append(command.CreateParameter(Name:="@Ten", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(Ten), Value:=Ten.ToString))
        command.Parameters.Append(command.CreateParameter(Name:="@CCCD", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(CCCD), Value:=CCCD.ToString))
        command.Parameters.Append(command.CreateParameter(Name:="@CMND", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(CMND), Value:=CMND.ToString))
        command.Parameters.Append(command.CreateParameter(Name:="@HC", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(HC), Value:=HC.ToString))
        command.Parameters.Append(command.CreateParameter(Name:="@CMSQ", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(CMSQ), Value:=CMSQ.ToString))
        command.Parameters.Append(command.CreateParameter(Name:="@SDDCN", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(SDDCN), Value:=SDDCN.ToString))
        command.Parameters.Append(command.CreateParameter(Name:="@TT", Type:=ADODB.DataTypeEnum.adVarChar, Direction:=ADODB.ParameterDirectionEnum.adParamInput, Size:=ReturnSizeOfProperty(TT), Value:=TT.ToString))
        command.Execute()

        command.ActiveConnection = Nothing
        connection.Close()

    End Sub
    Public Function ExportCustomer(Optional Type As ExportType = ExportType.Form) As String
        Dim exportString As String = ""

        Select Case Type
            Case ExportType.Form
                exportString = Gt & vbTab & ": " & Ten & vbCr & "Sinh năm" & vbTab & ": " & Sn & vbCr

                If CCCD <> "" Then exportString &= "Căn cước công dân số" & vbTab & ": " & AddSpaceCCCD(CCCD) & vbCr
                If CMND <> "" Then exportString &= "Chứng minh nhân dân số" & vbTab & ": " & AddSpaceCCCD(CMND) & vbCr
                If HC <> "" Then exportString &= "Hộ chiếu số" & vbTab & ": " & HC & vbCr
                If CMSQ <> "" Then exportString &= "Chứng minh sỹ quan số" & vbTab & ": " & CMSQ & vbCr
                If SDDCN <> "" Then exportString &= "Số định danh cá nhân" & vbTab & ": " & SDDCN & vbCr

                exportString &= "Thường trú" & vbTab & ": " & TT

            Case ExportType.OneLine
                exportString = ID & ";" & Gt & ";" & Ten & ";" & Sn & ";" & CCCD & ";" & CMND & ";" & HC & ";" & CMSQ & ";" & SDDCN & ";" & TT
                exportString = Regex.Replace(exportString, ";+", "; ")
        End Select

        Return exportString

    End Function
    
    Public Function GetPropertyValues() As Dictionary(Of String, String)
        Dim properties As PropertyInfo() = Me.GetType().GetProperties()
        Dim propertyDictionary As New Dictionary(Of String, String)
        For Each MeProperty As PropertyInfo In properties
            Dim propertyName As String = MeProperty.Name
            Dim propertyValue As Object = MeProperty.GetValue(Me)
            propertyDictionary.Add(propertyName, propertyValue)
        Next
        Return propertyDictionary
    End Function

End Class
