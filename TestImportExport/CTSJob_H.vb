'-----Class Definition-----
Imports System.Data.SqlClient
Public Class CCTSJob_H
    Private m_ConnStr As String
    Public Sub New()

    End Sub
    Public Sub New(pConnStr As String)
        m_ConnStr = pConnStr
    End Sub
    Public Sub SetConnect(pConnStr As String)
        m_ConnStr = pConnStr
    End Sub

    Private m_Id As Integer
    Public Property Id As Integer
        Get
            Return m_Id
        End Get
        Set(value As Integer)
            m_Id = value
        End Set
    End Property
    Private m_fiscalPeriod As String
    Public Property fiscalPeriod As String
        Get
            Return m_fiscalPeriod
        End Get
        Set(value As String)
            m_fiscalPeriod = value
        End Set
    End Property
    Private m_jobNo As String
    Public Property jobNo As String
        Get
            Return m_jobNo
        End Get
        Set(value As String)
            m_jobNo = value
        End Set
    End Property
    Private m_jobDate As Date
    Public Property jobDate As Date
        Get
            Return m_jobDate
        End Get
        Set(value As Date)
            m_jobDate = value
        End Set
    End Property
    Private m_agentId As String
    Public Property agentId As String
        Get
            Return m_agentId
        End Get
        Set(value As String)
            m_agentId = value
        End Set
    End Property
    Private m_agentJobPicPath As String
    Public Property agentJobPicPath As String
        Get
            Return m_agentJobPicPath
        End Get
        Set(value As String)
            m_agentJobPicPath = value
        End Set
    End Property
    Private m_companyId As String
    Public Property companyId As String
        Get
            Return m_companyId
        End Get
        Set(value As String)
            m_companyId = value
        End Set
    End Property
    Private m_companyBranch As String
    Public Property companyBranch As String
        Get
            Return m_companyBranch
        End Get
        Set(value As String)
            m_companyBranch = value
        End Set
    End Property
    Private m_companyAuthPicPath As String
    Public Property companyAuthPicPath As String
        Get
            Return m_companyAuthPicPath
        End Get
        Set(value As String)
            m_companyAuthPicPath = value
        End Set
    End Property
    Private m_supportBy As String
    Public Property supportBy As String
        Get
            Return m_supportBy
        End Get
        Set(value As String)
            m_supportBy = value
        End Set
    End Property
    Private m_customsPort As String
    Public Property customsPort As String
        Get
            Return m_customsPort
        End Get
        Set(value As String)
            m_customsPort = value
        End Set
    End Property
    Private m_jobStatus As String
    Public Property jobStatus As String
        Get
            Return m_jobStatus
        End Get
        Set(value As String)
            m_jobStatus = value
        End Set
    End Property
    Private m_lastUpdate As Date
    Public Property lastUpdate As Date
        Get
            Return m_lastUpdate
        End Get
        Set(value As Date)
            m_lastUpdate = value
        End Set
    End Property
    Private m_updateBy As String
    Public Property updateBy As String
        Get
            Return m_updateBy
        End Get
        Set(value As String)
            m_updateBy = value
        End Set
    End Property
    Private m_remark As String
    Public Property remark As String
        Get
            Return m_remark
        End Get
        Set(value As String)
            m_remark = value
        End Set
    End Property
    Public Function SaveData(pSQLWhere As String) As String
        Dim msg As String = ""
        Using cn As New SqlConnection(m_ConnStr)
            Try
                cn.Open()

                Using da As New SqlDataAdapter("SELECT * FROM dbo.CTSJob_H" & pSQLWhere, cn)
                    Using cb As New SqlCommandBuilder(da)
                        Using dt As New DataTable
                            da.Fill(dt)
                            Dim dr As DataRow = dt.NewRow
                            If dt.Rows.Count > 0 Then dr = dt.Rows(0)
                            If Me.Id = 0 Then
                                Dim retStr As String = Main.GetMaxByMask(m_ConnStr, String.Format("SELECT MAX(Id) as t FROM CTSJob_H "), "____")
                                Id = Convert.ToInt32("0" & retStr)
                            End If
                            dr("Id") = Me.Id
                            dr("fiscalPeriod") = Me.fiscalPeriod
                            dr("jobNo") = Me.jobNo
                            dr("jobDate") = Main.GetDBDate(Me.jobDate)
                            dr("agentId") = Me.agentId
                            dr("agentJobPicPath") = Me.agentJobPicPath
                            dr("companyId") = Me.companyId
                            dr("companyBranch") = Me.companyBranch
                            dr("companyAuthPicPath") = Me.companyAuthPicPath
                            dr("supportBy") = Me.supportBy
                            dr("customsPort") = Me.customsPort
                            dr("jobStatus") = Me.jobStatus
                            dr("lastUpdate") = Main.GetDBDate(Me.lastUpdate)
                            dr("updateBy") = Me.updateBy
                            dr("remark") = Me.remark
                            If dr.RowState = DataRowState.Detached Then dt.Rows.Add(dr)
                            da.Update(dt)
                            msg = "Save Complete"
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                msg = ex.Message
            End Try
        End Using
        Return msg
    End Function
    Public Sub AddNew()
        Id = 0
    End Sub
    Public Function GetData(pSQLWhere As String) As List(Of CCTSJob_H)
        Dim lst As New List(Of CCTSJob_H)
        Using cn As New SqlConnection(m_ConnStr)
            Dim row As CCTSJob_H
            Try
                cn.Open()
                Dim rd As SqlDataReader = New SqlCommand("SELECT * FROM dbo.CTSJob_H" & pSQLWhere, cn).ExecuteReader()
                While rd.Read()
                    row = New CCTSJob_H(m_ConnStr)
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("Id"))) = False Then
                        row.Id = rd.GetInt32(rd.GetOrdinal("Id"))
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("fiscalPeriod"))) = False Then
                        row.fiscalPeriod = rd.GetString(rd.GetOrdinal("fiscalPeriod")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("jobNo"))) = False Then
                        row.jobNo = rd.GetString(rd.GetOrdinal("jobNo")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("jobDate"))) = False Then
                        row.jobDate = rd.GetValue(rd.GetOrdinal("jobDate"))
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("agentId"))) = False Then
                        row.agentId = rd.GetString(rd.GetOrdinal("agentId")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("agentJobPicPath"))) = False Then
                        row.agentJobPicPath = rd.GetString(rd.GetOrdinal("agentJobPicPath")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("companyId"))) = False Then
                        row.companyId = rd.GetString(rd.GetOrdinal("companyId")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("companyBranch"))) = False Then
                        row.companyBranch = rd.GetString(rd.GetOrdinal("companyBranch")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("companyAuthPicPath"))) = False Then
                        row.companyAuthPicPath = rd.GetString(rd.GetOrdinal("companyAuthPicPath")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("supportBy"))) = False Then
                        row.supportBy = rd.GetString(rd.GetOrdinal("supportBy")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("customsPort"))) = False Then
                        row.customsPort = rd.GetString(rd.GetOrdinal("customsPort")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("jobStatus"))) = False Then
                        row.jobStatus = rd.GetString(rd.GetOrdinal("jobStatus")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("lastUpdate"))) = False Then
                        row.lastUpdate = rd.GetValue(rd.GetOrdinal("lastUpdate"))
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("updateBy"))) = False Then
                        row.updateBy = rd.GetString(rd.GetOrdinal("updateBy")).ToString()
                    End If
                    If IsDBNull(rd.GetValue(rd.GetOrdinal("remark"))) = False Then
                        row.remark = rd.GetString(rd.GetOrdinal("remark")).ToString()
                    End If
                    lst.Add(row)
                End While
            Catch ex As Exception
            End Try
        End Using
        Return lst
    End Function
    Public Function DeleteData(pSQLWhere As String) As String
        Dim msg As String = ""
        Using cn As New SqlConnection(m_ConnStr)
            Try
                cn.Open()

                Using cm As New SqlCommand("DELETE FROM dbo.CTSJob_H" + pSQLWhere, cn)
                    cm.CommandTimeout = 0
                    cm.CommandType = CommandType.Text
                    cm.ExecuteNonQuery()
                End Using
                cn.Close()
                msg = "Delete Complete"
            Catch ex As Exception
                msg = ex.Message
            End Try
        End Using
        Return msg
    End Function
End Class
