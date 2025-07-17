Imports System.IO
Imports System.Linq
Public Class frmEqualize_GenerateEqualisedReport

    'Dim objEqualize_clsBal As clsEqualize_BusinessAccessLayer
    Dim dtFundData As DataTable
    Dim dtSchemeData As DataTable
    Dim dtToEqualizeScheme As DataTable
    Dim dtToEqualPlanData As DataTable
    Dim dtDateData As DataTable
    Dim dtDgvSchemeData As DataTable
    Dim dtBeforeFromTemplate As New DataTable
    Dim dtBtweenFromToTemplate As New DataTable
    Dim dtBtweenFromToOpenigBalEffDate As New DataTable    'Added by Tushar as 31082012 to add the opening balance effective date.
    Dim dtTemplateColData As DataTable
    Dim dtSampleColData As DataTable
    Dim dtInputData As New DataTable
    Dim dtPlanData As New DataTable
    Dim dtEqualizeData As DataTable
    Dim dtRptTemplateData As DataTable
    Dim dtLastEqualizeDate As DataTable
    Dim dtOpeningData As DataTable

    Dim objclsExcel As clsEqualization_Excel
    'Dim dtTemplateColData As DataTable
    Dim DtColNames As DataTable
    Dim dtColtemp As DataColumn

    Dim IsFormInit As Boolean = False
    Dim strDate As String
    Dim strFundCode As String
    Dim drSelect() As DataRow
    Dim drSelectedScheme() As DataRow
    Dim drAddRow As DataRow
    Dim drAddRow1 As DataRow

    Dim strFromDate As String
    Dim newStrFromDate As String
    Dim SchemeOpenDate As String
    Dim strToDate As String
    Dim lngDateDiff As Long

    Dim SchemeId As Long
    Dim SchemeCode As String
    Dim PlanCode As String
    Dim strDateNotFound As String


    Dim XlProcessId As Integer
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim xlWorkSheetSample As Excel.Worksheet
    Dim xlWorkSheetCurrScheme As Excel.Worksheet

    Dim rptType As String
    Dim IsOpenDate As Boolean

    Dim RptTemplateId As Long
    Dim strQuery As String
    Dim strValuesQuery As String
    Dim Val As String
    Dim xlRange As String

    Dim strColID As String
    Dim strColValue As String

    Dim planStartDate As String
    Dim TemplateEffDt As String


    Dim userChngData As Boolean

    'Added by shweta (29 Aug 2011)
    'To add functionality for Yes to all and no to all
    Dim strYesNoFlag As String
    Dim dtGLReportData As DataTable
    Dim lngGLRptRowNum As Long

    'For Dividend Distribution Report
    Dim dtDivDistData As DataTable



    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            objfrmEqualizeGenerateEqualisedRpt.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub frmEqualize_GenerateEqualisedReport_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            objfrmEqualizeGenerateEqualisedRpt = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub frmEqualize_GenerateEqualisedReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            userChngData = True
            If IsNothing(objEqualizeDefaultReportPath) Then
                FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop
            Else
                FolderBrowserDialog1.SelectedPath = objEqualizeDefaultReportPath
            End If
            Dim buttonToolTip As New ToolTip()
            objEqualize_clsBal.SetToolTip(BtnClose, "To Exit Form.(Alt+C)", buttonToolTip)
            objEqualize_clsBal.SetToolTip(BtnGenerate, "To Generate Report.(Alt+G)", buttonToolTip)
            objEqualize_clsBal.SetToolTip(btnOkPLan, "To Close Panel .(Alt+O)", buttonToolTip)

            pnlMainData.Enabled = True
            pnlSchemeData.Enabled = False
            pnlSchemeData.Visible = False
            pnlSchemeData.SendToBack()
            pnlMainData.BringToFront()

            dtFundData = New DataTable
            dtFundData = objEqualize_clsDAL.GetApprovedData("User Mutual Fund Master", ClsCommon.userName)
            cmbMFund.DisplayMember = "Mutual_Fund_Code"
            cmbMFund.ValueMember = "Mutual_Fund_Name"
            cmbMFund.DataSource = dtFundData
            Dim dt As String = System.DateTime.Now.ToString("dd-MMM-yyyy")
            dtToEqualPlanData = New DataTable
            dtToEqualPlanData = objEqualize_clsDAL.GetApprovedData("GetToEqualizeData", "", dt)
            IsFormInit = True

            dtpBaseDate.Value = System.DateTime.Now.ToString("dd-MMM-yyyy")
            dtpLastRecordDate.Value = System.DateTime.Now.ToString("dd-MMM-yyyy")

            REM added by sameer on 18-Jan-2013
            With_Formula_Rights()
            pnldistrpt.Visible = False
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub
    REM added by sameer on 18-Jan-2013
    Private Sub With_Formula_Rights()
        Try
            Dim _ds As New DataSet
            _ds = objEqualize_clsDAL.USP_FormulaRight("EQUALIZE")

            If _ds.Tables(0).Rows.Count > 0 Then
                chk_WithFormula.Visible = True
            Else
                chk_WithFormula.Visible = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub cmbMFund_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMFund.SelectedIndexChanged
        Try
            If IsFormInit Then
                getSchemeData()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub rBtnEqualize_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rBtnEqualize.CheckedChanged
        Try
            If IsFormInit Then
                CheckedSchemeToEqualize()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub DTPFromDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPFromDate.ValueChanged
        Try
            If IsFormInit Then
                CheckedSchemeToEqualize(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub dtpToDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpToDate.ValueChanged
        Try
            If IsFormInit Then
                CheckedSchemeToEqualize(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub btnOkPLan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOkPLan.Click
        Try
            pnlMainData.Enabled = True
            pnlSchemeData.Enabled = False
            pnlSchemeData.Visible = False
            pnlSchemeData.SendToBack()
            pnlMainData.BringToFront()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub dgvSchemes_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSchemes.CellContentClick
        Try
            Dim ToDt As String
            Dim FromDt As String
            Dim SchemeId As String
            If e.RowIndex >= 0 Then
                If Me.dgvSchemes.CurrentCell.ColumnIndex = dgvSchemes.Columns("colSchemeCode").Index Then
                    FromDt = Convert.ToDateTime(DTPFromDate.Text).ToString("dd-MMM-yyyy")
                    ToDt = Convert.ToDateTime(dtpToDate.Text).ToString("dd-MMM-yyyy")
                    lblSchemeCd.Text = dgvSchemes.Rows(e.RowIndex).Cells("colSchemeCode").Value
                    lblSchemeNm.Text = dgvSchemes.Rows(e.RowIndex).Cells("colSchemeName").Value
                    SchemeId = dgvSchemes.Rows(e.RowIndex).Cells("AutoID").Value.ToString
                    dtPlanData = objEqualize_clsDAL.GetApprovedData("PlanWithDivCal", "", SchemeId, ToDt, FromDt)
                    lvPlanData.Items.Clear()
                    For index As Integer = 0 To dtPlanData.Rows.Count - 1
                        lvPlanData.Items.Add(dtPlanData.Rows(index)("Plan_Code").ToString)
                        lvPlanData.Items(index).SubItems.Add(dtPlanData.Rows(index)("Plan_Name").ToString)
                        lvPlanData.Items(index).SubItems.Add(dtPlanData.Rows(index)("Dividend_Frequency").ToString)
                    Next
                    pnlMainData.Enabled = False
                    pnlSchemeData.Enabled = True
                    pnlSchemeData.Visible = True
                    pnlSchemeData.BringToFront()
                    pnlMainData.SendToBack()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub BtnGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnGenerate.Click
        Try
            lstEqulizeStatus.Items.Clear()
            'Dim MFundId As Long
            Dim report_Type As String = "Generate Report"
            If chkEquReport.Checked = False And chkGenerateGLRpt.Checked = False And chkDivDistRpt.Checked = False And chk_DistriSurplus.Checked = False And chkDistributablerpt.Checked = False Then
                MessageBox.Show("Please Select Report to genarate.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            ElseIf chkEquReport.Checked = False And chkGenerateGLRpt.Checked Then
                report_Type = "Generate GL Report"
            End If
            Dim strRptPath As String = TxtReportPath.Text.Trim
            If strRptPath = "" Then
                MessageBox.Show("Report Save path is not assigned. Please assign the report save path first", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            If cmbRptFormat.SelectedIndex < 0 And chkEquReport.Checked Then
                MessageBox.Show("Please select type of Report ", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            'If rBtnEqualize.Checked Then
            '    If cmbMFund.SelectedIndex < 0 Then
            '        MessageBox.Show("Please Select the Mutual Fund", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            '        Exit Sub
            '    End If
            '    strFundCode = cmbMFund.GetItemText(cmbMFund.SelectedItem)
            '    drSelect = dtFundData.Select("Mutual_Fund_Code='" & strFundCode & "'")
            '    MFundId = 0
            '    If drSelect.Length > 0 Then
            '        MFundId = drSelect(0)("AutoId")
            '    End If
            '    Dim dataInUsed As Long = objEqualize_clsDAL.LockFund(MFundId) 
            '    If dataInUsed <= 0 Then
            '        MessageBox.Show("Selected Fund Is in Use. Please Try Again later.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            '        'objExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
            '        Exit Sub
            '    End If
            'End If

            REM added by sameer on 16-Jan-13
            If chk_WithFormula.Checked Then
                If MsgBox("Are you sure, you want the report with formula?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "DBEqualization") = MsgBoxResult.No Then
                    chk_WithFormula.Checked = False
                    Exit Sub
                End If
            End If

            REM added by sameer on 13-Mar-2013
            If chk_DistriSurplus.Checked = True Then
                Dim MFId As Integer
                drSelect = dtFundData.Select("Mutual_Fund_Code='" & strFundCode & "'")
                MFId = 0
                If drSelect.Length > 0 Then
                    MFId = drSelect(0)("AutoId")
                End If
                Generate_DistSurplus(MFId)
            ElseIf chkDistributablerpt.Checked = True Then
                ''Added by vijay 12072013 Start
                Dim MFId As Integer
                drSelect = dtFundData.Select("Mutual_Fund_Code='" & strFundCode & "'")
                MFId = 0
                If drSelect.Length > 0 Then
                    MFId = drSelect(0)("AutoId")
                End If
                Generate_DistReport(MFId)
                ''Added by vijay 12072013 End
            Else
                EqualizeData(report_Type)
                chk_WithFormula.Checked = False
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub Generate_DistReport(ByVal MFundID As Integer)
        Try
            Dim Rqry As String
            If Not Directory.Exists(TxtReportPath.Text.Trim) Then
                MessageBox.Show("Selected Path not exist. Please Select the Proper Path.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If

            If DateDiff(DateInterval.Day, DTPFromDate.Value, dtpToDate.Value) > 0 Then
                MessageBox.Show("This report is for single date plz select proper date.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If

            Dim _distSrp As DataTable
            Dim DS_str, DS_SchemeID As String
            'If _distSrp Is Nothing Then _distSrp = New DataSet
            '_distSrp = objEqualize_clsDAL.USP_DistrSurplus("COL_HEADER", MFundID)
            Rqry = "Select distinct isnull('ColumnValue' + convert(varchar,TC.ColSequenceNum),'') ColumnValue,TC.ColSequenceNum,DS.column_name,DS.Column_id" & vbCrLf
            Rqry &= " from Equalize_Distributable_Report DS" & vbCrLf
            Rqry &= "left Join Equalize_TxnTemplateColumn TC on DS.Template_Column=tc.ColHeader " & vbCrLf
            Rqry &= " and TC.TemplateID in(Select max(rptTemplateid) from Equalize_EqualizationData where FundId =" & MFundID & ")" & vbCrLf
            Rqry &= " Where DS.is_approved=1 and DS.mfundID=" & MFundID & " order by DS.Column_id"

            _distSrp = objEqualize_clsDAL.FillDataSet(Rqry)
            If _distSrp.DefaultView.ToTable(True, "ColumnValue").Rows.Count > 1 Then
            Else
                MsgBox("Data is not available for selected combination", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                Exit Sub
            End If

            Dim str, str1 As String
            If _distSrp Is Nothing = False Then
                If _distSrp.Rows.Count > 0 Then
                    DS_str = ""
                    For i As Integer = 0 To _distSrp.Rows.Count - 1
                        If _distSrp.Rows(i)("ColumnValue").ToString = "" Then
                            If _distSrp.Rows(i)("column_name").ToString = "DIVIDEDDIST" Then
                                DS_str &= "'=IF(K3<0,H3-1000-J3,H3-1000-J3-K3)'"
                                str = "Cast(isnull(" & _distSrp.Rows(i)("columnValue").ToString & ",0)as decimal(38,5))"
                            ElseIf _distSrp.Rows(i)("column_name").ToString = "Units" Then
                                'str1 = "Cast(replace(isnull(" & _distSrp.Rows(i)("columnValue").ToString & ",1),0,1)as decimal(38,5))"
                                str1 = "Cast(isnull(case when " & _distSrp.Rows(i)("columnValue").ToString & " = '0' then '1' else " & _distSrp.Rows(i)("columnValue").ToString & " end ,1)as decimal(38,5))"
                                DS_str &= str & "/" & str1
                            Else
                                DS_str &= "''"
                            End If

                        Else
                            If _distSrp.Rows(i)("column_name").ToString = "DIVIDEDDIST" Then
                                DS_str &= "'=IF(K3<0,H3-1000-J3,H3-1000-J3-K3)'"
                                str = "cast(isnull(" & _distSrp.Rows(i)("columnValue").ToString & ",0)as decimal(38,5))"
                            ElseIf _distSrp.Rows(i)("column_name").ToString = "Units" Then
                                'str1 = "Cast(replace(isnull(" & _distSrp.Rows(i)("columnValue").ToString & ",1),0,1)as decimal(38,5))"
                                str1 = "Cast(isnull(case when " & _distSrp.Rows(i)("columnValue").ToString & " = '0' then '1' else " & _distSrp.Rows(i)("columnValue").ToString & " end ,1)as decimal(38,5))"
                                DS_str &= str & "/" & str1
                            Else
                                DS_str &= _distSrp.Rows(i)("ColumnValue").ToString
                            End If
                        End If
                        DS_str &= " as [" & _distSrp.Rows(i)("column_name").ToString & "],"
                    Next
                    DS_str = DS_str.Substring(0, DS_str.Length - 1)
                Else
                    MsgBox("Data is not available for selected combination", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                    Exit Sub
                End If
            Else
                MsgBox("Data is not available for selected combination", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                Exit Sub
            End If
            DS_SchemeID = ""
            For i As Integer = 0 To dgvSchemes.Rows.Count - 1
                If dgvSchemes.Rows(i).Cells(0).Value = True Then
                    DS_SchemeID &= dgvSchemes.Rows(i).Cells(4).Value & ","
                End If
            Next
            DS_SchemeID = DS_SchemeID.Substring(0, DS_SchemeID.Length - 1)

            Dim _dsdt, _dsdt1, _dsdt2 As DataTable
            REM report qry

            Rqry = "Select a.AMFICode,a.[CAMS Code],a.SchemeName ,a.PlanCode ,a.PlanName ,a.[Last Record Date],a.NAV,b.NAV ,isnull(b.[As Per AMFI],0)as 'As Per AMFI',isnull(b.[As Per Unit],0)as 'As Per Unit',isnull(b.[UPR Per Unit],0) as 'UPR Per Unit',a.DIVIDEDDIST,a.Units ,b.Remarks ,isnull(a.[Closing unit],0)as 'BaseClosingUnit' ,isnull(b.[Closing unit],0) as 'CurrentClosingUnit'  from (Select distinct pm.AutoId as 'PlanId',pm.AMFICode,pm.RnTCode as'CAMS Code',sm.SchemeName,pm.PlanCode,pm.PlanName,'" & Convert.ToDateTime(dtpLastRecordDate.Value).ToString("dd-MMM-yyyy") & "'as 'Last Record Date' " & vbCrLf
            Rqry &= "," & DS_str & vbCrLf
            Rqry &= "from Equalize_MstScheme SM" & vbCrLf
            Rqry &= "join Equalize_MstPlan PM on SM.AutoId =PM.SchemeID and SM.IsApproved =PM.IsApproved " & vbCrLf
            Rqry &= "Join Equalize_EqualizationData ED on ED.FundId =SM.MFundID and ED.SchemeId =SM.AutoId " & vbCrLf
            Rqry &= "and ED.IsApproved =SM.IsApproved AND ED.PlanId =PM.AutoId " & vbCrLf
            Rqry &= "Where SM.MFundID =" & MFundID & " and SM.AutoId in (" & DS_SchemeID & ") and SM.IsApproved=1" & vbCrLf
            Rqry &= "and ED.EqualizeDate = '" & Convert.ToDateTime(dtpLastRecordDate.Value).ToString("dd-MMM-yyyy") & "')a" & vbCrLf
            Rqry &= " join(Select distinct pm.AMFICode,pm.RnTCode as'CAMS Code',sm.SchemeName,pm.PlanCode,pm.PlanName,'" & Convert.ToDateTime(dtpLastRecordDate.Value).ToString("dd-MMM-yyyy") & "'as 'Last Record Date' " & vbCrLf
            Rqry &= "," & DS_str & vbCrLf
            Rqry &= "from Equalize_MstScheme SM" & vbCrLf
            Rqry &= "join Equalize_MstPlan PM on SM.AutoId =PM.SchemeID and SM.IsApproved =PM.IsApproved " & vbCrLf
            Rqry &= "Join Equalize_EqualizationData ED on ED.FundId =SM.MFundID and ED.SchemeId =SM.AutoId " & vbCrLf
            Rqry &= "and ED.IsApproved =SM.IsApproved AND ED.PlanId =PM.AutoId " & vbCrLf
            Rqry &= "Where SM.MFundID =" & MFundID & " and SM.AutoId in (" & DS_SchemeID & ") and SM.IsApproved=1" & vbCrLf
            Rqry &= "and ED.EqualizeDate = '" & Convert.ToDateTime(dtpBaseDate.Value).ToString("dd-MMM-yyyy") & "')b on a.SchemeName=b.SchemeName and a.PlanCode=b.PlanCode " & vbCrLf
            Rqry &= " join (Select GrowthPlan_ID ,a.Plan_ID    from Equalize_Dist_Rp_Plan_Map a "
            Rqry &= "join Equalize_MstPlan b "
            Rqry &= " on a.Plan_ID =b.AutoId and a.SchemeID =b.SchemeID and a.Is_Approved=1 and b.IsApproved=1"
            Rqry &= "WHERE a.MFundID=" & MFundID & " and a.SchemeID  in(" & DS_SchemeID & ")) c "
            Rqry &= " on a.PlanId=c.Plan_ID"
            Rqry &= " group by c.Plan_ID,a.AMFICode,a.[CAMS Code],a.SchemeName ,a.PlanCode ,a.PlanName ,a.[Last Record Date],a.NAV,b.NAV ,isnull(b.[As Per AMFI],0) ,isnull(b.[As Per Unit],0) ,isnull(b.[UPR Per Unit],0),a.DIVIDEDDIST,a.Units ,b.Remarks ,isnull(a.[Closing unit],0),isnull(b.[Closing unit],0)"
            Rqry &= " Order by  c.Plan_ID"
            _dsdt = objEqualize_clsDAL.FillDataSet(Rqry)

            Rqry = "Select a.NAV as 'BaseNAV',b.NAV as 'CurrentNAV'  from (Select distinct sm.MFundID ,sm.AutoId as 'SchemeId' ,pm.AutoId as 'PlanId', pm.AMFICode,pm.RnTCode as'CAMS Code',sm.SchemeName,pm.PlanCode,pm.PlanName,'" & Convert.ToDateTime(dtpLastRecordDate.Value).ToString("dd-MMM-yyyy") & "'as 'Last Record Date' " & vbCrLf
            Rqry &= "," & DS_str & vbCrLf
            Rqry &= "from Equalize_MstScheme SM" & vbCrLf
            Rqry &= "join Equalize_MstPlan PM on SM.AutoId =PM.SchemeID and SM.IsApproved =PM.IsApproved " & vbCrLf
            Rqry &= "Join Equalize_EqualizationData ED on ED.FundId =SM.MFundID and ED.SchemeId =SM.AutoId " & vbCrLf
            Rqry &= "and ED.IsApproved =SM.IsApproved AND ED.PlanId =PM.AutoId " & vbCrLf
            Rqry &= "Where SM.MFundID =" & MFundID & " and SM.AutoId in (" & DS_SchemeID & ") and SM.IsApproved=1" & vbCrLf
            Rqry &= "and ED.EqualizeDate = '" & Convert.ToDateTime(dtpLastRecordDate.Value).ToString("dd-MMM-yyyy") & "')a" & vbCrLf
            Rqry &= " join(Select distinct pm.AMFICode,pm.RnTCode as'CAMS Code',sm.SchemeName,pm.PlanCode,pm.PlanName,'" & Convert.ToDateTime(dtpLastRecordDate.Value).ToString("dd-MMM-yyyy") & "'as 'Last Record Date' " & vbCrLf
            Rqry &= "," & DS_str & vbCrLf
            Rqry &= "from Equalize_MstScheme SM" & vbCrLf
            Rqry &= "join Equalize_MstPlan PM on SM.AutoId =PM.SchemeID and SM.IsApproved =PM.IsApproved " & vbCrLf
            Rqry &= "Join Equalize_EqualizationData ED on ED.FundId =SM.MFundID and ED.SchemeId =SM.AutoId " & vbCrLf
            Rqry &= "and ED.IsApproved =SM.IsApproved AND ED.PlanId =PM.AutoId " & vbCrLf
            Rqry &= "Where SM.MFundID =" & MFundID & " and SM.AutoId in (" & DS_SchemeID & ") and SM.IsApproved=1" & vbCrLf
            Rqry &= "and ED.EqualizeDate = '" & Convert.ToDateTime(dtpBaseDate.Value).ToString("dd-MMM-yyyy") & "')b on a.SchemeName=b.SchemeName and a.PlanCode=b.PlanCode " & vbCrLf
            Rqry &= " join (Select GrowthPlan_ID ,a.Plan_ID    from Equalize_Dist_Rp_Plan_Map a "
            Rqry &= "join Equalize_MstPlan b "
            Rqry &= " on a.Plan_ID =b.AutoId and a.SchemeID =b.SchemeID and a.Is_Approved=1 and b.IsApproved=1"
            Rqry &= "WHERE a.MFundID=" & MFundID & " and a.SchemeID  in(" & DS_SchemeID & ")) c "
            Rqry &= " on a.PlanId=c.GrowthPlan_ID "
            Rqry &= "Order by  c.Plan_ID"
            _dsdt1 = objEqualize_clsDAL.FillDataSet(Rqry)

            If _dsdt Is Nothing = False Then
                If _dsdt.Rows.Count > 0 Then
                    Dim DS_path As String
                    DS_path = Create_Folder(TxtReportPath.Text.Trim, "Distributable Report")
                    DS_path &= "\Distributable Report - " & Convert.ToDateTime(dtpBaseDate.Value).ToString("dd-MMM-yyyy") & ".xlsx"

                    If Write_To_Excel(DS_path, _dsdt, _dsdt1, _dsdt2) = True Then
                        MsgBox("Report generated successfully at : " & DS_path, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                    Else
                        MsgBox("Report is not generated", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                    End If
                Else
                    MsgBox("Data not available for selected combination", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub
    Public Function Write_To_Excel(ByVal PATH As String, ByVal Data_Table As DataTable, ByVal Data_Table1 As DataTable, ByVal Data_Table2 As DataTable) As Boolean
        Dim obj_cls_DAL As New clsCRS_DataAccessLayer
        Dim xl_app As Excel.Application
        Dim xl_wbk As Excel.Workbook
        Dim xl_wsh As Excel.Worksheet
        Dim Temppath As String = strABilling_user_TemplatePath & "\DistributableReport.xlt"
        ''"C:\Documents and Settings\user\Desktop\DistributableReport.xlt"
        If File.Exists(Temppath) Then
        Else
            MessageBox.Show("Template Not found at path " & Temppath, "DBEqualization")
        End If
        Try


            Write_To_Excel = True
            xl_app = New Excel.Application
            xl_wbk = xl_app.Workbooks.Open(Temppath)
            xl_wsh = xl_wbk.ActiveSheet
            'xl_app.Visible = True

            For Each xl_wsh In xl_wbk.Worksheets
                If xl_wbk.Worksheets.Count > 1 Then
                    xl_wbk.Worksheets(2).Delete()
                End If
            Next

            xl_wsh = xl_wbk.Sheets(1)
            xl_wsh.Name = "Distributable Report"
            xl_wsh.Range("B1").Value = Convert.ToDateTime(dtpBaseDate.Value).ToString("dd-MMM-yyyy")
            Dim Obj_Arry(Data_Table.Rows.Count, 11) As Object
            For r As Integer = 0 To Data_Table.Rows.Count - 1
                For c As Integer = 0 To 11
                    Obj_Arry(r + 0, c) = Data_Table.Rows(r)(c)
                Next
            Next
            xl_wsh.Range("A3").Select()
            xl_wsh.Range("A3").Resize(Data_Table.Rows.Count + 1, 11).Value = Obj_Arry
            Dim j As Integer = 3
            For i As Integer = 0 To Data_Table.Rows.Count - 1
                xl_wsh.Range("M" & j).Value = Data_Table.Rows(i)("Units").ToString
                xl_wsh.Range("W" & j).Value = Data_Table.Rows(i)("BaseClosingUnit").ToString
                xl_wsh.Range("X" & j).Value = Data_Table.Rows(i)("CurrentClosingUnit").ToString
                j = j + 1
            Next

            xl_wsh.Range("A3").Select()
            xl_wsh.Range("A3").Resize(Data_Table.Rows.Count + 1, 11).Value = Obj_Arry
            xl_wsh.Range("A3").Resize(Data_Table.Rows.Count, Data_Table.Columns.Count - 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            xl_wsh.Range("W3").Resize(Data_Table.Rows.Count, 3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            xl_wsh.Range("AA3").Resize(Data_Table.Rows.Count, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous



            If Data_Table1.Rows.Count > 0 Then
                Dim Obj_Arry1(Data_Table1.Rows.Count, Data_Table1.Columns.Count) As Object
                For r As Integer = 0 To Data_Table1.Rows.Count - 1
                    For c As Integer = 0 To Data_Table1.Columns.Count - 1
                        Obj_Arry1(r + 0, c) = Data_Table1.Rows(r)(c)
                    Next
                Next
                xl_wsh.Range("Q3").Select()
                xl_wsh.Range("Q3").Resize(Data_Table1.Rows.Count, Data_Table1.Columns.Count).Value = Obj_Arry1
                xl_wsh.Range("Q3").Resize(Data_Table1.Rows.Count, Data_Table1.Columns.Count + 3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

            End If
            If Data_Table.Rows.Count > 1 Then
                xl_wsh.Range("L3").AutoFill(xl_wsh.Range("L3:L" & 4 + (Data_Table.Rows.Count - 2)), Excel.XlAutoFillType.xlFillDefault)
                xl_wsh.Range("N3").AutoFill(xl_wsh.Range("N3:N" & 4 + (Data_Table.Rows.Count - 2)), Excel.XlAutoFillType.xlFillDefault)
                xl_wsh.Range("O3").AutoFill(xl_wsh.Range("O3:O" & 4 + (Data_Table.Rows.Count - 2)), Excel.XlAutoFillType.xlFillDefault)
                xl_wsh.Range("Y3").AutoFill(xl_wsh.Range("Y3:Y" & 4 + (Data_Table.Rows.Count - 2)), Excel.XlAutoFillType.xlFillDefault)
                xl_wsh.Range("AA3").AutoFill(xl_wsh.Range("AA3:AA" & 4 + (Data_Table.Rows.Count - 2)), Excel.XlAutoFillType.xlFillDefault)
                xl_wsh.Range("AB3").AutoFill(xl_wsh.Range("AB3:AB" & 4 + (Data_Table.Rows.Count - 2)), Excel.XlAutoFillType.xlFillDefault)
            End If
            If Data_Table1.Rows.Count > 1 Then
                xl_wsh.Range("S3").AutoFill(xl_wsh.Range("S3:S" & 4 + (Data_Table1.Rows.Count - 2)), Excel.XlAutoFillType.xlFillDefault)
                xl_wsh.Range("T3").AutoFill(xl_wsh.Range("T3:T" & 4 + (Data_Table1.Rows.Count - 2)), Excel.XlAutoFillType.xlFillDefault)
            End If
            xl_wsh.Columns.AutoFit()

            xl_wbk.SaveAs(PATH)

        Catch ex As Exception
            MsgBox("Error Source : " & ex.Source & vbCrLf & "Error Message : " & ex.Message & vbCrLf & "Error Occured in Method:-" & vbCrLf & ex.StackTrace, MsgBoxStyle.Information, Me.Text)
            Write_To_Excel = False
        Finally
            xl_wbk.Close()
            xl_app.Quit()
            obj_cls_DAL.killExcelProcess()
        End Try
    End Function


    Public Sub EqualizeData(ByVal strEquOrRpt As String)
        Dim MFundId As Long
        Try
            Dim boolGetTempData As Boolean
            Dim boolRprGenrate As Boolean = False
            Dim usedWrkShts As Integer = 0
            Dim strRptPath As String = TxtReportPath.Text.Trim
            Dim SavePath As String = strRptPath
            'Dim fileDetail As IO.FileInfo
            Dim extc As String = ""
            'Dim FileNamelen As Integer
            Dim RptFileName As String = ""
            Dim strStartEqualizationDate As String
            Dim boolEqDataNotFound As Boolean = False

            lstEqulizeStatus.Items.Clear()

            strFundCode = cmbMFund.GetItemText(cmbMFund.SelectedItem)
            drSelect = dtFundData.Select("Mutual_Fund_Code='" & strFundCode & "'")
            MFundId = 0
            If drSelect.Length > 0 Then
                MFundId = drSelect(0)("AutoId")
            End If
            'If strEquOrRpt = "Generate Report" Then
            '    If strRptPath = "" Then
            '        MessageBox.Show("Report Save path is not assigned. Please assign the report save path first", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            '        Exit Sub
            '    End If
            '    If cmbRptFormat.SelectedIndex < 0 Then
            '        MessageBox.Show("Please select type of Report ", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            '        Exit Sub
            '    End If
            'End If
            If chkGenerateGLRpt.Checked Or chkDivDistRpt.Checked Or chkEquReport.Checked Then
                If strRptPath = "" Then
                    MessageBox.Show("Report Save path is not assigned. Please assign the report save path first", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Exit Sub
                End If
                If Not Directory.Exists(strRptPath) Then
                    MessageBox.Show("Selected Path not exist. Please Select the Proper Path.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Exit Sub
                End If
            End If
            '--Added by shweta(08 Feb 2012)
            'For Print
            Dim strPrintFilePath As String = ""
            If chkPrint.Checked Then
                If TxtReportPath.Text <> "" Then
                    strPrintFilePath = TxtReportPath.Text & "\Test.xlsx"
                    PrintDialog1.ShowDialog()
                End If
            End If
            '====================

            rptType = cmbRptFormat.GetItemText(cmbRptFormat.SelectedItem)
            strFromDate = Convert.ToDateTime(DTPFromDate.Text).ToString("dd-MMM-yyyy")
            'Convert.ToDateTime(DTPFromDate.Text).AddDays(-1).ToString("dd-MMM-yyyy")
            strStartEqualizationDate = strFromDate
            strToDate = dtpToDate.Text

            If Convert.ToDateTime(strFromDate) > Convert.ToDateTime(strToDate) Then
                MessageBox.Show("From date must be less than To Date", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If

            dtDgvSchemeData = dgvSchemes.DataSource
            If IsNothing(dtDgvSchemeData) Then
                MessageBox.Show("Scheme not exist to Equalize data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            If dtDgvSchemeData.Rows.Count < 0 Then
                MessageBox.Show("Scheme not exist to Equalize data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            drSelectedScheme = dtDgvSchemeData.Select("ColChk =True")
            'If drSelectedScheme.Length <= 0 Then
            '    MessageBox.Show("Please select atleast 1 scheme to Calculate Equalization", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            '    Exit Sub
            'End If
            Dim boolSchemeSelected As Boolean = False
            For index As Integer = 0 To dtDgvSchemeData.Rows.Count - 1
                If dtDgvSchemeData.Rows(index)("ColChk") Then
                    boolSchemeSelected = True
                    Exit For
                End If
            Next
            If boolSchemeSelected = False Then
                MessageBox.Show("Please select atleast 1 scheme for " & strEquOrRpt, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            dtGLReportData = New DataTable
            lngGLRptRowNum = 0
            dtDivDistData = New DataTable
            strFromDate = Convert.ToDateTime(strFromDate).ToString("dd-MMM-yyyy")
            If strEquOrRpt = "Generate Report" And chkEquReport.Checked Then 'If rBtnReport.Checked Then
                'fileDetail = My.Computer.FileSystem.GetFileInfo(strRptPath)
                'If extc.ToLower <> ".xls" And extc.ToLower <> ".xlsx" Then
                '    MessageBox.Show("Selected File Path  extension is invalid", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                '    Exit Sub
                'End If
                'FileNamelen = strRptPath.LastIndexOf(extc)
                'RptFileName = strRptPath.Substring(0, FileNamelen)
                'Dim DirName As String = "" '= strRptPath.Substring(0, strRptPath.LastIndexOf("\"))
                'If strRptPath.LastIndexOf("\") >= 0 Then
                '    DirName = strRptPath.Substring(0, strRptPath.LastIndexOf("\"))
                'End If
                If Not Directory.Exists(strRptPath) Then
                    MessageBox.Show("Selected Path not exist. Please Select the Proper Path.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Exit Sub
                End If
                extc = ".xlsx"
                If strRptPath.Substring(strRptPath.Length - 1, 1) = "\" Then
                    strRptPath = strRptPath & "Equalization Report" '& strFromDate & " To " & strToDate    
                Else
                    strRptPath = strRptPath & "\" & "Equalization Report" '& strFromDate & " To " & strToDate                   
                End If
                objEqualize_clsBal.NewFolderCheckOrCreate(strRptPath)
                RptFileName = strRptPath & "\" & strFundCode
                strRptPath = strRptPath & "\" & strFundCode & extc
                objclsExcel = New clsEqualization_Excel
                objclsExcel.Initialise_ExcelObj(xlApp, XlProcessId) ' xlWorkBook,
                'xlApp.DefaultSaveFormat = Excel.XlFileFormat.xlExcel7
                If rptType = "Plan on Single Worksheet" Then
                    CreateNewWorkbook(strRptPath)
                    usedWrkShts = 1
                End If
            ElseIf strEquOrRpt = "Equalization" Then
                objclsExcel = New clsEqualization_Excel
                objclsExcel.Initialise_ExcelObj(xlApp, XlProcessId) ' xlWorkBook,
                CreateNewWorkbook()
                usedWrkShts = 1
            End If
            Me.Cursor = Cursors.WaitCursor
            boolGetTempData = False
            For index As Integer = 0 To dtDgvSchemeData.Rows.Count - 1 'drSelectedScheme.Length - 1

                strFromDate = Convert.ToDateTime(DTPFromDate.Text).ToString("dd-MMM-yyyy")
                strStartEqualizationDate = strFromDate
                strToDate = dtpToDate.Text
                If dtDgvSchemeData.Rows(index)("ColChk") = False Then
                    Continue For
                End If

                'Start of Added By Tejas Munankar on 08 Sep 2014

                Dim dt1 As DataTable = objEqualize_clsDAL.GetShouldEqualize(DTPFromDate.Text, Convert.ToInt64(dtDgvSchemeData.Rows(index)("AutoId").ToString), StrProcess:="PendingApproval")

                'Added Below Code for desktop Testing
                'Dim Strsql As String = ""
                'Strsql = "Select * From Equalize_TxnInputData Where SchemeId = '" & Convert.ToInt64(dtDgvSchemeData.Rows(index)("AutoId").ToString) & "' and DataDate = '" & DTPFromDate.Text & "' And IsApproved In (0,2)"
                'Dim dt1 As DataTable = objEqualize_clsDAL.FillDataSet(Strsql)

                If dt1.Rows.Count > 0 Then
                    Me.Cursor = Cursors.Arrow
                    Dim lstItem As New ListViewItem
                    lstItem.Text = dtDgvSchemeData.Rows(index)("Scheme_Code").ToString
                    lstItem.SubItems.Add("Input data is pending for approval for date :" & DTPFromDate.Text & ".")
                    lstEqulizeStatus.Items.Add(lstItem)
                    MessageBox.Show(dtDgvSchemeData.Rows(index)("Scheme_Code").ToString & " requires to be Reauthorized!", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Continue For
                End If
                'END of Added By Tejas Munankar on 08 Sep 2014


                'Added By Sagar Shah on 08 May 2014

                'Start Of Changed By Tejas On 08 Sep 2014
                'Dim dt As DataTable = objEqualize_clsDAL.GetShouldEqualize(DTPFromDate.Text, Convert.ToInt64(dtDgvSchemeData.Rows(index)("AutoId").ToString))
                Dim dt As DataTable = objEqualize_clsDAL.GetShouldEqualize(DTPFromDate.Text, Convert.ToInt64(dtDgvSchemeData.Rows(index)("AutoId").ToString), StrProcess:="PendingApproval-6Eye")

                'Added Below Code for desktop Testing
                'Strsql = "Select * From Equalize_TxnInputData Where SchemeId = '" & Convert.ToInt64(dtDgvSchemeData.Rows(index)("AutoId").ToString) & "' and DataDate = '" & DTPFromDate.Text & "' And IsNewPlan = 1 And IsReApproved is Null And IsApproved In (0,1,2)"
                'Dim dt As DataTable = objEqualize_clsDAL.FillDataSet(Strsql)

                'End Of Changed By Tejas On 08 Sep 2014

                If dt.Rows.Count > 0 Then
                    Me.Cursor = Cursors.Arrow
                    Dim lstItem As New ListViewItem
                    lstItem.Text = dtDgvSchemeData.Rows(index)("Scheme_Code").ToString
                    lstItem.SubItems.Add("Requires 6 Eye Check as New Plan exists.")
                    lstEqulizeStatus.Items.Add(lstItem)
                    MessageBox.Show(dtDgvSchemeData.Rows(index)("Scheme_Code").ToString & " requires to be Reauthorized!", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Continue For
                End If
                'Added By Sagar Shah on 08 May 2014 END

                boolEqDataNotFound = False
                SchemeCode = dtDgvSchemeData.Rows(index)("Scheme_Code").ToString
                SchemeId = Convert.ToInt64(dtDgvSchemeData.Rows(index)("AutoId").ToString)
                dtPlanData = objEqualize_clsDAL.GetApprovedData("PlanWithDivCal", "", SchemeId, strToDate, strFromDate)
                If dtPlanData.Rows.Count <= 0 Then
                    Me.Cursor = Cursors.Arrow
                    Dim lstItem As New ListViewItem
                    lstItem.Text = SchemeCode
                    lstItem.SubItems.Add("Plan not Exist")
                    lstEqulizeStatus.Items.Add(lstItem)
                    Continue For
                End If
                newStrFromDate = strFromDate

                'if Data is never equalized for selected Scheme then start the equaliztion from Open Date
                SchemeOpenDate = "" '(Added By Shweta(19 Jun 2012))
                dtLastEqualizeDate = New DataTable
                dtLastEqualizeDate = objEqualize_clsDAL.GetApprovedData("Approved Input Data Open Date", SchemeId)
                If dtLastEqualizeDate.Rows.Count > 0 Then
                    SchemeOpenDate = dtLastEqualizeDate.Rows(0)("DataDate").ToString
                End If
                If SchemeOpenDate = "" Then
                    Dim lstItem As New ListViewItem
                    lstItem.Text = SchemeCode
                    lstItem.SubItems.Add("Input Data Not exist")
                    lstEqulizeStatus.Items.Add(lstItem)
                    Continue For
                End If
                If strEquOrRpt = "Equalization" Then
                    'Get the Last Equalized date 
                    dtLastEqualizeDate = New DataTable
                    dtLastEqualizeDate = objEqualize_clsDAL.GetApprovedData("Equalize_Data_LastEqualizeDate", "", SchemeId)
                    newStrFromDate = ""
                    If dtLastEqualizeDate.Rows.Count > 0 Then
                        newStrFromDate = dtLastEqualizeDate.Rows(0)("dataDate").ToString
                    End If
                    If newStrFromDate = "" Then
                        boolEqDataNotFound = False
                    End If

                    If newStrFromDate = "" Then
                        'if Equalization is never done for Selected Scheme
                        strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                        strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                    Else
                        If Convert.ToDateTime(newStrFromDate) < Convert.ToDateTime(strFromDate) Then
                            'If Start date is greater than last equalization date then start equalization after Last equalized date
                            strStartEqualizationDate = Convert.ToDateTime(newStrFromDate).AddDays(1).ToString("dd-MMM-yyyy")
                            strFromDate = Convert.ToDateTime(strStartEqualizationDate).AddDays(-1).ToString("dd-MMM-yyyy")
                        Else
                            If Convert.ToDateTime(SchemeOpenDate) > Convert.ToDateTime(strFromDate) Then
                                'If Start date is less than open date then start equalization then start equalization from open date
                                strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                                'Convert.ToDateTime(SchemeOpenDate).AddDays(2).ToString("dd-MMM-yyyy")
                                strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                            Else
                                'If Start date is greater than open date then start equalization from From date 
                                strStartEqualizationDate = Convert.ToDateTime(strFromDate).ToString("dd-MMM-yyyy")
                                'Select From date less than 2 day from equalization date or Scheme open date
                                If Convert.ToDateTime(strStartEqualizationDate).AddDays(-1) < Convert.ToDateTime(SchemeOpenDate) Then
                                    strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                                Else
                                    strFromDate = Convert.ToDateTime(strStartEqualizationDate).AddDays(-1).ToString("dd-MMM-yyyy")
                                End If
                            End If
                        End If
                    End If
                Else 'If strEquOrRpt = "Generate Report" And chkEquReport.Checked Then
                    'Check Data already Equalize or Not
                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalize_Data_From_To_To", SchemeId.ToString, DTPFromDate.Text, strToDate)
                    If dtEqualizeData.Rows.Count <= 0 Then
                        Dim lstItem As New ListViewItem
                        lstItem.Text = SchemeCode
                        lstItem.SubItems.Add("Data is not Equalized for selected period")
                        lstEqulizeStatus.Items.Add(lstItem)
                        Continue For
                    Else
                        'Get the Min Equalize date as Start Date
                        'And Max Equalize date as end date between the selected Period
                        strStartEqualizationDate = Convert.ToDateTime(dtEqualizeData.Rows(0)("dataDate").ToString).ToString("dd-MMM-yyyy")
                        If Convert.ToDateTime(strStartEqualizationDate).AddDays(-1) < Convert.ToDateTime(SchemeOpenDate) Then
                            strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                            'strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                        Else
                            strFromDate = Convert.ToDateTime(dtEqualizeData.Rows(0)("dataDate").ToString).AddDays(-1).ToString("dd-MMM-yyyy")
                        End If
                        strToDate = Convert.ToDateTime(dtEqualizeData.Rows(dtEqualizeData.Rows.Count - 1)("dataDate").ToString).ToString("dd-MMM-yyyy")
                    End If
                End If

                'If boolGetTempData = False Or newStrFromDate <> strFromDate Then
                'To check Template exist or not to create report
                dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Get_Template_Before_From_Date", SchemeId, strFromDate, strToDate, MFundId)
                If dtBeforeFromTemplate.Rows.Count <= 0 Then
                    Dim lstItem As New ListViewItem
                    lstItem.Text = SchemeCode
                    lstItem.SubItems.Add("Template not exist for Selected period")
                    lstEqulizeStatus.Items.Add(lstItem)
                    'Me.Cursor = Cursors.Arrow
                    'MessageBox.Show("Template Not exist for Selected period of Scheme Code = " & SchemeCode & "." & vbCrLf & "Please create Template to calculate Data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Continue For
                End If
                dtBtweenFromToTemplate = objEqualize_clsDAL.GetApprovedData("Get_Template_Between_From_Date", SchemeId, strFromDate, strToDate, MFundId)


                'merge first template to start the calculation in the one datatable
                drAddRow = dtBtweenFromToTemplate.NewRow
                For index1 As Integer = 0 To dtBeforeFromTemplate.Columns.Count - 1
                    drAddRow(index1) = dtBeforeFromTemplate.Rows(0)(index1)
                Next
                dtBtweenFromToTemplate.Rows.InsertAt(drAddRow, 0)
                dtBtweenFromToTemplate.AcceptChanges()

                ''Added by Tushar as on 31082012 for opening balance effective date - START
                'dtBtweenFromToTemplate.AcceptChanges()
                'dtBtweenFromToOpenigBalEffDate = objEqualize_clsDAL.GetApprovedData("Get_EffectiveDate_Between_From_Date_For_OpeningBalance", SchemeId, strFromDate, strToDate, MFundId)

                'For InsertRow1 As Integer = 0 To dtBtweenFromToOpenigBalEffDate.Rows.Count - 1
                '    drAddRow1 = dtBtweenFromToTemplate.NewRow
                '    For index1 As Integer = 0 To dtBtweenFromToOpenigBalEffDate.Columns.Count - 1
                '        drAddRow1(index1) = dtBtweenFromToOpenigBalEffDate.Rows(0)(index1)
                '    Next
                '    dtBtweenFromToTemplate.Rows.InsertAt(drAddRow1, 0)
                '    dtBtweenFromToTemplate.AcceptChanges()
                'Next
                'Dim dv As DataView = dtBtweenFromToTemplate.DefaultView
                'dv.Sort = dtBtweenFromToTemplate.Columns("EffectiveDate").ColumnName & " " & "asc"
                'dtBtweenFromToTemplate = dv.ToTable

                ''Added by Tushar as on 31082012 for opening balance effective date  -   END
                dtBtweenFromToTemplate.AcceptChanges()

                dtTemplateColData = New DataTable
                dtTemplateColData = objEqualize_clsDAL.GetApprovedData("Get_All_Selected_Template_Columns", SchemeId, strFromDate, strToDate, MFundId)
                boolGetTempData = True

                'Check Input data exist or not for selected From to To date.
                dtInputData = objEqualize_clsDAL.GetApprovedData("Input_Data_From_To_To", SchemeId.ToString, strFromDate, strToDate)
                If dtInputData.Rows.Count <= 0 Then
                    Dim lstItem As New ListViewItem
                    lstItem.Text = SchemeCode
                    lstItem.SubItems.Add("Input Data Not exist for selected period")
                    lstEqulizeStatus.Items.Add(lstItem)
                    'Me.Cursor = Cursors.Arrow
                    'MessageBox.Show("Input Data Not exist for selected period for Scheme code " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Continue For
                End If
                'Dim isValid As Boolean = True
                If strEquOrRpt = "Equalization" Then 'If rBtnEqualize.Checked Then
                    While xlWorkBook.Worksheets.Count > 1
                        xlWorkSheet = xlWorkBook.Worksheets(2)
                        xlWorkSheet.Delete()
                    End While
                    'Dim datauploaded As Boolean = SaveEqualData(MFundId, SchemeId, strFundCode, SchemeCode, strStartEqualizationDate, strFromDate, strToDate, dtPlanData, dtInputData) 'strStartEqualizationDate
                    'Changed By Shweta 14 Mar 2012
                    Dim datauploaded As Boolean = SaveEqualData(MFundId, SchemeId, strFundCode, SchemeCode, strStartEqualizationDate, strFromDate, strToDate, dtPlanData, dtInputData, SchemeOpenDate) 'strStartEqualizationDate
                End If


                'Added by Tushar as on 31082012 for opening balance effective date - START

                dtBtweenFromToTemplate = objEqualize_clsDAL.GetApprovedData("Get_EffectiveDate_Between_From_Date_For_OpeningBalance", SchemeId, strFromDate, strToDate, MFundId)

                'dtBtweenFromToTemplate.AcceptChanges()
                'dtBtweenFromToOpenigBalEffDate = objEqualize_clsDAL.GetApprovedData("Get_EffectiveDate_Between_From_Date_For_OpeningBalance", SchemeId, strFromDate, strToDate, MFundId)

                'For InsertRow1 As Integer = 0 To dtBtweenFromToOpenigBalEffDate.Rows.Count - 1
                '    drAddRow1 = dtBtweenFromToTemplate.NewRow
                '    For index1 As Integer = 0 To dtBtweenFromToOpenigBalEffDate.Columns.Count - 1
                '        drAddRow1(index1) = dtBtweenFromToOpenigBalEffDate.Rows(0)(index1)
                '    Next
                '    dtBtweenFromToTemplate.Rows.InsertAt(drAddRow1, 0)
                '    dtBtweenFromToTemplate.AcceptChanges()
                'Next
                'Dim dv As DataView = dtBtweenFromToTemplate.DefaultView
                'dv.Sort = dtBtweenFromToTemplate.Columns("EffectiveDate").ColumnName & " " & "asc"
                'dtBtweenFromToTemplate = dv.ToTable

                'Added by Tushar as on 31082012 for opening balance effective date  -   END
                dtBtweenFromToTemplate.AcceptChanges()


                If strEquOrRpt = "Generate Report" And chkEquReport.Checked Then 'If rBtnReport.Checked Then
                    strDateNotFound = CheckAllDataExistsOrNot(SchemeCode, dtInputData, dtPlanData, strFromDate, strToDate)
                    If strDateNotFound <> "" Then
                        Dim lstItem As New ListViewItem
                        lstItem.Text = SchemeCode
                        lstItem.SubItems.Add("Input Data Not exist of the day " & strDateNotFound)
                        lstEqulizeStatus.Items.Add(lstItem)
                        'Me.Cursor = Cursors.Arrow
                        'MessageBox.Show("Input Data Not exist of the day " & strDateNotFound & " For Scheme code " & SchemeCode & " . Please Upload the data First", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Continue For
                    End If
                    If rptType = "Plan on Single Worksheet" Then
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                        xlWorkSheet.Activate()
                        If usedWrkShts > 0 Then xlWorkBook.Worksheets.Add()
                        xlWorkSheetCurrScheme = xlWorkBook.Worksheets(1)
                        If usedWrkShts > 0 Then xlWorkSheetCurrScheme.Move(, xlWorkBook.Worksheets(usedWrkShts + 1))
                        'strRptPath = strRptPath & "\" & SchemeCode
                        'NewFolderCheckOrCreate
                        xlWorkSheetCurrScheme.Name = "Scheme-" & SchemeCode
                        usedWrkShts = usedWrkShts + 1
                        'SavePath = RptFileName & "-" & SchemeCode & extc
                        SavePath = SavePath

                        'Added by Tushar as on 17082012 - START
                        GenerateRptForSingleWorksheet(usedWrkShts, dtInputData, dtPlanData, dtBtweenFromToTemplate, dtTemplateColData, SavePath, strStartEqualizationDate, strToDate)
                        boolRprGenrate = True
                        'Added by Tushar as on 17082012 - END

                    Else
                        'objBAL.NewFolderCheckOrCreate(strRptPath)
                        SavePath = RptFileName & "-" & SchemeCode & extc
                        CreateNewWorkbook(SavePath)
                        usedWrkShts = 1

                        'Added by Tushar as on 17082012 - START
                        GenerateRptForDiffWorksheet(usedWrkShts, dtInputData, dtPlanData, dtBtweenFromToTemplate, dtTemplateColData, SavePath, strStartEqualizationDate, strToDate)
                        boolRprGenrate = True
                        'Added by Tushar as on 17082012 - END

                    End If

                    'Commented by Tushar as on 17082012 - START
                    'xlApp.Visible = False
                    'GenerateRpt(usedWrkShts, dtInputData, dtPlanData, dtBtweenFromToTemplate, dtTemplateColData, SavePath, strStartEqualizationDate, strToDate)
                    'boolRprGenrate = True
                    'Commented by Tushar as on 17082012 - END
                End If
                'Added by Shweta(15 Sep 2011)
                'For Generate GL Report
                If chkGenerateGLRpt.Checked Then
                    GetGLReportData(dtBtweenFromToTemplate, strFromDate, strToDate, SchemeId)
                End If
                If chkDivDistRpt.Checked Then
                    If strEquOrRpt = "Equalization" Then
                        GetDivDistReportEqualize(dtPlanData, SchemeCode, strToDate, SchemeId)
                    Else
                        GetDivDataFromDataBase(SchemeCode, SchemeId, strToDate)
                    End If
                End If
            Next
            If chkGenerateGLRpt.Checked Then
                GenerateGLReport(strFundCode, strFromDate & " To " & strToDate)
            End If
            If strEquOrRpt = "Generate Report" And chkEquReport.Checked Then 'If rBtnReport.Checked Then
                If boolRprGenrate Then
                    If rptType = "Plan on Single Worksheet" Then
                        xlApp.DisplayAlerts = False
                        xlWorkSheetSample.Delete()
                        xlWorkBook.Save()
                        'xlWorkBook.SaveAs(strRptPath) ', Excel.XlFileFormat.xlExcel7
                        xlWorkBook.Close()
                    End If
                    objclsExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
                    Me.Cursor = Cursors.Arrow
                    MessageBox.Show("Report Generation complete", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Else
                    objclsExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
                End If
            ElseIf strEquOrRpt = "Equalization" Then
                objclsExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
                'MessageBox.Show("Equalization Completed.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            End If
            userChngData = False
            chkSelectAll.Checked = False
            'For i As Integer = 0 To dgvSchemes.Rows.Count - 1 'Commented by Pravin 16112013 start
            '    dgvSchemes.Rows(i).Cells("colChk").Value = False
            'Next 'Commented by Pravin 16112013 end
            userChngData = True
            CheckedSchemeToEqualize(True)
            'Added by shweta on (25 Aug 2011)
            'To Avoid Message box
            If lstEqulizeStatus.Items.Count > 0 Then
                If strEquOrRpt = "Equalization" Then lblEqualizeStatus.Text = "Equalization Status"
                If strEquOrRpt = "Generate Report" Then lblEqualizeStatus.Text = "Generate Report Status"
                pnlMainData.Enabled = False
                pnlMainData.SendToBack()
                pnlEqualizeStatus.BringToFront()
                pnlEqualizeStatus.Enabled = True
            End If
            '=====================================
        Catch ex As Exception
            Me.Cursor = Cursors.Arrow
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            objclsExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
        Finally
            Me.Cursor = Cursors.Arrow
            If strEquOrRpt = "Equalization" Then 'If rBtnEqualize.Checked Then
                objEqualize_clsDAL.UnLockFund(MFundId)
            End If
        End Try
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub CreateNewWorkbook(Optional ByVal SavePath As String = "")
        Try
            Dim misvalues As Object = System.Reflection.Missing.Value
            xlWorkBook = xlApp.Workbooks.Add(misvalues)
            If SavePath <> "" Then
                xlWorkBook.SaveAs(SavePath)
                xlWorkBook.Close()
                xlWorkBook = xlApp.Workbooks.Open(SavePath)
            End If
            While xlWorkBook.Worksheets.Count > 1
                xlWorkSheet = xlWorkBook.Worksheets(2)
                xlWorkSheet.Delete()
            End While
            xlWorkSheetSample = xlWorkBook.Worksheets(1)
            xlWorkSheetSample.Name = "Sample"
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub


    Private Function SaveEqualData(ByVal MFundId As Long, ByVal SchemeId As Long, ByVal FundCode As String, ByVal SchemeCode As String, ByVal StartEqFromDt As String, ByVal FromDt As String, ByVal ToDt As String, ByVal dtPlanData As DataTable, ByVal dtInputData As DataTable, Optional ByVal SchemeOpenData As String = "") As Boolean
        Dim DataUploaded As Boolean = False
        Try
            Dim FormulaStartColumn As Integer
            Dim boolInValidData As Boolean = False
            Dim dtDateData As DataTable

            Dim dataExist As Long = 0
            Dim dataExist1 As Long = 0
            Dim usedWrkShts As Integer

            Dim ColNum As Long
            Dim rownum As Long
            Dim LastDate As String
            Dim chkDt As Date
            Dim addData As Boolean
            Dim PlanCntDataExist As Long
            Dim sDate As String
            Dim PlanId As Long
            Dim InputDataStartDate As String
            Dim LogMsg As String = ""
            Dim MaxColCnt As Long = 1
            Dim strOverwrite As String = ""

            dataExist = objEqualize_clsDAL.CheckDataExist("Equalize_EqualizeData", SchemeId, StartEqFromDt, ToDt)
            dataExist1 = objEqualize_clsDAL.CheckDataExist("Equalization_Data", SchemeId, StartEqFromDt, ToDt)

            If dataExist > 0 Or dataExist1 > 0 Then
                Me.Cursor = Cursors.Arrow
                Dim t As String = vbNo
                If strYesNoFlag = "" Then
                    Dim objFrmMsgBox As New frmEqualize_MsgBox
                    objfrmEqualizeGenerateRpt.Enabled = False
                    objFrmMsgBox.lblMsg.Text = "Data already equalized for selected Period. Do you want Replace the data?"
                    objFrmMsgBox.ShowDialog()
                    t = objEqualizemsgResult
                    If t = "YesToAll" Then
                        t = vbYes
                        strYesNoFlag = "YesToAll"
                    ElseIf t = "NoToAll" Then
                        t = vbNo
                        strYesNoFlag = "NoToAll"
                    ElseIf t = "Yes" Then
                        t = vbYes
                    ElseIf t = "No" Then
                        t = vbNo
                    End If
                    objEqualizemsgResult = ""
                    objfrmEqualizeGenerateRpt.Enabled = True
                ElseIf strYesNoFlag = "YesToAll" Then
                    t = vbYes
                ElseIf strYesNoFlag = "NoToAll" Then
                    t = vbNo
                End If
                If t = vbYes Then
                    strOverwrite = "Equalization Overwrite."
                    objEqualize_clsDAL.DeleteEqualizeDataExist(SchemeId, StartEqFromDt, ToDt, ClsCommon.userName)
                End If
                If t = vbNo Then
                    Dim lstItem1 As New ListViewItem
                    lstItem1.Text = SchemeCode
                    lstItem1.SubItems.Add(LogMsg & vbCrLf & "Equalization Cancelled.")
                    lstEqulizeStatus.Items.Add(lstItem1)
                    Return False
                End If
                Me.Cursor = Cursors.WaitCursor
            End If
            dtDateData = objEqualize_clsDAL.GetApprovedData("Input_Data_Dates_From_To_To", SchemeId.ToString, FromDt, ToDt)
            usedWrkShts = 1
            xlWorkSheetSample = xlWorkBook.Worksheets("Sample")
            For index As Integer = 0 To dtPlanData.Rows.Count - 1
                FormulaStartColumn = 1
                xlWorkSheet = xlWorkBook.Worksheets(1)
                xlWorkSheet.Activate()
                xlApp.Visible = False
                If usedWrkShts > 0 Then xlWorkBook.Worksheets.Add()
                xlWorkSheet = xlWorkBook.Worksheets(1)
                If usedWrkShts > 0 Then xlWorkSheet.Move(, xlWorkBook.Worksheets(usedWrkShts + 1))
                xlWorkSheet.Name = "Plan-" & dtPlanData.Rows(index)("Plan_Code").ToString
                xlWorkSheet.Cells(1, 1).value = 2
                'For index1 As Integer = 12 To dtInputData.Columns.Count - 1        'Commented by Tushar as on 21082012
                For index1 As Integer = 12 To dtInputData.Columns.Count - 3        'Added by Tushar as on 21082012
                    xlWorkSheet.Cells(1, index1 - 10).value = dtInputData.Columns(index1).ColumnName
                Next

                'Added by Tushar as on 21082012 - START
                ColShift = 0
                ColShift = objIER
                For index1 As Integer = dtInputData.Columns.Count - 2 To dtInputData.Columns.Count - 1         'Added by Tushar as on 21082012
                    xlWorkSheet.Cells(1, ColShift).value = dtInputData.Columns(index1).ColumnName
                    ColShift = ColShift + 1
                Next
                ColShift = 0
                'Added by Tushar as on 21082012 - END

                usedWrkShts = usedWrkShts + 1
            Next
            'FormulaStartColumn = dtInputData.Columns.Count - 11     'Commented by Tushar as on 21082012
            FormulaStartColumn = dtInputData.Columns.Count - 11 - 2     'Added by Tushar as on 21082012
            MaxColCnt = FormulaStartColumn
            ColNum = 0
            rownum = 2
            LastDate = ""
            chkDt = FromDt
            addData = True
            PlanCntDataExist = 0
            Dim planCloseDate As String
            While chkDt <= Convert.ToDateTime(ToDt)
                sDate = chkDt.ToString("dd-MMM-yyyy")
                addData = True
                PlanCntDataExist = 0
                For index1 As Integer = 0 To dtPlanData.Rows.Count - 1
                    IsOpenDate = False
                    boolInValidData = False
                    PlanCode = dtPlanData.Rows(index1)("Plan_Code").ToString
                    PlanId = dtPlanData.Rows(index1)("PlanId")
                    InputDataStartDate = dtPlanData.Rows(index1)("Start_Date").ToString
                    planCloseDate = dtPlanData.Rows(index1)("ClosedDate").ToString
                    If Convert.ToDateTime(InputDataStartDate) <= Convert.ToDateTime(sDate) Then
                        'If planCloseDate <> "" Then
                        '    If IsDate(planCloseDate) Then
                        '        If Convert.ToDateTime(planCloseDate) <= Convert.ToDateTime(sDate) Then
                        '            Continue For
                        '        End If
                        '    End If
                        'End If
                        drSelect = dtInputData.Select("DataDate = '" & sDate & "' and PlanId =" & PlanId)
                        If drSelect.Length > 0 Then
                            If drSelect(0)("InputID").ToString <> "" Then
                                If drSelect(0)("InputID") <= 0 Then
                                    boolInValidData = True
                                Else
                                    xlWorkSheet = xlWorkBook.Sheets("Plan-" & dtPlanData.Rows(index1)("Plan_Code").ToString)
                                    xlWorkSheet.Activate()
                                    rownum = xlWorkSheet.Cells(1, 1).value
                                    'Copy Row Data
                                    'For index As Integer = 11 To dtInputData.Columns.Count - 1     'Commented by Tushar as on 21082012
                                    For index As Integer = 11 To dtInputData.Columns.Count - 3      'Added by Tushar as on 21082012
                                        xlWorkSheet.Cells(rownum, index - 10).value = drSelect(0)(index)
                                    Next

                                    ColShift = 0
                                    ColShift = 0
                                    ColShift = objIER
                                    For index As Integer = dtInputData.Columns.Count - 2 To dtInputData.Columns.Count - 1         'Added by Tushar as on 21082012
                                        xlWorkSheet.Cells(rownum, ColShift).value = drSelect(0)(index)
                                        ColShift = ColShift + 1
                                    Next
                                    ColShift = 0
                                    dtEqualizeData = New DataTable
                                    RptTemplateId = 0
                                    dtTemplateColData = New DataTable
                                    dtRptTemplateData = New DataTable

                                    'Get Template Data for Current Row
                                    dtRptTemplateData = objEqualize_clsDAL.GetApprovedData("Get_Template_Before_From_Date", SchemeId, sDate, , MFundId)
                                    If dtRptTemplateData.Rows.Count > 0 Then
                                        RptTemplateId = Convert.ToInt64(dtRptTemplateData.Rows(0)("AutoID").ToString)
                                        dtTemplateColData = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)
                                        TemplateEffDt = dtRptTemplateData.Rows(0)("EffectiveDate").ToString

                                        Dim copyTemplate As Boolean = True

                                        'If Selected Date Is Start Date Then Get Data From Opening Balance
                                        If IsDate(InputDataStartDate) Then
                                            If Convert.ToDateTime(sDate) = Convert.ToDateTime(InputDataStartDate) Then
                                                IsOpenDate = True
                                                ColNum = 1

                                                'Added By Shweta (14 Mar 2012)
                                                Dim ChkPlanNewStart As Boolean = False
                                                If Convert.ToDateTime(sDate) > Convert.ToDateTime(SchemeOpenDate) Then
                                                    ChkPlanNewStart = True
                                                Else
                                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Opening_Data", PlanId, RptTemplateId, sDate, SchemeId)     'Added by Tushar as on 09102012 , sDate , SchemeId
                                                    If dtEqualizeData.Rows.Count > 0 Then
                                                        For index As Integer = 0 To dtTemplateColData.Rows.Count - 1
                                                            xlWorkSheet.Cells(rownum, FormulaStartColumn + ColNum).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                            ColNum = ColNum + 1
                                                            If MaxColCnt < FormulaStartColumn + ColNum Then
                                                                MaxColCnt = FormulaStartColumn + ColNum
                                                            End If
                                                        Next
                                                        'Added By Shweta (14 Mar 2012)
                                                    Else
                                                        ChkPlanNewStart = True
                                                    End If
                                                End If

                                                If ChkPlanNewStart Then
                                                    'Dim finalColLetter As String = objclsExcel.ExcelColName(MaxColCnt - 1)    '-1 Added by Tushar as on 02092012
                                                    Dim finalColLetter As String = objclsExcel.ExcelColName(objIER + 1)    'Added by Tushar as on 04092012

                                                    xlWorkSheetSample.Cells.Clear()
                                                    xlWorkSheet.Range("A" & rownum & ":" & finalColLetter & rownum).Copy()
                                                    xlWorkSheetSample.Activate()
                                                    xlWorkSheetSample.Range("A3").Select()
                                                    xlWorkSheetSample.Paste()
                                                    'Copy formula of Template Column
                                                    ColNum = 1


                                                    ''If Added by Tushar as on 31122012 to check the opening balance    -   START
                                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Opening_Data", PlanId, RptTemplateId, sDate, SchemeId)     'Added by Tushar as on 09102012 , sDate , SchemeId
                                                    If dtEqualizeData.Rows.Count > 0 Then
                                                        For index As Integer = 0 To dtTemplateColData.Rows.Count - 1
                                                            xlWorkSheetSample.Cells(3, FormulaStartColumn + ColNum).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                            ColNum = ColNum + 1
                                                            If MaxColCnt < FormulaStartColumn + ColNum Then
                                                                MaxColCnt = FormulaStartColumn + ColNum
                                                            End If
                                                        Next

                                                    Else
                                                        For index As Integer = 0 To dtTemplateColData.Rows.Count - 1
                                                            xlWorkSheetSample.Cells(3, FormulaStartColumn + ColNum).value = dtTemplateColData.Rows(index)("Column_Formula")
                                                            ColNum = ColNum + 1
                                                            If MaxColCnt < FormulaStartColumn + ColNum Then
                                                                MaxColCnt = FormulaStartColumn + ColNum
                                                            End If
                                                        Next
                                                    End If
                                                    ''If Added by Tushar as on 31122012 to check the opening balance    -   END

                                                    '' Commented by Tushar as on 31122012   -   START
                                                    ''For index As Integer = 0 To dtTemplateColData.Rows.Count - 1
                                                    ''    xlWorkSheetSample.Cells(3, FormulaStartColumn + ColNum).value = dtTemplateColData.Rows(index)("Column_Formula")
                                                    ''    ColNum = ColNum + 1
                                                    ''    If MaxColCnt < FormulaStartColumn + ColNum Then
                                                    ''        MaxColCnt = FormulaStartColumn + ColNum
                                                    ''    End If
                                                    ''Next
                                                    '' Commented by Tushar as on 31122012   -   END

                                                    finalColLetter = objclsExcel.ExcelColName(MaxColCnt - 1)          '-1 Added by Tushar as on 02092012
                                                    xlWorkSheetSample.Range("A3:" & finalColLetter & 3).Copy()
                                                    xlWorkSheet.Activate()
                                                    'xlWorkSheet.Range("A" & rownum).Select()
                                                    xlWorkSheet.Range("A" & rownum).PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, False)
                                                    '===============================================
                                                End If
                                            End If
                                        End If

                                        If IsOpenDate = False Then
                                            'If Selected Date Is Template Effective Date Then Check If data Exist For Opening Balance
                                            If IsDate(TemplateEffDt) Then
                                                If Convert.ToDateTime(sDate) = Convert.ToDateTime(TemplateEffDt) Then
                                                    IsOpenDate = True
                                                    REM dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Opening_Data", PlanId, RptTemplateId) rem commented by Tushar as on 29082012)
                                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Opening_Data", PlanId, "", sDate, SchemeId.ToString) REM  added get the opening balance (Tushar as on 29082012)
                                                Else     'Else codition added to check the opening balance (Tushar as on 29082012)
                                                    'If rownum <> 2 Then
                                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Opening_Data", PlanId, "", sDate, SchemeId.ToString)
                                                    If dtEqualizeData.Rows.Count > 0 Then
                                                        IsOpenDate = True
                                                    End If
                                                    'End If
                                                End If
                                            End If
                                            If IsOpenDate = False Then
                                                'If Its A first row then Get Equalization Data 
                                                'Which is alraedy stored DataBase
                                                If rownum = 2 Then
                                                    copyTemplate = False
                                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, sDate)
                                                End If
                                            End If

                                            'End If
                                            'If IsOpenDate = False Then
                                            ColNum = 1
                                            If dtEqualizeData.Rows.Count > 0 Then
                                                'If data Is need to paste from Database
                                                'Then save only values
                                                For index As Integer = 0 To dtTemplateColData.Rows.Count - 1
                                                    xlWorkSheet.Cells(rownum, FormulaStartColumn + ColNum).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                    ColNum = ColNum + 1
                                                    If MaxColCnt < FormulaStartColumn + ColNum Then
                                                        MaxColCnt = FormulaStartColumn + ColNum
                                                    End If
                                                Next
                                            Else
                                                If copyTemplate Then
                                                    'Dim finalColLetter As String = objclsExcel.ExcelColName(MaxColCnt - 1)      ' -1 Added by Tushar as on 02092012
                                                    Dim finalColLetter As String = objclsExcel.ExcelColName(objIER + 1)      'Added by Tushar as on 04092012

                                                    xlWorkSheetSample.Cells.Clear()
                                                    xlWorkSheet.Range("A" & rownum - 1 & ":" & finalColLetter & rownum).Copy()
                                                    xlWorkSheetSample.Activate()
                                                    xlWorkSheetSample.Range("A2").Select()
                                                    xlWorkSheetSample.Paste()
                                                    ''Copy formula of Template Column
                                                    ColNum = 1
                                                    For index As Integer = 0 To dtTemplateColData.Rows.Count - 1
                                                        xlWorkSheetSample.Cells(3, FormulaStartColumn + ColNum).value = dtTemplateColData.Rows(index)("Column_Formula")
                                                        ColNum = ColNum + 1
                                                        If MaxColCnt < FormulaStartColumn + ColNum Then
                                                            MaxColCnt = FormulaStartColumn + ColNum
                                                        End If
                                                    Next
                                                    finalColLetter = objclsExcel.ExcelColName(MaxColCnt - 1)          '-1 Added by Tushar as on 02092012
                                                    xlWorkSheetSample.Range("A3:" & finalColLetter & 3).Copy()
                                                    xlWorkSheet.Activate()
                                                    xlWorkSheet.Range("A" & rownum).Select()
                                                    xlWorkSheet.Paste()
                                                End If
                                            End If
                                        End If
                                    End If
                                    If Convert.ToDateTime(sDate) >= Convert.ToDateTime(StartEqFromDt) Then
                                        If dtTemplateColData.Rows.Count > 0 Then
                                            'Newly added
                                            ColNum = 1
                                            strColID = ""
                                            strColValue = ""
                                            xlWorkSheet.Columns.AutoFit()
                                            For index As Integer = 0 To dtTemplateColData.Rows.Count - 1
                                                If index > 0 Then
                                                    strColID = strColID & ","
                                                    strColValue = strColValue & ","
                                                End If
                                                Val = 0
                                                If Not IsNothing(xlWorkSheet.Cells(rownum, FormulaStartColumn + ColNum).value) Then
                                                    If xlWorkSheet.Cells(rownum, FormulaStartColumn + ColNum).value.ToString <> "" Then
                                                        If IsNumeric(xlWorkSheet.Cells(rownum, FormulaStartColumn + ColNum).value.ToString) Then
                                                            Val = Convert.ToDouble(xlWorkSheet.Cells(rownum, FormulaStartColumn + ColNum).value)
                                                        Else
                                                            Val = xlWorkSheet.Cells(rownum, FormulaStartColumn + ColNum).value.ToString()
                                                        End If
                                                    End If
                                                End If
                                                strColID = strColID & dtTemplateColData.Rows(index)("AutoId")
                                                strColValue = strColValue & Val
                                                ColNum = ColNum + 1
                                            Next
                                            objEqualize_clsDAL.AddEqualizationData(MFundId, SchemeId, PlanId, RptTemplateId, ClsCommon.userName.Trim, 1, sDate, IsOpenDate, strColID, strColValue)
                                            ''---------------------------------------------
                                        End If
                                    End If
                                    xlWorkSheet.Cells(1, 1).value = rownum + 1
                                End If
                            Else
                                boolInValidData = True
                            End If
                        Else
                            boolInValidData = True
                        End If
                        If boolInValidData Then
                            addData = False
                            If chkDt.DayOfWeek <> DayOfWeek.Sunday Then
                                LogMsg = ""
                                If LastDate <> "" Then
                                    Me.Cursor = Cursors.Arrow
                                    LogMsg = "Data equalized for Fund Code :" & FundCode & " And Scheme Code : " & SchemeCode & ". From Date: " & StartEqFromDt & " and To Date :" & LastDate & " by User : " & ClsCommon.userName.Trim
                                    objEqualize_clsDAL.GenerateSysLog("Insert Log", "", LogMsg, ClsCommon.userName.Trim, "", "Generate Reports")
                                    LogMsg = "Data equalized for Fund Code :" & FundCode & " And Scheme Code : " & SchemeCode & "." & vbCrLf & "From Date: " & StartEqFromDt & " and To Date :" & LastDate
                                End If
                                Dim lstItem1 As New ListViewItem
                                lstItem1.Text = SchemeCode
                                lstItem1.SubItems.Add(strOverwrite & " " & vbCrLf & LogMsg & vbCrLf & "Input Data not exist for Date :" & Convert.ToDateTime(sDate).ToString("dd-MMM-yyyy") & "." & vbCrLf & "Failed to complete Equalization.")
                                lstEqulizeStatus.Items.Add(lstItem1)
                                'MessageBox.Show(LogMsg & vbCrLf & "Input Data not exist for Date :" & Convert.ToDateTime(sDate).ToString("dd-MMM-yyyy") & "." & vbCrLf & "Failed to complete Equalization.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                                Return False
                            End If
                        End If
                    End If
                Next
                If Convert.ToDateTime(sDate) < Convert.ToDateTime(StartEqFromDt) Then
                    addData = False
                End If
                If addData Then
                    LastDate = sDate
                    objEqualize_clsDAL.AddEqualizeData(MFundId, SchemeId, sDate, ClsCommon.userName, 0)
                End If
                chkDt = chkDt.AddDays(1)
            End While
            If LastDate <> "" Then
                Me.Cursor = Cursors.Arrow
                LogMsg = ""
                LogMsg = "Data equalized for Fund Code :" & FundCode & " And Scheme Code : " & SchemeCode & ". From Date: " & FromDt & " and To Date :" & LastDate & " by User : " & ClsCommon.userName.Trim
                objEqualize_clsDAL.GenerateSysLog("Insert Log", "", LogMsg, ClsCommon.userName.Trim, "", "Generate Reports")
            End If
            Me.Cursor = Cursors.Arrow
            Dim lstItem As New ListViewItem
            lstItem.Text = SchemeCode
            lstItem.SubItems.Add(strOverwrite & " Data Equalization completed successfully")
            lstEqulizeStatus.Items.Add(lstItem)
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
        Return DataUploaded
    End Function


    REM added by sameer on 13-Mar-13
    Private Sub Generate_DistSurplus(ByVal MFundID As Integer)
        Try
            Dim Rqry As String
            If Not Directory.Exists(TxtReportPath.Text.Trim) Then
                MessageBox.Show("Selected Path not exist. Please Select the Proper Path.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If

            If DateDiff(DateInterval.Day, DTPFromDate.Value, dtpToDate.Value) > 0 Then
                MessageBox.Show("This report is for single date plz select proper date.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If

            Dim _distSrp As DataTable
            Dim DS_str, DS_SchemeID As String
            'If _distSrp Is Nothing Then _distSrp = New DataSet
            '_distSrp = objEqualize_clsDAL.USP_DistrSurplus("COL_HEADER", MFundID)
            Rqry = "Select distinct isnull('ColumnValue' + convert(varchar,TC.ColSequenceNum),'') ColumnValue,TC.ColSequenceNum,DS.column_name,DS.Column_id,DS.OutputHeader" & vbCrLf 'Added Last Column by pundlik on 06-Jan-2013
            Rqry &= "from Equalize_Distributable_Surplus DS" & vbCrLf
            Rqry &= "left Join Equalize_TxnTemplateColumn TC on DS.Template_Column=tc.ColHeader " & vbCrLf
            Rqry &= "and TC.TemplateID in(Select max(rptTemplateid) from Equalize_EqualizationData where FundId =" & MFundID & ")" & vbCrLf
            Rqry &= "Where DS.is_approved=1 and DS.mfundID=" & MFundID & " order by DS.Column_id"

            _distSrp = objEqualize_clsDAL.FillDataSet(Rqry)

            If _distSrp Is Nothing = False Then
                If _distSrp.Rows.Count > 0 Then
                    DS_str = ""
                    ''For i As Integer = 0 To _distSrp.Rows.Count - 1
                    ''    If _distSrp.Rows(i)("ColumnValue").ToString <> "" Then
                    ''        '    DS_str &= "''"
                    ''        'Else
                    ''        DS_str &= _distSrp.Rows(i)("ColumnValue").ToString

                    ''        'Commented and Added By pundlik on 06-Jan-2014
                    ''        If _distSrp.Rows(i)("OutputHeader").ToString <> "" Then
                    ''            DS_str &= " as [" & _distSrp.Rows(i)("OutputHeader").ToString & "],"
                    ''            'Else
                    ''            '    DS_str &= " as [" & _distSrp.Rows(i)("column_name").ToString & "],"
                    ''        End If
                    ''    End If
                    ''    'DS_str &= " as [" & _distSrp.Rows(i)("column_name").ToString & "],"
                    ''Next

                    For i As Integer = 0 To _distSrp.Rows.Count - 1
                        If _distSrp.Rows(i)("OutputHeader").ToString <> "" Then
                            If _distSrp.Rows(i)("ColumnValue").ToString <> "" Then
                                DS_str &= _distSrp.Rows(i)("ColumnValue").ToString
                            Else
                                DS_str &= "''"
                            End If
                            DS_str &= " as [" & _distSrp.Rows(i)("OutputHeader").ToString & "],"
                        End If
                    Next
                    DS_str = DS_str.Substring(0, DS_str.Length - 1)
                Else
                    MsgBox("Data not available for selected combination", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                End If
            Else
                MsgBox("Data not available for selected combination", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
            End If
            DS_SchemeID = ""
            For i As Integer = 0 To dgvSchemes.Rows.Count - 1
                If dgvSchemes.Rows(i).Cells(0).Value = True Then
                    DS_SchemeID &= dgvSchemes.Rows(i).Cells(4).Value & ","
                End If
            Next
            DS_SchemeID = DS_SchemeID.Substring(0, DS_SchemeID.Length - 1)

            REM report qry

            Rqry = "Select distinct SM.SchemeCode, PM.PlanCode ,PM.PlanName " & vbCrLf
            Rqry &= "," & DS_str & vbCrLf
            Rqry &= "from Equalize_MstScheme SM" & vbCrLf
            Rqry &= "join Equalize_MstPlan PM on SM.AutoId =PM.SchemeID and SM.IsApproved =PM.IsApproved " & vbCrLf
            Rqry &= "Join Equalize_EqualizationData ED on ED.FundId =SM.MFundID and ED.SchemeId =SM.AutoId " & vbCrLf
            Rqry &= "and ED.IsApproved =SM.IsApproved AND ED.PlanId =PM.AutoId " & vbCrLf
            Rqry &= "Where SM.MFundID =" & MFundID & " and SM.AutoId in (" & DS_SchemeID & ") and SM.IsApproved=1" & vbCrLf
            Rqry &= "and ED.EqualizeDate = '" & Convert.ToDateTime(DTPFromDate.Value).ToString("dd-MMM-yyyy") & "'" & vbCrLf
            'Rqry &= "AND RptTemplateId =(Select max(rptTemplateid) from Equalize_EqualizationData where FundId =" & MFundID & ")" & vbCrLf
            Rqry &= "Order by SM.SchemeCode, PM.PlanCode"

            Dim _dsdt As DataTable
            _dsdt = objEqualize_clsDAL.FillDataSet(Rqry)
            If _dsdt Is Nothing = False Then
                If _dsdt.Rows.Count > 0 Then
                    Dim DS_path As String
                    DS_path = Create_Folder(TxtReportPath.Text.Trim, "Distributable Surplus")
                    DS_path &= "\Distributable Surplus - " & Convert.ToDateTime(DTPFromDate.Value).ToString("dd-MMM-yyyy") & ".xlsx"

                    If Dt_To_Excel(DS_path, _dsdt) = True Then
                        MsgBox("Report generated successfully at : " & DS_path, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                    Else
                        MsgBox("Report is not generated", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                    End If
                Else
                    MsgBox("Data not available for selected combination", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, Me.Text)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub
    Public Function Create_Folder(ByVal PATH As String, ByVal FOLDER_NAME As String) As String
        Try
            Dim Report_folder As String
            Report_folder = PATH & "\" & FOLDER_NAME

            Dim i As Integer = 0
            Do While Directory.Exists(Report_folder)
                i = i + 1
                Report_folder = PATH & "\" & FOLDER_NAME & "_" & i
            Loop
            Directory.CreateDirectory(Report_folder)
            Create_Folder = Report_folder
        Catch ex As Exception
            MsgBox("Error Source : " & ex.Source & vbCrLf & "Error Message : " & ex.Message & vbCrLf & "Error Occured in Method:-" & vbCrLf & ex.StackTrace, MsgBoxStyle.Information, Me.Text)
            Create_Folder = ""
        End Try
    End Function
    Public Function Dt_To_Excel(ByVal PATH As String, ByVal Data_Table As DataTable) As Boolean
        Dim obj_cls_DAL As New clsCRS_DataAccessLayer
        Dim xl_app As Excel.Application
        Dim xl_wbk As Excel.Workbook
        Dim xl_wsh As Excel.Worksheet
        Try
            Dt_To_Excel = True

            xl_app = New Excel.Application
            xl_wbk = xl_app.Workbooks.Add
            xl_app.Visible = False

            For Each xl_wsh In xl_wbk.Worksheets
                If xl_wbk.Worksheets.Count > 1 Then
                    xl_wbk.Worksheets(2).Delete()
                End If
            Next

            xl_wsh = xl_wbk.Sheets(1)
            xl_wsh.Name = "Distributable Surplus"

            Dim Obj_Arry(Data_Table.Rows.Count, Data_Table.Columns.Count - 1) As Object
            For c As Integer = 0 To Data_Table.Columns.Count - 1
                Obj_Arry(0, c) = Data_Table.Columns(c).ColumnName
            Next
            For r As Integer = 0 To Data_Table.Rows.Count - 1
                For c As Integer = 0 To Data_Table.Columns.Count - 1
                    Obj_Arry(r + 1, c) = Data_Table.Rows(r)(c)
                Next
            Next
            xl_wsh.Cells(1, 1) = lblFundName.Text : xl_wsh.Cells(2, 1) = DTPFromDate.Value.ToString("dd-MMM-yyyy")

            xl_wsh.Range("A5").Select()
            xl_wsh.Range("A5").Resize(Data_Table.Rows.Count + 1, Data_Table.Columns.Count).Value = Obj_Arry
            xl_wsh.Range("A5").Resize(Data_Table.Rows.Count + 1, Data_Table.Columns.Count).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            xl_wsh.Range("A5").Resize(Data_Table.Rows.Count + 1, Data_Table.Columns.Count).NumberFormat = "#,##0.00"
            xl_wsh.Range("A5").Resize(1, Data_Table.Columns.Count).Select()
            xl_app.Selection.Font.Bold = True
            xl_app.Selection.Interior.Color = RGB(186, 186, 186)
            xl_wsh.Columns.AutoFit()

            xl_wbk.SaveAs(PATH)

        Catch ex As Exception
            MsgBox("Error Source : " & ex.Source & vbCrLf & "Error Message : " & ex.Message & vbCrLf & "Error Occured in Method:-" & vbCrLf & ex.StackTrace, MsgBoxStyle.Information, Me.Text)
            Dt_To_Excel = False
        Finally
            xl_wbk.Close()
            xl_app.Quit()
            obj_cls_DAL.killExcelProcess()
        End Try
    End Function

    Public Sub getSchemeData()
        Try
            lblFundName.Text = ""
            If cmbMFund.SelectedIndex >= 0 Then
                userChngData = False
                strFundCode = cmbMFund.GetItemText(cmbMFund.SelectedItem)
                lblFundName.Text = cmbMFund.GetItemText(cmbMFund.SelectedValue)
                drSelect = dtFundData.Select("Mutual_Fund_Code ='" & strFundCode & "'")
                If drSelect.Length > 0 Then
                    dtSchemeData = objEqualize_clsDAL.GetApprovedData("Scheme Master", drSelect(0)("AutoId"))
                    SetDatagridView(dtSchemeData)
                    strFromDate = DTPFromDate.Text
                    strToDate = dtpToDate.Text
                    lngDateDiff = DateDiff(DateInterval.Day, Convert.ToDateTime(strFromDate), Convert.ToDateTime(strToDate))
                    If rBtnEqualize.Checked And lngDateDiff = 0 Then
                        CheckedSchemeToEqualize(True)
                    Else
                        dgvSchemes.Columns("colIsEqualized").Visible = False
                    End If
                End If
                userChngData = True
            End If
        Catch ex As Exception
            userChngData = True
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub SetDatagridView(ByVal dtData As DataTable)
        Try
            dtColtemp = New DataColumn
            dtColtemp.ColumnName = "colChk"
            dtColtemp.DataType = System.Type.GetType("System.Boolean")
            dtColtemp.DefaultValue = 0
            dtSchemeData.Columns.Add(dtColtemp)
            dtColtemp = New DataColumn
            dtColtemp.ColumnName = "colIsEqualized"
            dtColtemp.DataType = System.Type.GetType("System.Boolean")
            dtColtemp.DefaultValue = 0
            dtData.Columns.Add(dtColtemp)
            dtData.Columns("colChk").SetOrdinal(0)
            dtData.Columns("colIsEqualized").SetOrdinal(1)
            dtData.Columns("Scheme_Code").SetOrdinal(2)
            dtData.Columns("Scheme_Name").SetOrdinal(3)

            dgvSchemes.Columns("colChk").DataPropertyName = "colChk"
            dgvSchemes.Columns("colIsEqualized").DataPropertyName = "colIsEqualized"
            dgvSchemes.Columns("colSchemeCode").DataPropertyName = "Scheme_Code"
            dgvSchemes.Columns("colSchemeName").DataPropertyName = "Scheme_Name"
            dgvSchemes.Columns("colIsEqualized").ReadOnly = True
            dgvSchemes.DataSource = dtData
            For index As Integer = 2 To dgvSchemes.Columns.Count - 1
                If dgvSchemes.Columns(index).Name <> "colSchemeCode" And dgvSchemes.Columns(index).Name <> "colSchemeName" Then
                    dgvSchemes.Columns(index).Visible = False
                End If
                dgvSchemes.Columns(index).ReadOnly = True
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    ''' <summary>
    ''' To Set Checked State in DataGridView if Scheme Need to Equalize Today and If Already Scheme is already Equalize 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckedSchemeToEqualize(Optional ByVal boolGetNewData As Boolean = False)
        Try
            Dim strFundId As String = ""
            Dim schemeId As Long
            If IsFormInit Then
                strFromDate = DTPFromDate.Text
                strToDate = dtpToDate.Text
                lngDateDiff = DateDiff(DateInterval.Day, Convert.ToDateTime(strFromDate), Convert.ToDateTime(strToDate))
                If rBtnEqualize.Checked And lngDateDiff = 0 Then
                    If boolGetNewData Then
                        strFundCode = cmbMFund.GetItemText(cmbMFund.SelectedItem)
                        strFundId = ""
                        drSelect = dtFundData.Select("Mutual_Fund_Code ='" & strFundCode & "'")
                        If drSelect.Length > 0 Then
                            strFundId = drSelect(0)("AutoID").ToString
                        End If
                        dtToEqualizeScheme = objEqualize_clsDAL.GetApprovedData("GetToEqualizeData", strFundId, strFromDate)
                        dtDateData = objEqualize_clsDAL.GetApprovedData("Equalize_Data", strFundId, strFromDate)
                    End If
                    For i As Integer = 0 To dgvSchemes.Rows.Count - 1
                        schemeId = dgvSchemes.Rows(i).Cells("AutoId").Value
                        drSelect = dtDateData.Select("SchemeId = " & schemeId)
                        If drSelect.Length > 0 Then
                            dgvSchemes.Rows(i).Cells("colIsEqualized").Value = True
                        Else
                            dgvSchemes.Rows(i).Cells("colIsEqualized").Value = False
                            drSelect = dtToEqualizeScheme.Select("SchemeId = " & schemeId)
                            If drSelect.Length > 0 Then
                                dgvSchemes.Rows(i).Cells("colChk").Value = True
                            Else
                                'dgvSchemes.Rows(i).Cells("colChk").Value = False 'Commented by Pravin 16112013
                            End If
                        End If
                    Next
                    dgvSchemes.Columns("colIsEqualized").Visible = True
                Else
                    dgvSchemes.Columns("colIsEqualized").Visible = False
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Function GetLastDateToEqualize(ByVal strDate As String) As String
        Dim strLastDate As String = ""
        Dim ChkDt As Date
        Try
            ChkDt = Convert.ToDateTime(strDate).AddDays(-1)
            If ChkDt.DayOfWeek = DayOfWeek.Sunday Then
                ChkDt = ChkDt.AddDays(-1)
            End If
            strLastDate = ChkDt.ToString("dd-MMM-yyyy")
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
        Return strLastDate
    End Function


    ''' <summary>
    ''' To check all Date data is Exists or not.
    ''' </summary>
    ''' <param name="SchemeCode"></param>
    ''' <param name="dtData"></param>
    ''' <param name="dtPlanData"></param>
    ''' <param name="FromDt"></param>
    ''' <param name="ToDt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckAllDataExistsOrNot(ByVal SchemeCode As String, ByVal dtData As DataTable, ByVal dtPlanData As DataTable, ByVal FromDt As String, ByVal ToDt As String) As String
        Dim strDtNotFound As String = ""
        Try
            Dim dataExist As Long = 0
            Dim strDt As String
            Dim drSelect() As DataRow
            Dim dt As Date
            If dtPlanData.Rows.Count > 0 Then
                dt = Convert.ToDateTime(FromDt)
                While dt <= Convert.ToDateTime(ToDt)
                    If dt.DayOfWeek <> DayOfWeek.Sunday Then
                        strDt = dt.ToString("dd-MMM-yyyy")
                        drSelect = dtData.Select("DataDate ='" & strDt & "'")
                        If drSelect.Length < 1 Then
                            strDtNotFound = strDt
                            Exit While
                        End If
                    End If
                    dt = dt.AddDays(1)
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
        Return strDtNotFound
    End Function

    'Private Sub GenerateRpt(ByVal startWrkSht As Integer, ByVal dtData As DataTable, ByVal dtPlanData As DataTable, ByVal dtTemplateData As DataTable, ByVal dtTemplateColData As DataTable, ByVal strRptPath As String, ByVal strEqualizationDate As String, Optional ByVal strToDateRpt As String = "")
    '    Try
    '        Dim xlShtLastColName As String
    '        Dim xlShtLastColNum As String

    '        Dim StartColName As String
    '        Dim startColNum As Integer = 11
    '        Dim strColName As String = ""

    '        Dim RowCnt As Long = 1
    '        Dim usedWrkShts As Integer = startWrkSht
    '        Dim PlanCode As String = ""
    '        Dim PlanId As String = ""
    '        Dim MaxColCnt As Long = 1

    '        Dim drSelect() As DataRow
    '        Dim col, row As Integer

    '        Dim CurrenTtemplateID As Long
    '        Dim FormulaEndRow As Long
    '        Dim FormulaStartRow As Long
    '        Dim formulaRowCnt As Long
    '        Dim colCnt As Integer = dtData.Columns.Count - startColNum + 1
    '        MaxColCnt = dtData.Columns.Count - startColNum
    '        Dim ChkDt As String
    '        Dim AddHeader As Boolean
    '        Dim TempColCnt As Long

    '        Dim CurrentEffDt As String
    '        Dim EffectiveDt As String

    '        Dim FirstColorIndex As Boolean
    '        Dim MaxColName As String = ""
    '        Dim drSelectCol() As DataRow
    '        Dim IntFormulaStartcol As Integer
    '        Dim FormulaStartColName As String
    '        Dim CurrentColName As String = ""
    '        Dim strStartPlan As String
    '        Dim strEndPlan As String
    '        Dim TotalPlanColNum As Long
    '        Dim strTotalPlanCol As String

    '        Dim CopyStrtColNum As Integer
    '        Dim strPastColName As String
    '        Dim colNumToPaste As Long
    '        Dim stColPlanNum As Long
    '        Dim endColPlanNum As Long
    '        Dim AddNewRow As Boolean
    '        dtSampleColData = New DataTable
    '        dtSampleColData = dtTemplateColData.Copy
    '        dtSampleColData.Rows.Clear()
    '        strToDateRpt = dtpToDate.Text
    '        xlApp.Visible = False
    '        Dim PlanClosedDate As String
    '        If dtData.Rows.Count > 0 Then
    '            For index As Integer = 0 To dtPlanData.Rows.Count - 1
    '                PlanCode = dtPlanData.Rows(index)("Plan_Code").ToString
    '                PlanId = dtPlanData.Rows(index)("PlanId").ToString
    '                planStartDate = dtPlanData.Rows(index)("Start_Date").ToString
    '                PlanClosedDate = dtPlanData.Rows(index)("ClosedDate").ToString
    '                If strToDateRpt <> "" Then
    '                    If IsDate(strToDateRpt) Then
    '                        If IsDate(planStartDate) Then
    '                            If Convert.ToDateTime(planStartDate) > Convert.ToDateTime(strToDateRpt) Then
    '                                Continue For
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '                drSelect = dtData.Select("PlanId =" & PlanId)
    '                If drSelect.Length > 0 Then
    '                    xlWorkSheet = xlWorkBook.Worksheets(1)
    '                    xlWorkSheet.Activate()
    '                    If usedWrkShts > 0 Then xlWorkBook.Worksheets.Add()
    '                    xlWorkSheet = xlWorkBook.Worksheets(1)
    '                    If usedWrkShts > 0 Then xlWorkSheet.Move(, xlWorkBook.Worksheets(usedWrkShts + 1))
    '                    xlWorkSheet.Name = "Plan-" & PlanCode
    '                    xlWorkSheet.Activate()

    '                    'Paste Row  Data 
    '                    Dim rawData(drSelect.Length, dtData.Columns.Count - startColNum) As Object
    '                    Dim intlen As Integer = drSelect.Rank
    '                    ' Copy the column names to the first row of the object array
    '                    For col = 0 To dtData.Columns.Count - startColNum - 1
    '                        rawData(0, col) = dtData.Columns(col + startColNum).ColumnName.ToUpper
    '                    Next
    '                    ' Copy the values to the object array
    '                    For col = 0 To dtData.Columns.Count - startColNum - 1
    '                        'Changed by shweta (25 Jan 2012)START================
    '                        '-------To Change the Date format
    '                        If col = 1 Then
    '                            For row = 0 To drSelect.Length - 1
    '                                If IsDate(drSelect(row)(col + startColNum)) Then
    '                                    rawData(row + 1, col) = Convert.ToDateTime(drSelect(row)(col + startColNum))
    '                                Else
    '                                    rawData(row + 1, col) = drSelect(row)(col + startColNum)
    '                                End If
    '                            Next
    '                        Else
    '                            For row = 0 To drSelect.Length - 1
    '                                rawData(row + 1, col) = drSelect(row)(col + startColNum)
    '                            Next
    '                        End If
    '                        'Changed by shweta (25 Jan 2012)END================
    '                    Next

    '                    'Calculate the final column letter
    '                    Dim finalColLetter As String = String.Empty
    '                    Dim finalColLetter1 As String = String.Empty

    '                    finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum)
    '                    Dim excelRange As String = String.Format("A" & RowCnt & ":{0}{1}", finalColLetter, drSelect.Length + RowCnt)
    '                    xlWorkSheet.Range(excelRange, Type.Missing).NumberFormat = "@"
    '                    xlWorkSheet.Range(excelRange, Type.Missing).Select()
    '                    xlWorkSheet.Range(excelRange, Type.Missing).Value2 = rawData
    '                    xlWorkSheet.Columns.AutoFit()

    '                    'Added by Tushar as on 06082012 - START
    '                    'If rptType = "Plan on Different Worksheet" Then
    '                    '    finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum)
    '                    '    finalColLetter1 = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum - 1)
    '                    '    xlWorkSheet.Range(finalColLetter1 & 1 & ":" & finalColLetter & drSelect.Length + RowCnt).Cut()
    '                    '    xlWorkSheet.Columns.Cells(1, objIER).select()
    '                    '    xlWorkSheet.Paste()
    '                    'End If
    '                    'Added by Tushar as on 06082012 - END

    '                    xlWorkSheet.Range("1:1").Font.Bold = True
    '                    'Changed by shweta (25 Jan 2012)START================
    '                    '-------To Change the Date format
    '                    xlWorkSheet.Range("C:" & finalColLetter).Cells.NumberFormat = "#,##0.00"
    '                    xlWorkSheet.Range("B:B").NumberFormat = "dd-mmm-yyyy"
    '                    'Changed by shweta (25 Jan 2012)END================
    '                    'xlApp.Visible = True

    '                    'Commented by Tushar as on 06082012 - START
    '                    'IntFormulaStartcol = dtData.Columns.Count - startColNum
    '                    'Commented by Tushar as on 06082012 - END

    '                    'Added by Tushar as on 06082012 - START
    '                    'IntFormulaStartcol = dtData.Columns.Count - startColNum - 2    'Commented by Tushar as on 16082012 
    '                    IntFormulaStartcol = dtData.Columns.Count - startColNum         'Added by Tushar as on 06082012
    '                    'Added by Tushar as on 06082012 - END
    '                    FormulaStartColName = objclsExcel.ExcelColName(IntFormulaStartcol)
    '                    MaxColName = ""

    '                    'xlApp.Visible = True
    '                    'Dim CurrenTtemplateID As Long
    '                    FormulaEndRow = 3
    '                    FormulaStartRow = 3
    '                    'Commented by Tushar as on 06082012 - START
    '                    'colCnt = dtData.Columns.Count - startColNum + 1
    '                    'MaxColCnt = dtData.Columns.Count - startColNum
    '                    'Commented by Tushar as on 06082012 - END

    '                    'Added by Tushar as on 06082012 - START
    '                    'colCnt = dtData.Columns.Count - startColNum - 1         'Commented by Tushar as on 16082012 
    '                    'MaxColCnt = dtData.Columns.Count - startColNum - 2      'Commented by Tushar as on 16082012 

    '                    colCnt = dtData.Columns.Count - startColNum + 1         'Added by Tushar as on 16082012 
    '                    MaxColCnt = dtData.Columns.Count - startColNum         'Added by Tushar as on 16082012 
    '                    'Added by Tushar as on 06082012 - END

    '                    AddHeader = True
    '                    TempColCnt = 0

    '                    StartColName = objclsExcel.ExcelColName(colCnt)
    '                    xlShtLastColNum = colCnt
    '                    'Get previous Day  Data
    '                    Dim ColNum As Integer = 1
    '                    'ChkDt = xlWorkSheet.Cells(2, 2).value

    '                    'If IsDate(planStartDate) Then
    '                    '    If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
    '                    '        dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, ChkDt)

    '                    '        If dtEqualizeData.Rows.Count > 0 Then
    '                    '            RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
    '                    '            dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)
    '                    '        End If

    '                    '        For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
    '                    '            xlWorkSheet.Cells(2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
    '                    '            ColNum = ColNum + 1
    '                    '            xlShtLastColNum = i + colCnt
    '                    '        Next
    '                    '    Else
    '                    '        dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, Convert.ToDateTime(planStartDate).ToString("dd-MMM-yyyy"))

    '                    '        If dtEqualizeData.Rows.Count > 0 Then
    '                    '            RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
    '                    '            dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)
    '                    '        End If
    '                    '        For index2 As Integer = 2 To drSelect.Length + 1
    '                    '            ChkDt = xlWorkSheet.Cells(index2, 2).value
    '                    '            If Convert.ToDateTime(ChkDt) = Convert.ToDateTime(planStartDate) Then
    '                    '                For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
    '                    '                    xlWorkSheet.Cells(index2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
    '                    '                    ColNum = ColNum + 1
    '                    '                    xlShtLastColNum = i + colCnt
    '                    '                Next
    '                    '                Exit For
    '                    '            End If
    '                    '        Next
    '                    '    End If
    '                    'End If
    '                    'xlApp.Visible = True

    '                    'If drSelect.Length = 1 Then
    '                    '    FormulaStartRow = 3
    '                    'End If


    '                    For index1 As Integer = 0 To dtTemplateData.Rows.Count - 1
    '                        formulaRowCnt = 2
    '                        CurrentEffDt = dtTemplateData.Rows(index1)("EffectiveDate").ToString
    '                        If index1 < dtTemplateData.Rows.Count - 1 Then
    '                            EffectiveDt = dtTemplateData.Rows(index1 + 1)("EffectiveDate").ToString
    '                            ChkDt = xlWorkSheet.Cells(FormulaStartRow, 2).Text
    '                            If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(EffectiveDt) Then
    '                                Continue For
    '                            End If
    '                            For index2 As Integer = FormulaStartRow To drSelect.Length + RowCnt
    '                                ChkDt = xlWorkSheet.Cells(index2, 2).Text
    '                                If Convert.ToDateTime(ChkDt) < Convert.ToDateTime(EffectiveDt) Then
    '                                    FormulaEndRow = index2
    '                                    formulaRowCnt = formulaRowCnt + 1
    '                                Else
    '                                    Exit For
    '                                End If
    '                            Next
    '                        Else
    '                            FormulaEndRow = drSelect.Length + RowCnt
    '                            formulaRowCnt = FormulaEndRow - FormulaStartRow + 3
    '                        End If

    '                        CurrenTtemplateID = dtTemplateData.Rows(index1)("AutoID")
    '                        drSelectCol = dtTemplateColData.Select("TemplateID ='" & CurrenTtemplateID & "'", "ColSequenceNum")
    '                        xlWorkSheetSample.Cells.Clear()
    '                        'Commented by Tushar as on 06082012 - START
    '                        'xlWorkSheet.Range("A" & FormulaStartRow - 1 & ":" & finalColLetter & FormulaEndRow).Copy()
    '                        'Commented by Tushar as on 06082012 - END

    '                        'Added by Tushar as on 06082012 - START
    '                        'finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum - 2)  'Commented by Tushar as on 16082012 
    '                        xlWorkSheet.Range("A" & FormulaStartRow - 1 & ":" & finalColLetter & FormulaEndRow).Copy()
    '                        'Added by Tushar as on 06082012 - END

    '                        xlWorkSheetSample.Activate()
    '                        xlWorkSheetSample.Range("A2").Select()
    '                        xlWorkSheetSample.Paste()

    '                        xlShtLastColName = objclsExcel.ExcelColName(xlShtLastColNum)
    '                        xlWorkSheet.Range(StartColName & FormulaStartRow - 1 & ":" & xlShtLastColName & FormulaStartRow - 1).Copy()
    '                        xlWorkSheetSample.Activate()
    '                        xlWorkSheetSample.Range(StartColName & 2).PasteSpecial(Excel.XlPasteType.xlPasteValues)

    '                        'xlApp.Visible = True
    '                        For i As Integer = 0 To drSelectCol.Length - 1
    '                            If AddHeader Then
    '                                xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
    '                            Else
    '                                If IsNothing(xlWorkSheet.Cells(1, i + colCnt).value) Then
    '                                    xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
    '                                ElseIf xlWorkSheet.Cells(1, i + colCnt).value.ToString = "" Then
    '                                    xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
    '                                End If
    '                            End If
    '                            If drSelect.Length > 1 Then
    '                                xlWorkSheetSample.Cells(3, i + colCnt).Value = drSelectCol(i)("ColFormula").ToString
    '                                strColName = objclsExcel.ExcelColName(i + colCnt)
    '                                If formulaRowCnt > 3 Then

    '                                    xlWorkSheetSample.Range(strColName & 3).AutoFill(xlWorkSheetSample.Range(strColName & 3 & ":" & strColName & (formulaRowCnt)), Excel.XlAutoFillType.xlFillDefault)
    '                                End If
    '                            End If
    '                            'to Show Decimal number
    '                            xlRange = strColName & 3 & ":" & strColName & (formulaRowCnt)
    '                            Dim Val As Integer = 0
    '                            If drSelectCol(i)("ColDecimalNum").ToString() <> "" Then
    '                                Val = drSelectCol(i)("ColDecimalNum")
    '                            End If
    '                            objclsExcel.SetNumberFormatToColumn(Val, xlRange, xlWorkSheetSample)

    '                            If dtSampleColData.Rows.Count - 1 < i Then
    '                                MaxColCnt = MaxColCnt + 1
    '                                drAddRow = dtSampleColData.NewRow
    '                                drAddRow("ColHeader") = drSelectCol(i)("ColHeader").ToString
    '                                drAddRow("ColFormula") = drSelectCol(i)("ColFormula").ToString
    '                                drAddRow("ColShowFormula") = drSelectCol(i)("ColShowFormula").ToString
    '                                drAddRow("ColIsSchemeWise") = drSelectCol(i)("ColIsSchemeWise").ToString
    '                                drAddRow("ColShowTotal") = drSelectCol(i)("ColShowTotal")
    '                                drAddRow("ColDecimalNum") = drSelectCol(i)("ColDecimalNum")
    '                                dtSampleColData.Rows.Add(drAddRow)
    '                                xlShtLastColNum = xlShtLastColNum + 1
    '                            Else
    '                                'Commented by Tushar as on 06082012 - START
    '                                'MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count
    '                                'Commented by Tushar as on 06082012 - END
    '                                'Added by Tushar as on 0608212 - START
    '                                'MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count - 2    'Commented by Tushar as on 16082012                                    
    '                                MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count     'Added by Tushar as on 16082012
    '                                'Added by Tushar as on 0608212 - END
    '                            End If
    '                        Next
    '                        If drSelect.Length > 1 Then
    '                            xlWorkSheetSample.Range("A3:" & strColName & (formulaRowCnt)).Copy()
    '                            xlWorkSheet.Activate()
    '                            xlWorkSheet.Range("A" & FormulaStartRow).Select()
    '                            xlWorkSheet.Paste()
    '                        End If
    '                        If AddHeader Then
    '                            CopyStrtColNum = FormulaStartRow - 1
    '                        Else
    '                            CopyStrtColNum = FormulaStartRow
    '                        End If

    '                        ChkDt = xlWorkSheet.Cells(CopyStrtColNum, 2).Text
    '                        ColNum = 1
    '                        If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
    '                            dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, ChkDt)

    '                            'If dtEqualizeData.Rows.Count > 0 Then
    '                            '    RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
    '                            '    dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)

    '                            '    For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
    '                            '        xlWorkSheet.Cells(CopyStrtColNum, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
    '                            '        ColNum = ColNum + 1
    '                            '        xlShtLastColNum = i + colCnt
    '                            '    Next
    '                            'End If

    '                            If dtEqualizeData.Rows.Count > 0 Then
    '                                For i As Integer = 0 To 50
    '                                    If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
    '                                        If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
    '                                            xlWorkSheet.Cells(CopyStrtColNum, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
    '                                            ColNum = ColNum + 1
    '                                            xlShtLastColNum = i + colCnt
    '                                        End If
    '                                    End If
    '                                Next
    '                            End If
    '                        Else
    '                            'Get Last date of selecte Region
    '                            ChkDt = xlWorkSheet.Cells(FormulaStartRow + formulaRowCnt - 3, 2).Text.ToString
    '                            If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
    '                                dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, Convert.ToDateTime(planStartDate).ToString("dd-MMM-yyyy"))

    '                                'If dtEqualizeData.Rows.Count > 0 Then
    '                                '    RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
    '                                '    dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)
    '                                'End If
    '                                'For index2 As Integer = CopyStrtColNum To CopyStrtColNum + formulaRowCnt - 3
    '                                '    ChkDt = xlWorkSheet.Cells(index2, 2).value
    '                                '    If Convert.ToDateTime(ChkDt) = Convert.ToDateTime(planStartDate) Then
    '                                '        For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
    '                                '            xlWorkSheet.Cells(index2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
    '                                '            ColNum = ColNum + 1
    '                                '            xlShtLastColNum = i + colCnt
    '                                '        Next
    '                                '        Exit For
    '                                '    End If
    '                                'Next

    '                                'Changed on 18 Aug 2011
    '                                For index2 As Integer = CopyStrtColNum To CopyStrtColNum + formulaRowCnt - 3
    '                                    ChkDt = xlWorkSheet.Cells(index2, 2).Text
    '                                    If Convert.ToDateTime(ChkDt) = Convert.ToDateTime(planStartDate) Then
    '                                        If dtEqualizeData.Rows.Count > 0 Then
    '                                            For i As Integer = 0 To 50
    '                                                If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
    '                                                    If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
    '                                                        xlWorkSheet.Cells(index2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
    '                                                        ColNum = ColNum + 1
    '                                                        xlShtLastColNum = i + colCnt
    '                                                    End If
    '                                                End If
    '                                            Next
    '                                        End If
    '                                    End If
    '                                Next
    '                                '=====================================
    '                            End If
    '                        End If


    '                        'Added by Tushar as on 16082012 - START
    '                        If rptType = "Plan on Different Worksheet" Then
    '                            finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum)
    '                            finalColLetter1 = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum - 1)
    '                            xlWorkSheet.Range(finalColLetter1 & 1 & ":" & finalColLetter & drSelect.Length + RowCnt).Cut()
    '                            xlWorkSheet.Columns.Cells(1, objIER).select()
    '                            'xlWorkSheet.Selection.Insert(Shift:=Excel.xlToRight)
    '                            xlWorkSheet.Application.Selection.insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight)
    '                            'xlWorkSheet.Paste()
    '                        End If
    '                        'Added by Tushar as on 16082012 - END


    '                        If Not AddHeader Then xlWorkSheet.Range(FormulaStartRow & ":" & FormulaStartRow).Interior.ColorIndex = 47.6
    '                        FormulaStartRow = FormulaEndRow + 1
    '                        AddHeader = False
    '                    Next

    '                    If rptType = "Plan on Different Worksheet" Then
    '                        MaxColName = objclsExcel.ExcelColName(MaxColCnt - 2)        'Added by Tushar as on 16082012
    '                        FormatPlanWiseExcelSheet(xlWorkSheet, MaxColCnt - 2, drSelect.Length + 1, IntFormulaStartcol, strEqualizationDate)   'Added by Tushar as on 16082012
    '                    Else
    '                        MaxColName = objclsExcel.ExcelColName(MaxColCnt)
    '                        FormatPlanWiseExcelSheet(xlWorkSheet, MaxColCnt, drSelect.Length + 1, IntFormulaStartcol, strEqualizationDate)   'Added by Tushar as on 16082012
    '                    End If

    '                    ''FormatPlanWiseExcelSheet(xlWorkSheet, MaxColCnt, drSelect.Length + 1, IntFormulaStartcol, strEqualizationDate)    'Commented by Tushar as on 16082012
    '                    usedWrkShts = usedWrkShts + 1
    '                End If
    '            Next

    '            If rptType = "Plan on Different Worksheet" Then
    '                xlApp.DisplayAlerts = False
    '                xlWorkSheetSample.Delete()
    '                xlWorkBook.Save()
    '                xlWorkBook.Close()
    '            Else
    '                'If All Plan in one sheet
    '                If usedWrkShts > 1 Then
    '                    xlApp.Visible = False
    '                    xlWorkSheetCurrScheme.Cells.Clear()
    '                    xlWorkSheetCurrScheme.Activate()
    '                    xlWorkSheet.Range("A:B").Copy()
    '                    xlWorkSheetCurrScheme.Range("A1").Select()
    '                    xlWorkSheetCurrScheme.Paste()
    '                    xlWorkSheetCurrScheme.Range("1:1").Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

    '                    Dim LastRowCnt As Long
    '                    LastRowCnt = xlWorkSheetCurrScheme.Cells.Find(What:="*", After:=xlWorkSheetCurrScheme.Cells(2, 1), _
    '                                             SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '                                             SearchDirection:=Excel.XlSearchDirection.xlPrevious).Row

    '                    FirstColorIndex = True

    '                    CurrentColName = ""
    '                    colNumToPaste = 3
    '                    stColPlanNum = 3
    '                    endColPlanNum = 3
    '                    AddNewRow = True
    '                    CopyStrtColNum = IntFormulaStartcol

    '                    Dim StrColName2 As String = ""
    '                    Dim ColMainNum As Integer = 3
    '                    Dim wrkShtNum As Integer = 0

    '                    Dim addStartColName As String
    '                    Dim AddEndColName As String

    '                    Dim addStartColNameMain As String
    '                    Dim AddEndColNameMain As String

    '                    Dim StartPlanNum As Integer = 3
    '                    Dim EndPlanNum As Integer = 3
    '                    Dim Num As Long
    '                    Dim strNum As String
    '                    Dim colInputData As Long = 0
    '                    For index As Integer = startWrkSht + 1 To usedWrkShts
    '                        xlWorkSheet = xlWorkBook.Worksheets(index)
    '                        ColMainNum = 3
    '                        CopyStrtColNum = 3
    '                        colNumToPaste = 3 + wrkShtNum
    '                        xlWorkSheet.Range("1:1").Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

    '                        'Copy Scheme Total 
    '                        If wrkShtNum = 0 Then
    '                            strPastColName = objclsExcel.ExcelColName(colNumToPaste)
    '                            Num = 2 + objEqualizeTotalSchemeInputCol
    '                            strNum = objclsExcel.ExcelColName(Num)
    '                            xlWorkSheet.Range("C:" & strNum).Copy()
    '                            xlWorkSheetCurrScheme.Activate()
    '                            xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
    '                            xlWorkSheetCurrScheme.Paste()
    '                            xlWorkSheetCurrScheme.Cells(1, 3).value = xlWorkSheetCurrScheme.Name
    '                            xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, "C1:" & strNum & "1", True)
    '                        End If
    '                        '=========================================
    '                        CopyStrtColNum = objEqualizeTotalSchemeInputCol + 3 '11
    '                        ColMainNum = objEqualizeTotalSchemeInputCol + 3 '11
    '                        colNumToPaste = objEqualizeTotalSchemeInputCol + 3 '11
    '                        If wrkShtNum > 0 Then
    '                            colNumToPaste = colNumToPaste + (objEqualizeTotalInputCol * wrkShtNum)
    '                            addStartColName = objclsExcel.ExcelColName(colNumToPaste)
    '                            addStartColNameMain = objclsExcel.ExcelColName(colNumToPaste + (objEqualizeTotalInputCol - 1))
    '                            xlWorkSheetCurrScheme.Range(addStartColName & ":" & addStartColNameMain).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

    '                            addStartColName = objclsExcel.ExcelColName(ColMainNum)
    '                            addStartColNameMain = objclsExcel.ExcelColName(ColMainNum - 1 + (objEqualizeTotalInputCol * wrkShtNum))
    '                            xlWorkSheet.Range(addStartColName & ":" & addStartColNameMain).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

    '                            CopyStrtColNum = ColMainNum - 1 + (objEqualizeTotalInputCol * wrkShtNum) + 1
    '                            ColMainNum = CopyStrtColNum
    '                            colNumToPaste = CopyStrtColNum
    '                        End If

    '                        'To Copy INPUT DATA
    '                        strPastColName = objclsExcel.ExcelColName(colNumToPaste)
    '                        strColName = objclsExcel.ExcelColName(ColMainNum)
    '                        CopyStrtColNum = ColMainNum + objEqualizeTotalInputCol - 1
    '                        StrColName2 = objclsExcel.ExcelColName(CopyStrtColNum)

    '                        xlWorkSheet.Range(strColName & ":" & StrColName2).Copy()
    '                        xlWorkSheetCurrScheme.Activate()
    '                        xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
    '                        xlWorkSheetCurrScheme.Paste()
    '                        xlWorkSheetCurrScheme.Cells(1, colNumToPaste).value = xlWorkSheet.Name
    '                        xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strColName & "1:" & StrColName2 & "1", True)
    '                        '==============================================

    '                        CopyStrtColNum = CopyStrtColNum + 1
    '                        ColMainNum = CopyStrtColNum
    '                        colNumToPaste = colNumToPaste + objEqualizeTotalInputCol + (wrkShtNum)
    '                        Dim oldColNumToPaste As Long = colNumToPaste
    '                        For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
    '                            Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
    '                            If wrkShtNum > 0 And (val = "" Or val = "TRUE" Or val = "1") Then
    '                                ColMainNum = ColMainNum + 1
    '                                colNumToPaste = colNumToPaste + 1
    '                                Continue For
    '                            End If
    '                            If wrkShtNum > 0 Then
    '                                addStartColName = objclsExcel.ExcelColName(ColMainNum)
    '                                addStartColNameMain = objclsExcel.ExcelColName(ColMainNum + 1)
    '                                ColMainNum = ColMainNum + wrkShtNum
    '                                AddEndColName = objclsExcel.ExcelColName(ColMainNum - 1)
    '                                AddEndColNameMain = objclsExcel.ExcelColName(ColMainNum)
    '                                xlWorkSheet.Range(addStartColName & ":" & AddEndColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
    '                                strPastColName = objclsExcel.ExcelColName(colNumToPaste)
    '                                xlWorkSheetCurrScheme.Range(strPastColName & ":" & strPastColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
    '                            End If

    '                            'strPastColName = objExcel.ExcelColName(colNumToPaste)
    '                            'strEndPlan = strPastColName

    '                            'strColName = objExcel.ExcelColName(ColMainNum)

    '                            ''Changed
    '                            'xlWorkSheet.Range(strColName & ":" & strColName).Copy()
    '                            'xlWorkSheetCurrScheme.Activate()
    '                            'xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
    '                            'xlWorkSheetCurrScheme.Paste()
    '                            'xlWorkSheetCurrScheme.Cells(2, colNumToPaste).Value = xlWorkSheet.Name.ToString
    '                            'xlWorkSheetCurrScheme.Cells(1, colNumToPaste).Value = xlWorkSheet.Cells(2, ColMainNum).value

    '                            ColMainNum = ColMainNum + 1
    '                            colNumToPaste = colNumToPaste + wrkShtNum + 1
    '                        Next

    '                        'CopyStrtColNum = CopyStrtColNum + 1
    '                        ColMainNum = CopyStrtColNum
    '                        colNumToPaste = oldColNumToPaste 'colNumToPaste + objEqualizeTotalInputCol + (wrkShtNum)
    '                        For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
    '                            Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
    '                            If wrkShtNum > 0 And (val = "" Or val = "TRUE" Or val = "1") Then
    '                                ColMainNum = ColMainNum + 1
    '                                colNumToPaste = colNumToPaste + 1
    '                                Continue For
    '                            End If

    '                            If wrkShtNum > 0 Then
    '                                '    addStartColName = objExcel.ExcelColName(ColMainNum)
    '                                '    addStartColNameMain = objExcel.ExcelColName(ColMainNum + 1)
    '                                ColMainNum = ColMainNum + wrkShtNum
    '                                '    AddEndColName = objExcel.ExcelColName(ColMainNum - 1)
    '                                '    AddEndColNameMain = objExcel.ExcelColName(ColMainNum)
    '                                '    xlWorkSheet.Range(addStartColName & ":" & AddEndColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
    '                                strPastColName = objclsExcel.ExcelColName(colNumToPaste)
    '                                '    xlWorkSheetCurrScheme.Range(strPastColName & ":" & strPastColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
    '                            End If

    '                            strPastColName = objclsExcel.ExcelColName(colNumToPaste)
    '                            strEndPlan = strPastColName

    '                            strColName = objclsExcel.ExcelColName(ColMainNum)

    '                            'Changed
    '                            xlWorkSheet.Range(strColName & ":" & strColName).Copy()
    '                            xlWorkSheetCurrScheme.Activate()
    '                            xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
    '                            xlWorkSheetCurrScheme.Paste()
    '                            xlWorkSheetCurrScheme.Cells(2, colNumToPaste).Value = xlWorkSheet.Name.ToString
    '                            xlWorkSheetCurrScheme.Cells(1, colNumToPaste).Value = xlWorkSheet.Cells(2, ColMainNum).value

    '                            ColMainNum = ColMainNum + 1
    '                            colNumToPaste = colNumToPaste + wrkShtNum + 1
    '                        Next

    '                        wrkShtNum = wrkShtNum + 1
    '                    Next

    '                    ''Added By Shweta'
    '                    ''To avoid ms calculation due to formula contains data for next column
    '                    'wrkShtNum = 0

    '                    'For index As Integer = startWrkSht + 1 To usedWrkShts
    '                    '    xlWorkSheet = xlWorkBook.Worksheets(index)
    '                    '    ColMainNum = 3
    '                    '    CopyStrtColNum = 3
    '                    '    colNumToPaste = 3 + wrkShtNum

    '                    '    CopyStrtColNum = objEqualizeTotalSchemeInputCol + 3 '11
    '                    '    ColMainNum = objEqualizeTotalSchemeInputCol + 3 '11
    '                    '    colNumToPaste = objEqualizeTotalSchemeInputCol + 3 '11

    '                    '    If wrkShtNum > 0 Then
    '                    '        colNumToPaste = colNumToPaste + (objEqualizeTotalInputCol * wrkShtNum)
    '                    '        CopyStrtColNum = ColMainNum - 1 + (objEqualizeTotalInputCol * wrkShtNum) + 1
    '                    '        ColMainNum = CopyStrtColNum
    '                    '        colNumToPaste = CopyStrtColNum
    '                    '    End If


    '                    '    CopyStrtColNum = CopyStrtColNum + 1
    '                    '    ColMainNum = CopyStrtColNum
    '                    '    colNumToPaste = colNumToPaste + objEqualizeTotalInputCol + (wrkShtNum)

    '                    '    For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
    '                    '        Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
    '                    '        If wrkShtNum > 0 And (val = "" Or val = "TRUE" Or val = "1") Then
    '                    '            ColMainNum = ColMainNum + 1
    '                    '            colNumToPaste = colNumToPaste + 1
    '                    '            Continue For
    '                    '        End If

    '                    '        If wrkShtNum > 0 Then
    '                    '            addStartColName = objExcel.ExcelColName(ColMainNum)
    '                    '            addStartColNameMain = objExcel.ExcelColName(ColMainNum + 1)

    '                    '            ColMainNum = ColMainNum + wrkShtNum

    '                    '            AddEndColName = objExcel.ExcelColName(ColMainNum - 1)
    '                    '            AddEndColNameMain = objExcel.ExcelColName(ColMainNum)

    '                    '            'xlWorkSheet.Range(addStartColName & ":" & AddEndColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

    '                    '            strPastColName = objExcel.ExcelColName(colNumToPaste)
    '                    '            'xlWorkSheetCurrScheme.Range(strPastColName & ":" & strPastColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
    '                    '        End If

    '                    '        strPastColName = objExcel.ExcelColName(colNumToPaste)
    '                    '        strEndPlan = strPastColName

    '                    '        strColName = objExcel.ExcelColName(ColMainNum)

    '                    '        'Changed
    '                    '        xlWorkSheet.Range(strColName & ":" & strColName).Copy()
    '                    '        xlWorkSheetCurrScheme.Activate()
    '                    '        xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
    '                    '        xlWorkSheetCurrScheme.Paste()
    '                    '        xlWorkSheetCurrScheme.Cells(2, colNumToPaste).Value = xlWorkSheet.Name.ToString
    '                    '        xlWorkSheetCurrScheme.Cells(1, colNumToPaste).Value = xlWorkSheet.Cells(2, ColMainNum).value

    '                    '        ColMainNum = ColMainNum + 1
    '                    '        colNumToPaste = colNumToPaste + wrkShtNum + 1
    '                    '    Next

    '                    '    wrkShtNum = wrkShtNum + 1

    '                    'Next



    '                    'Changed By Shweta(21 Jun 2012)
    '                    ' StartPlanNum = 11 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol) + 1
    '                    StartPlanNum = objEqualizeTotalInputCol + 1 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol) + 1
    '                    '===========================================
    '                    'StartPlanNum = 11 + ((usedWrkShts - startWrkSht) * 8)
    '                    EndPlanNum = StartPlanNum
    '                    For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
    '                        Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
    '                        If val = "" Or val = "TRUE" Or val = "1" Then
    '                            EndPlanNum = StartPlanNum
    '                        Else
    '                            EndPlanNum = StartPlanNum + (usedWrkShts - startWrkSht - 1)
    '                        End If

    '                        strStartPlan = objclsExcel.ExcelColName(StartPlanNum)
    '                        strEndPlan = objclsExcel.ExcelColName(EndPlanNum)


    '                        'To add Total column In Worksheet 
    '                        If val <> "" And val <> "TRUE" And val <> "1" Then
    '                            val = dtSampleColData.Rows(index1)("ColShowTotal").ToString.Trim.ToUpper()
    '                            If val = "" Or val = "TRUE" Or val = "1" Then
    '                                TotalPlanColNum = EndPlanNum + 1
    '                                strTotalPlanCol = objclsExcel.ExcelColName(TotalPlanColNum)
    '                                xlWorkSheetCurrScheme.Range(strTotalPlanCol & ":" & strTotalPlanCol).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
    '                                xlWorkSheetCurrScheme.Cells(2, strTotalPlanCol).Value = "Total"
    '                                xlWorkSheetCurrScheme.Cells(2, strTotalPlanCol).Font.Bold = True
    '                                xlWorkSheetCurrScheme.Cells(3, strTotalPlanCol).Value = "=SUM(" & strStartPlan & "3:" & strEndPlan & "3)" '=SUM(C3:C3)
    '                                xlWorkSheetCurrScheme.Range(strTotalPlanCol & 3).AutoFill(xlWorkSheetCurrScheme.Range(strTotalPlanCol & "3:" & strTotalPlanCol & LastRowCnt), Excel.XlAutoFillType.xlFillDefault)
    '                                xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strTotalPlanCol & "2:" & strTotalPlanCol & LastRowCnt, False)

    '                                strEndPlan = strTotalPlanCol
    '                                EndPlanNum = TotalPlanColNum
    '                            End If
    '                        End If
    '                        '--------------------------------    

    '                        If FirstColorIndex Then
    '                            xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & LastRowCnt).Interior.ColorIndex = 35
    '                            FirstColorIndex = False
    '                        Else
    '                            xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & LastRowCnt).Interior.ColorIndex = 36
    '                            FirstColorIndex = True
    '                        End If

    '                        xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & "1").Select()
    '                        xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strStartPlan & "1:" & strEndPlan & "1", True)
    '                        StartPlanNum = EndPlanNum + 1
    '                    Next


    '                    xlWorkSheetCurrScheme.Columns.AutoFit()
    '                    'To hide the Row Data
    '                    StartPlanNum = 3
    '                    'Changed By Shweta(21 Jun 2012)
    '                    'EndPlanNum = 11 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol)
    '                    EndPlanNum = objEqualizeTotalInputCol + 1 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol)
    '                    '======================================================
    '                    'strStartPlan = objExcel.ExcelColName(StartPlanNum)
    '                    strEndPlan = objclsExcel.ExcelColName(EndPlanNum)
    '                    xlWorkSheetCurrScheme.Range("C:" & strEndPlan).EntireColumn.Hidden = True

    '                    For index As Integer = startWrkSht + 1 To usedWrkShts
    '                        xlWorkSheet = xlWorkBook.Worksheets(startWrkSht + 1)
    '                        xlWorkSheet.Delete()
    '                    Next
    '                End If

    '            End If
    '        End If



    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '    End Try
    'End Sub


    Private Sub GenerateRptForDiffWorksheet(ByVal startWrkSht As Integer, ByVal dtData As DataTable, ByVal dtPlanData As DataTable, ByVal dtTemplateData As DataTable, ByVal dtTemplateColData As DataTable, ByVal strRptPath As String, ByVal strEqualizationDate As String, Optional ByVal strToDateRpt As String = "")
        Try
            Dim xlShtLastColName As String
            Dim xlShtLastColNum As String

            Dim StartColName As String
            Dim startColNum As Integer = 11
            Dim strColName As String = ""

            Dim RowCnt As Long = 1
            Dim usedWrkShts As Integer = startWrkSht
            Dim PlanCode As String = ""
            Dim PlanId As String = ""
            Dim MaxColCnt As Long = 1

            Dim drSelect() As DataRow
            Dim col, row As Integer

            Dim CurrenTtemplateID As Long
            Dim FormulaEndRow As Long
            Dim FormulaStartRow As Long
            Dim formulaRowCnt As Long
            Dim colCnt As Integer = dtData.Columns.Count - startColNum + 1
            MaxColCnt = dtData.Columns.Count - startColNum
            Dim ChkDt As String
            Dim ChkDt1 As String ' Added by Tushar as on 03012013
            Dim AddHeader As Boolean
            Dim TempColCnt As Long

            Dim CurrentEffDt As String
            Dim EffectiveDt As String

            Dim FirstColorIndex As Boolean
            Dim MaxColName As String = ""
            Dim drSelectCol() As DataRow
            Dim IntFormulaStartcol As Integer
            Dim FormulaStartColName As String
            Dim CurrentColName As String = ""
            Dim strStartPlan As String
            Dim strEndPlan As String
            Dim TotalPlanColNum As Long
            Dim strTotalPlanCol As String

            Dim CopyStrtColNum As Integer
            Dim strPastColName As String
            Dim colNumToPaste As Long
            Dim stColPlanNum As Long
            Dim endColPlanNum As Long
            Dim AddNewRow As Boolean
            dtSampleColData = New DataTable
            dtSampleColData = dtTemplateColData.Copy
            dtSampleColData.Rows.Clear()
            strToDateRpt = dtpToDate.Text
            xlApp.Visible = False
            Dim PlanClosedDate As String
            If dtData.Rows.Count > 0 Then
                For index As Integer = 0 To dtPlanData.Rows.Count - 1
                    PlanCode = dtPlanData.Rows(index)("Plan_Code").ToString
                    PlanId = dtPlanData.Rows(index)("PlanId").ToString
                    planStartDate = dtPlanData.Rows(index)("Start_Date").ToString
                    PlanClosedDate = dtPlanData.Rows(index)("ClosedDate").ToString
                    If strToDateRpt <> "" Then
                        If IsDate(strToDateRpt) Then
                            If IsDate(planStartDate) Then
                                If Convert.ToDateTime(planStartDate) > Convert.ToDateTime(strToDateRpt) Then
                                    Continue For
                                End If
                            End If
                        End If
                    End If
                    drSelect = dtData.Select("PlanId =" & PlanId)
                    If drSelect.Length > 0 Then
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                        xlWorkSheet.Activate()
                        If usedWrkShts > 0 Then xlWorkBook.Worksheets.Add()
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                        If usedWrkShts > 0 Then xlWorkSheet.Move(, xlWorkBook.Worksheets(usedWrkShts + 1))
                        xlWorkSheet.Name = "Plan-" & PlanCode
                        xlWorkSheet.Activate()

                        'Paste Row  Data 
                        Dim rawData(drSelect.Length, dtData.Columns.Count - startColNum) As Object
                        Dim intlen As Integer = drSelect.Rank
                        ' Copy the column names to the first row of the object array
                        For col = 0 To dtData.Columns.Count - startColNum - 1
                            rawData(0, col) = dtData.Columns(col + startColNum).ColumnName.ToUpper
                        Next
                        ' Copy the values to the object array
                        For col = 0 To dtData.Columns.Count - startColNum - 1
                            'Changed by shweta (25 Jan 2012)START================
                            '-------To Change the Date format
                            If col = 1 Then
                                For row = 0 To drSelect.Length - 1
                                    If IsDate(drSelect(row)(col + startColNum)) Then
                                        rawData(row + 1, col) = Convert.ToDateTime(drSelect(row)(col + startColNum))
                                    Else
                                        rawData(row + 1, col) = drSelect(row)(col + startColNum)
                                    End If
                                Next
                            Else
                                For row = 0 To drSelect.Length - 1
                                    rawData(row + 1, col) = drSelect(row)(col + startColNum)
                                Next
                            End If
                            'Changed by shweta (25 Jan 2012)END================
                        Next

                        'Calculate the final column letter
                        Dim finalColLetter As String = String.Empty
                        Dim finalColLetter1 As String = String.Empty
                        finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum)
                        Dim excelRange As String = String.Format("A" & RowCnt & ":{0}{1}", finalColLetter, drSelect.Length + RowCnt)
                        xlWorkSheet.Range(excelRange, Type.Missing).NumberFormat = "@"
                        xlWorkSheet.Range(excelRange, Type.Missing).Select()
                        xlWorkSheet.Range(excelRange, Type.Missing).Value2 = rawData
                        xlWorkSheet.Columns.AutoFit()

                        'Added by Tushar as on 06082012 - START
                        finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum)
                        finalColLetter1 = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum - 1)
                        xlWorkSheet.Range(finalColLetter1 & 1 & ":" & finalColLetter & drSelect.Length + RowCnt).Cut()
                        xlWorkSheet.Columns.Cells(1, objIER).select()
                        xlWorkSheet.Paste()

                        'Added by Tushar as on 06082012 - END

                        xlWorkSheet.Range("1:1").Font.Bold = True
                        'Changed by shweta (25 Jan 2012)START================
                        '-------To Change the Date format
                        xlWorkSheet.Range("C:" & finalColLetter).Cells.NumberFormat = "#,##0.00"
                        xlWorkSheet.Range("B:B").NumberFormat = "dd-mmm-yyyy"
                        'Changed by shweta (25 Jan 2012)END================
                        'xlApp.Visible = True

                        'Commented by Tushar as on 06082012 - START
                        'IntFormulaStartcol = dtData.Columns.Count - startColNum
                        'Commented by Tushar as on 06082012 - END

                        'Added by Tushar as on 06082012 - START
                        IntFormulaStartcol = dtData.Columns.Count - startColNum - 2    'Commented by Tushar as on 16082012 
                        'IntFormulaStartcol = dtData.Columns.Count - startColNum         'Added by Tushar as on 06082012
                        'Added by Tushar as on 06082012 - END
                        FormulaStartColName = objclsExcel.ExcelColName(IntFormulaStartcol)
                        MaxColName = ""

                        'xlApp.Visible = True
                        'Dim CurrenTtemplateID As Long
                        FormulaEndRow = 3
                        FormulaStartRow = 3
                        'Commented by Tushar as on 06082012 - START
                        'colCnt = dtData.Columns.Count - startColNum + 1
                        'MaxColCnt = dtData.Columns.Count - startColNum
                        'Commented by Tushar as on 06082012 - END

                        'Added by Tushar as on 06082012 - START
                        'colCnt = dtData.Columns.Count - startColNum - 1         'Commented by Tushar as on 16082012 
                        'MaxColCnt = dtData.Columns.Count - startColNum - 2      'Commented by Tushar as on 16082012 

                        colCnt = dtData.Columns.Count - startColNum - 1         'Added by Tushar as on 16082012 
                        MaxColCnt = dtData.Columns.Count - startColNum - 2        'Added by Tushar as on 16082012 
                        'Added by Tushar as on 06082012 - END

                        AddHeader = True
                        TempColCnt = 0

                        StartColName = objclsExcel.ExcelColName(colCnt)
                        xlShtLastColNum = colCnt
                        'Get previous Day  Data
                        Dim ColNum As Integer = 1

                        For index1 As Integer = 0 To dtTemplateData.Rows.Count - 1
                            formulaRowCnt = 2
                            CurrentEffDt = dtTemplateData.Rows(index1)("EffectiveDate").ToString
                            If index1 < dtTemplateData.Rows.Count - 1 Then
                                EffectiveDt = dtTemplateData.Rows(index1 + 1)("EffectiveDate").ToString
                                ChkDt = xlWorkSheet.Cells(FormulaStartRow, 2).Text

                                If Convert.ToDateTime(ChkDt) > Convert.ToDateTime(EffectiveDt) Then        'Replace  >= with > 28122012 by Tushar 
                                    Continue For
                                End If
                                For index2 As Integer = FormulaStartRow To drSelect.Length + RowCnt
                                    ChkDt = xlWorkSheet.Cells(index2, 2).Text
                                    If Convert.ToDateTime(ChkDt) < Convert.ToDateTime(EffectiveDt) Then
                                        FormulaEndRow = index2
                                        formulaRowCnt = formulaRowCnt + 1
                                    ElseIf Convert.ToDateTime(xlWorkSheet.Cells(FormulaStartRow, 2).Text) = Convert.ToDateTime(EffectiveDt) Then  'Else if added by Tushar as on 28122012
                                        If FormulaStartRow = 3 Then
                                            FormulaEndRow = index2 - 1
                                            formulaRowCnt = formulaRowCnt - 1
                                        Else
                                            FormulaEndRow = index2
                                            formulaRowCnt = formulaRowCnt + 1

                                        End If
                                        'FormulaEndRow = index2
                                        'formulaRowCnt = formulaRowCnt + 1
                                        Exit For
                                    Else
                                        Exit For
                                    End If
                                Next
                            Else
                                FormulaEndRow = drSelect.Length + RowCnt
                                formulaRowCnt = FormulaEndRow - FormulaStartRow + 3
                            End If

                            CurrenTtemplateID = dtTemplateData.Rows(index1)("AutoID")
                            drSelectCol = dtTemplateColData.Select("TemplateID ='" & CurrenTtemplateID & "'", "ColSequenceNum")
                            xlWorkSheetSample.Cells.Clear()
                            'Added by Tushar as on 06082012 - START
                            'xlWorkSheet.Range("A" & FormulaStartRow - 1 & ":" & finalColLetter & FormulaEndRow).Copy()
                            'Added by Tushar as on 06082012 - END

                            'Added by Tushar as on 06082012 - START
                            finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum - 2)  'Added by Tushar as on 16082012 
                            xlWorkSheet.Range("A" & FormulaStartRow - 1 & ":" & finalColLetter & FormulaEndRow).Copy()
                            'Added by Tushar as on 06082012 - END

                            xlWorkSheetSample.Activate()
                            xlWorkSheetSample.Range("A2").Select()
                            xlWorkSheetSample.Paste()

                            'Added by Tushar as on 29122012  - START
                            xlWorkSheet.Activate()
                            finalColLetter = objclsExcel.ExcelColName(objIER)  'Added by Tushar as on 16082012 
                            finalColLetter1 = objclsExcel.ExcelColName(objIER + 1)
                            xlWorkSheet.Range(finalColLetter & FormulaStartRow - 1 & ":" & finalColLetter1 & FormulaEndRow).Copy()

                            xlWorkSheetSample.Activate()
                            xlWorkSheetSample.Range(finalColLetter & 2).Select()
                            xlWorkSheetSample.Paste()
                            'Added by Tushar as on 29122012  - END  

                            xlShtLastColName = objclsExcel.ExcelColName(xlShtLastColNum)
                            xlWorkSheet.Range(StartColName & FormulaStartRow - 1 & ":" & xlShtLastColName & FormulaStartRow - 1).Copy()
                            xlWorkSheetSample.Activate()
                            xlWorkSheetSample.Range(StartColName & 2).PasteSpecial(Excel.XlPasteType.xlPasteValues)

                            'xlApp.Visible = True
                            For i As Integer = 0 To drSelectCol.Length - 1
                                If AddHeader Then
                                    xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                Else
                                    If IsNothing(xlWorkSheet.Cells(1, i + colCnt).value) Then
                                        xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                    ElseIf xlWorkSheet.Cells(1, i + colCnt).value.ToString = "" Then
                                        xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                    End If
                                End If
                                If drSelect.Length > 1 Then
                                    strColName = objclsExcel.ExcelColName(i + colCnt)
                                    If (xlWorkSheetSample.Cells(3, 1).text().ToString().Trim()) <> "" Then  'IF condition added by Tushar as 03012012
                                        xlWorkSheetSample.Cells(3, i + colCnt).Value = drSelectCol(i)("ColFormula").ToString

                                        If formulaRowCnt > 3 Then
                                            xlWorkSheetSample.Range(strColName & 3).AutoFill(xlWorkSheetSample.Range(strColName & 3 & ":" & strColName & (formulaRowCnt)), Excel.XlAutoFillType.xlFillDefault)
                                        End If
                                    End If

                                End If


                                ''Commented by Tushar as on 03022013
                                'To show formula or not
                                'If Condition added by Tushar as on 15122012    -   START
                                ''If UCase(drSelectCol(i)("ColShowFormula").ToString) = "FALSE" Then
                                ''    xlWorkSheetSample.Range(strColName & 3 & ":" & strColName & formulaRowCnt).Copy()
                                ''    xlWorkSheetSample.Range(strColName & 3).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                ''End If
                                'If Condition added by Tushar as on 15122012    -   END


                                'to Show Decimal number
                                strColName = objclsExcel.ExcelColName(i + colCnt)  ''Added by Tushar as on 04012013
                                If formulaRowCnt = 1 And FormulaStartRow = 3 Then
                                    xlRange = strColName & 2 & ":" & strColName & (formulaRowCnt)
                                Else
                                    xlRange = strColName & 3 & ":" & strColName & (formulaRowCnt)
                                End If
                                xlRange = strColName & 3 & ":" & strColName & (formulaRowCnt)
                                Dim Val As Integer = 0
                                If drSelectCol(i)("ColDecimalNum").ToString() <> "" Then
                                    Val = drSelectCol(i)("ColDecimalNum")
                                End If
                                objclsExcel.SetNumberFormatToColumn(Val, xlRange, xlWorkSheetSample)

                                If dtSampleColData.Rows.Count - 1 < i Then
                                    MaxColCnt = MaxColCnt + 1
                                    drAddRow = dtSampleColData.NewRow
                                    drAddRow("ColHeader") = drSelectCol(i)("ColHeader").ToString
                                    drAddRow("ColFormula") = drSelectCol(i)("ColFormula").ToString
                                    drAddRow("ColShowFormula") = drSelectCol(i)("ColShowFormula").ToString
                                    drAddRow("ColIsSchemeWise") = drSelectCol(i)("ColIsSchemeWise").ToString
                                    drAddRow("ColShowTotal") = drSelectCol(i)("ColShowTotal")
                                    drAddRow("ColDecimalNum") = drSelectCol(i)("ColDecimalNum")
                                    dtSampleColData.Rows.Add(drAddRow)
                                    xlShtLastColNum = xlShtLastColNum + 1
                                Else
                                    'Added by Tushar as on 06082012 - START
                                    'MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count
                                    'Added by Tushar as on 06082012 - END
                                    'Added by Tushar as on 0608212 - START
                                    MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count - 2    'Commented by Tushar as on 16082012                                    
                                    'MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count     'Added by Tushar as on 16082012
                                    'Added by Tushar as on 0608212 - END
                                End If
                            Next
                            If drSelect.Length > 1 Then

                                If formulaRowCnt = 1 And FormulaStartRow = 3 Then
                                    xlWorkSheetSample.Range("A2:" & strColName & "2").Copy()
                                    xlWorkSheet.Activate()
                                    xlWorkSheet.Range("A" & 2).Select()
                                    xlWorkSheet.Paste()
                                Else
                                    xlWorkSheetSample.Range("A3:" & strColName & (formulaRowCnt)).Copy()
                                    xlWorkSheet.Activate()
                                    xlWorkSheet.Range("A" & FormulaStartRow).Select()
                                    xlWorkSheet.Paste()
                                End If


                            End If
                            If AddHeader Then
                                CopyStrtColNum = FormulaStartRow - 1
                            Else
                                CopyStrtColNum = FormulaStartRow
                            End If

                            ChkDt = xlWorkSheet.Cells(CopyStrtColNum, 2).Text
                            ColNum = 1

                            If (Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate)) Then

                                If (Convert.ToDateTime(ChkDt) = Convert.ToDateTime(CurrentEffDt) Or AddHeader = True) Then        ''IF conditio added by Tushar as on 29122012
                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, ChkDt)
                                    If dtEqualizeData.Rows.Count > 0 Then
                                        For i As Integer = 0 To 50
                                            If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
                                                If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
                                                    'Need to check tushar as on 18082012
                                                    xlWorkSheet.Cells(CopyStrtColNum, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                    ColNum = ColNum + 1
                                                    xlShtLastColNum = i + colCnt
                                                End If
                                            End If
                                        Next
                                    End If
                                End If

                                'Added by Tushar as on 28122012 - start
                                ColNum = 1
                                If (xlWorkSheetSample.Cells(FormulaStartRow, 2).text().ToString().Trim()) <> "" Then  'Added by Tushar as on 04012013
                                    If (Convert.ToDateTime(xlWorkSheet.Cells(FormulaStartRow, 2).Text) = Convert.ToDateTime(EffectiveDt)) Then
                                        AddHeader = False
                                        ChkDt1 = xlWorkSheet.Cells(FormulaStartRow, 2).Text    'Added by Tushar as on 03012013
                                        dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, ChkDt1)

                                        If dtEqualizeData.Rows.Count > 0 Then
                                            For i As Integer = 0 To 50
                                                If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
                                                    If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
                                                        xlWorkSheet.Cells(FormulaStartRow, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                        ColNum = ColNum + 1
                                                        xlShtLastColNum = i + colCnt
                                                    End If
                                                End If
                                            Next
                                        End If

                                    End If
                                End If

                                'Added by Tushar as on 28122012 - end

                            Else
                                'Get Last date of selecte Region

                                ChkDt = xlWorkSheet.Cells(FormulaStartRow + formulaRowCnt - 3, 2).Text.ToString

                                If IsDate(ChkDt) Then   'IF condition added by Tushar as on 03012013
                                    If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
                                        dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, Convert.ToDateTime(planStartDate).ToString("dd-MMM-yyyy"))
                                        'Changed on 18 Aug 2011
                                        For index2 As Integer = CopyStrtColNum To CopyStrtColNum + formulaRowCnt - 3
                                            ChkDt = xlWorkSheet.Cells(index2, 2).Text
                                            If Convert.ToDateTime(ChkDt) = Convert.ToDateTime(planStartDate) Then
                                                If dtEqualizeData.Rows.Count > 0 Then
                                                    For i As Integer = 0 To 50
                                                        If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
                                                            If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
                                                                xlWorkSheet.Cells(index2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                                ColNum = ColNum + 1
                                                                xlShtLastColNum = i + colCnt
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            End If

                                            'Added by Tushar as on 28122012 - start
                                            ColNum = 1
                                            If (xlWorkSheetSample.Cells(FormulaStartRow, 2).text().ToString().Trim()) <> "" Then  'Added by Tushar as on 04012013
                                                If (Convert.ToDateTime(xlWorkSheet.Cells(FormulaStartRow, 2).Text) = Convert.ToDateTime(EffectiveDt)) Then
                                                    AddHeader = False
                                                    'Added by Tushar as on 03012013 Start
                                                    ChkDt1 = ""
                                                    ChkDt1 = xlWorkSheet.Cells(FormulaStartRow, 2).Text
                                                    'Added by Tushar as on 03012013 End
                                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, ChkDt1)    'Replace Convert.ToDateTime(EffectiveDt) with chkDt1

                                                    If dtEqualizeData.Rows.Count > 0 Then
                                                        For i As Integer = 0 To 50
                                                            If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
                                                                If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
                                                                    xlWorkSheet.Cells(FormulaStartRow, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                                    ColNum = ColNum + 1
                                                                    xlShtLastColNum = i + colCnt
                                                                End If
                                                            End If
                                                        Next
                                                    End If

                                                End If
                                            End If

                                            'Added by Tushar as on 28122012 - end
                                        Next

                                    End If

                                    '=====================================
                                End If
                            End If

                            ''If Not AddHeader Then xlWorkSheet.Range(FormulaStartRow & ":" & FormulaStartRow).Interior.ColorIndex = 47.6       'Commented by Tushar as on 28122012
                            If (xlWorkSheet.Cells(FormulaStartRow, 2).text().ToString().Trim()) <> "" Then  'Added by Tushar as on 04012013
                                If Not AddHeader And ((Convert.ToDateTime(xlWorkSheet.Cells(FormulaStartRow, 2).Text) = Convert.ToDateTime(EffectiveDt)) Or (Convert.ToDateTime(xlWorkSheet.Cells(FormulaStartRow, 2).Text) = Convert.ToDateTime(CurrentEffDt))) Then
                                    xlWorkSheet.Range(FormulaStartRow & ":" & FormulaStartRow).Interior.ColorIndex = 47.6
                                End If
                            End If

                            FormulaStartRow = FormulaEndRow + 1
                            AddHeader = False
                            ChkDt1 = "" 'Added by Tushar as on 03012013
                        Next

                        MaxColName = objclsExcel.ExcelColName(MaxColCnt)        'Added by Tushar as on 16082012
                        FormatPlanWiseExcelSheet(xlWorkSheet, MaxColCnt, drSelect.Length + 1, IntFormulaStartcol, strEqualizationDate)   'Added by Tushar as on 16082012
                        'Added by Tushar as on 02092012
                        finalColLetter = objclsExcel.ExcelColName(objIER - 1)   'Added by Tushar as on 02092012
                        xlWorkSheet.Range(finalColLetter & ":" & finalColLetter).Columns.Insert(Excel.XlDirection.xlToRight)

                        'Added by Tushar as on 02092012
                        usedWrkShts = usedWrkShts + 1
                    End If
                Next
            End If

            REM added by Sameer on 16-Jan-2013
            Excel_Paste_Special(xlWorkBook)

            xlApp.DisplayAlerts = False
            xlWorkSheetSample.Delete()
            xlWorkBook.Save()

        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    REM added by Sameer on 16-Jan-13
    Private Sub Excel_Paste_Special(ByVal xl_wkb As Excel.Workbook)
        Try
            If chk_WithFormula.Checked Then Exit Sub

            Dim xl_sht As Excel.Worksheet
            For Each xl_sht In xl_wkb.Sheets
                xl_sht.Select() : xl_sht.UsedRange.Copy() : xl_sht.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteValues)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub GenerateRptForSingleWorksheet(ByVal startWrkSht As Integer, ByVal dtData As DataTable, ByVal dtPlanData As DataTable, ByVal dtTemplateData As DataTable, ByVal dtTemplateColData As DataTable, ByVal strRptPath As String, ByVal strEqualizationDate As String, Optional ByVal strToDateRpt As String = "")
        Try
            Dim xlShtLastColName As String
            Dim xlShtLastColNum As String

            Dim StartColName As String
            Dim startColNum As Integer = 11
            Dim strColName As String = ""

            Dim RowCnt As Long = 1
            Dim usedWrkShts As Integer = startWrkSht
            Dim PlanCode As String = ""
            Dim PlanId As String = ""
            Dim MaxColCnt As Long = 1

            Dim drSelect() As DataRow
            Dim col, row As Integer

            Dim CurrenTtemplateID As Long
            Dim FormulaEndRow As Long
            Dim FormulaStartRow As Long
            Dim formulaRowCnt As Long
            Dim colCnt As Integer = dtData.Columns.Count - startColNum + 1
            MaxColCnt = dtData.Columns.Count - startColNum
            Dim ChkDt As String
            Dim AddHeader As Boolean
            Dim TempColCnt As Long

            Dim CurrentEffDt As String
            Dim EffectiveDt As String

            Dim FirstColorIndex As Boolean
            Dim MaxColName As String = ""
            Dim drSelectCol() As DataRow
            Dim IntFormulaStartcol As Integer
            Dim FormulaStartColName As String
            Dim CurrentColName As String = ""
            Dim strStartPlan As String
            Dim strEndPlan As String
            Dim TotalPlanColNum As Long
            Dim strTotalPlanCol As String

            Dim CopyStrtColNum As Integer
            Dim strPastColName As String
            Dim colNumToPaste As Long
            Dim stColPlanNum As Long
            Dim endColPlanNum As Long
            Dim AddNewRow As Boolean
            dtSampleColData = New DataTable
            dtSampleColData = dtTemplateColData.Copy
            dtSampleColData.Rows.Clear()
            strToDateRpt = dtpToDate.Text
            xlApp.Visible = False
            Dim PlanClosedDate As String
            If dtData.Rows.Count > 0 Then
                For index As Integer = 0 To dtPlanData.Rows.Count - 1
                    PlanCode = dtPlanData.Rows(index)("Plan_Code").ToString
                    PlanId = dtPlanData.Rows(index)("PlanId").ToString
                    planStartDate = dtPlanData.Rows(index)("Start_Date").ToString
                    PlanClosedDate = dtPlanData.Rows(index)("ClosedDate").ToString
                    If strToDateRpt <> "" Then
                        If IsDate(strToDateRpt) Then
                            If IsDate(planStartDate) Then
                                If Convert.ToDateTime(planStartDate) > Convert.ToDateTime(strToDateRpt) Then
                                    Continue For
                                End If
                            End If
                        End If
                    End If
                    drSelect = dtData.Select("PlanId =" & PlanId)
                    If drSelect.Length > 0 Then
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                        xlWorkSheet.Activate()
                        If usedWrkShts > 0 Then xlWorkBook.Worksheets.Add()
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                        If usedWrkShts > 0 Then xlWorkSheet.Move(, xlWorkBook.Worksheets(usedWrkShts + 1))
                        xlWorkSheet.Name = "Plan-" & PlanCode
                        xlWorkSheet.Activate()

                        'Paste Row  Data 
                        Dim rawData(drSelect.Length, dtData.Columns.Count - startColNum) As Object
                        Dim intlen As Integer = drSelect.Rank
                        ' Copy the column names to the first row of the object array
                        'For col = 0 To dtData.Columns.Count - startColNum - 1  'Commented bz Tushar as on 17082012
                        For col = 0 To dtData.Columns.Count - startColNum - 3
                            rawData(0, col) = dtData.Columns(col + startColNum).ColumnName.ToUpper
                        Next
                        ' Copy the values to the object array
                        'For col = 0 To dtData.Columns.Count - startColNum - 1
                        For col = 0 To dtData.Columns.Count - startColNum - 3
                            'Changed by shweta (25 Jan 2012)START================
                            '-------To Change the Date format
                            If col = 1 Then
                                For row = 0 To drSelect.Length - 1
                                    If IsDate(drSelect(row)(col + startColNum)) Then
                                        rawData(row + 1, col) = Convert.ToDateTime(drSelect(row)(col + startColNum))
                                    Else
                                        rawData(row + 1, col) = drSelect(row)(col + startColNum)
                                    End If
                                Next
                            Else
                                For row = 0 To drSelect.Length - 1
                                    rawData(row + 1, col) = drSelect(row)(col + startColNum)
                                Next
                            End If
                            'Changed by shweta (25 Jan 2012)END================
                        Next

                        'Calculate the final column letter
                        Dim finalColLetter As String = String.Empty
                        Dim finalColLetter1 As String = String.Empty

                        'finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum)  'Commented by Tushar as on 170802012
                        finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum - 2)   'Added by Tushar as on 170802012
                        Dim excelRange As String = String.Format("A" & RowCnt & ":{0}{1}", finalColLetter, drSelect.Length + RowCnt)
                        xlWorkSheet.Range(excelRange, Type.Missing).NumberFormat = "@"
                        xlWorkSheet.Range(excelRange, Type.Missing).Select()
                        xlWorkSheet.Range(excelRange, Type.Missing).Value2 = rawData
                        xlWorkSheet.Columns.AutoFit()
                        xlApp.Visible = False ' True
                        xlWorkSheet.Range("1:1").Font.Bold = True
                        'Changed by shweta (25 Jan 2012)START================
                        '-------To Change the Date format
                        xlWorkSheet.Range("C:" & finalColLetter).Cells.NumberFormat = "#,##0.00"
                        xlWorkSheet.Range("B:B").NumberFormat = "dd-mmm-yyyy"
                        'Changed by shweta (25 Jan 2012)END================
                        'xlApp.Visible = True

                        'Commented by Tushar as on 06082012 - START
                        'IntFormulaStartcol = dtData.Columns.Count - startColNum
                        'Commented by Tushar as on 06082012 - END

                        'Added by Tushar as on 06082012 - START
                        IntFormulaStartcol = dtData.Columns.Count - startColNum - 2    'Added  by Tushar as on 16082012 
                        'IntFormulaStartcol = dtData.Columns.Count - startColNum         'Commented by Tushar as on 06082012
                        'Added by Tushar as on 06082012 - END
                        FormulaStartColName = objclsExcel.ExcelColName(IntFormulaStartcol)
                        MaxColName = ""

                        'xlApp.Visible = True
                        'Dim CurrenTtemplateID As Long
                        FormulaEndRow = 3
                        FormulaStartRow = 3
                        'Commented by Tushar as on 06082012 - START
                        'colCnt = dtData.Columns.Count - startColNum + 1
                        'MaxColCnt = dtData.Columns.Count - startColNum
                        'Commented by Tushar as on 06082012 - END

                        'Added by Tushar as on 06082012 - START
                        colCnt = dtData.Columns.Count - startColNum - 1         'Added by Tushar as on 16082012 
                        MaxColCnt = dtData.Columns.Count - startColNum - 2      'Added by Tushar as on 16082012 

                        'colCnt = dtData.Columns.Count - startColNum + 1         'Added by Tushar as on 16082012 
                        'MaxColCnt = dtData.Columns.Count - startColNum         'Added by Tushar as on 16082012 
                        'Added by Tushar as on 06082012 - END

                        AddHeader = True
                        TempColCnt = 0

                        StartColName = objclsExcel.ExcelColName(colCnt)
                        xlShtLastColNum = colCnt
                        'Get previous Day  Data
                        Dim ColNum As Integer = 1

                        For index1 As Integer = 0 To dtTemplateData.Rows.Count - 1
                            formulaRowCnt = 2
                            CurrentEffDt = dtTemplateData.Rows(index1)("EffectiveDate").ToString
                            If index1 < dtTemplateData.Rows.Count - 1 Then
                                EffectiveDt = dtTemplateData.Rows(index1 + 1)("EffectiveDate").ToString
                                ChkDt = xlWorkSheet.Cells(FormulaStartRow, 2).Text
                                If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(EffectiveDt) Then
                                    Continue For
                                End If
                                For index2 As Integer = FormulaStartRow To drSelect.Length + RowCnt
                                    ChkDt = xlWorkSheet.Cells(index2, 2).Text
                                    If Convert.ToDateTime(ChkDt) < Convert.ToDateTime(EffectiveDt) Then
                                        FormulaEndRow = index2
                                        formulaRowCnt = formulaRowCnt + 1
                                    Else
                                        Exit For
                                    End If
                                Next
                            Else
                                FormulaEndRow = drSelect.Length + RowCnt
                                formulaRowCnt = FormulaEndRow - FormulaStartRow + 3
                            End If

                            CurrenTtemplateID = dtTemplateData.Rows(index1)("AutoID")
                            drSelectCol = dtTemplateColData.Select("TemplateID ='" & CurrenTtemplateID & "'", "ColSequenceNum")
                            xlWorkSheetSample.Cells.Clear()
                            'Commented by Tushar as on 06082012 - START
                            'xlWorkSheet.Range("A" & FormulaStartRow - 1 & ":" & finalColLetter & FormulaEndRow).Copy()
                            'Commented by Tushar as on 06082012 - END

                            'Added by Tushar as on 06082012 - START
                            'finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum - 2)  'Commented by Tushar as on 16082012 
                            xlWorkSheet.Range("A" & FormulaStartRow - 1 & ":" & finalColLetter & FormulaEndRow).Copy()
                            'Added by Tushar as on 06082012 - END

                            xlWorkSheetSample.Activate()
                            xlWorkSheetSample.Range("A2").Select()
                            xlWorkSheetSample.Paste()

                            xlShtLastColName = objclsExcel.ExcelColName(xlShtLastColNum)
                            xlWorkSheet.Range(StartColName & FormulaStartRow - 1 & ":" & xlShtLastColName & FormulaStartRow - 1).Copy()
                            xlWorkSheetSample.Activate()
                            xlWorkSheetSample.Range(StartColName & 2).PasteSpecial(Excel.XlPasteType.xlPasteValues)

                            'xlApp.Visible = True
                            For i As Integer = 0 To drSelectCol.Length - 1
                                If AddHeader Then
                                    xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                Else
                                    If IsNothing(xlWorkSheet.Cells(1, i + colCnt).value) Then
                                        xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                    ElseIf xlWorkSheet.Cells(1, i + colCnt).value.ToString = "" Then
                                        xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                    End If
                                End If
                                If drSelect.Length > 1 Then
                                    xlWorkSheetSample.Cells(3, i + colCnt).Value = drSelectCol(i)("ColFormula").ToString
                                    strColName = objclsExcel.ExcelColName(i + colCnt)
                                    If formulaRowCnt > 3 Then

                                        xlWorkSheetSample.Range(strColName & 3).AutoFill(xlWorkSheetSample.Range(strColName & 3 & ":" & strColName & (formulaRowCnt)), Excel.XlAutoFillType.xlFillDefault)
                                    End If
                                End If
                                'to Show Decimal number
                                xlRange = strColName & 3 & ":" & strColName & (formulaRowCnt)
                                Dim Val As Integer = 0
                                If drSelectCol(i)("ColDecimalNum").ToString() <> "" Then
                                    Val = drSelectCol(i)("ColDecimalNum")
                                End If
                                objclsExcel.SetNumberFormatToColumn(Val, xlRange, xlWorkSheetSample)

                                If dtSampleColData.Rows.Count - 1 < i Then
                                    MaxColCnt = MaxColCnt + 1
                                    drAddRow = dtSampleColData.NewRow
                                    drAddRow("ColHeader") = drSelectCol(i)("ColHeader").ToString
                                    drAddRow("ColFormula") = drSelectCol(i)("ColFormula").ToString
                                    drAddRow("ColShowFormula") = drSelectCol(i)("ColShowFormula").ToString
                                    drAddRow("ColIsSchemeWise") = drSelectCol(i)("ColIsSchemeWise").ToString
                                    drAddRow("ColShowTotal") = drSelectCol(i)("ColShowTotal")
                                    drAddRow("ColDecimalNum") = drSelectCol(i)("ColDecimalNum")
                                    dtSampleColData.Rows.Add(drAddRow)
                                    xlShtLastColNum = xlShtLastColNum + 1
                                Else
                                    'Commented by Tushar as on 06082012 - START
                                    'MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count
                                    'Commented by Tushar as on 06082012 - END
                                    'Added by Tushar as on 0608212 - START
                                    'MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count - 2    'Commented by Tushar as on 16082012                                    
                                    MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count - 2    'Added by Tushar as on 16082012
                                    'Added by Tushar as on 0608212 - END
                                End If
                            Next
                            If drSelect.Length > 1 Then
                                xlWorkSheetSample.Range("A3:" & strColName & (formulaRowCnt)).Copy()
                                xlWorkSheet.Activate()
                                xlWorkSheet.Range("A" & FormulaStartRow).Select()
                                xlWorkSheet.Paste()
                            End If
                            If AddHeader Then
                                CopyStrtColNum = FormulaStartRow - 1
                            Else
                                CopyStrtColNum = FormulaStartRow
                            End If

                            ChkDt = xlWorkSheet.Cells(CopyStrtColNum, 2).Text
                            ColNum = 1
                            If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
                                dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, ChkDt)

                                'If dtEqualizeData.Rows.Count > 0 Then
                                '    RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
                                '    dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)

                                '    For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
                                '        xlWorkSheet.Cells(CopyStrtColNum, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                '        ColNum = ColNum + 1
                                '        xlShtLastColNum = i + colCnt
                                '    Next
                                'End If

                                If dtEqualizeData.Rows.Count > 0 Then
                                    For i As Integer = 0 To 50
                                        If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
                                            If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
                                                xlWorkSheet.Cells(CopyStrtColNum, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                ColNum = ColNum + 1
                                                xlShtLastColNum = i + colCnt
                                            End If
                                        End If
                                    Next
                                End If
                            Else
                                'Get Last date of selecte Region
                                ChkDt = xlWorkSheet.Cells(FormulaStartRow + formulaRowCnt - 3, 2).Text.ToString
                                If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, Convert.ToDateTime(planStartDate).ToString("dd-MMM-yyyy"))

                                    'Changed on 18 Aug 2011
                                    For index2 As Integer = CopyStrtColNum To CopyStrtColNum + formulaRowCnt - 3
                                        ChkDt = xlWorkSheet.Cells(index2, 2).Text
                                        If Convert.ToDateTime(ChkDt) = Convert.ToDateTime(planStartDate) Then
                                            If dtEqualizeData.Rows.Count > 0 Then
                                                For i As Integer = 0 To 50
                                                    If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
                                                        If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
                                                            xlWorkSheet.Cells(index2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                            ColNum = ColNum + 1
                                                            xlShtLastColNum = i + colCnt
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                    '=====================================
                                End If
                            End If

                            If Not AddHeader Then xlWorkSheet.Range(FormulaStartRow & ":" & FormulaStartRow).Interior.ColorIndex = 47.6
                            FormulaStartRow = FormulaEndRow + 1
                            AddHeader = False
                        Next

                        MaxColName = objclsExcel.ExcelColName(MaxColCnt)

                        FormatPlanWiseExcelSheet(xlWorkSheet, MaxColCnt, drSelect.Length + 1, IntFormulaStartcol, strEqualizationDate)   'Added by Tushar as on 16082012

                        ''FormatPlanWiseExcelSheet(xlWorkSheet, MaxColCnt, drSelect.Length + 1, IntFormulaStartcol, strEqualizationDate)    'Commented by Tushar as on 16082012
                        usedWrkShts = usedWrkShts + 1
                    End If
                Next

                If usedWrkShts > 1 Then
                    xlWorkSheetCurrScheme.Cells.Clear()
                    xlWorkSheetCurrScheme.Activate()
                    xlWorkSheet.Range("A:B").Copy()
                    xlWorkSheetCurrScheme.Range("A1").Select()
                    xlWorkSheetCurrScheme.Paste()
                    xlWorkSheetCurrScheme.Range("1:1").Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

                    Dim LastRowCnt As Long
                    LastRowCnt = xlWorkSheetCurrScheme.Cells.Find(What:="*", After:=xlWorkSheetCurrScheme.Cells(2, 1),
                                             SearchOrder:=Excel.XlSearchOrder.xlByRows,
                                             SearchDirection:=Excel.XlSearchDirection.xlPrevious).Row

                    FirstColorIndex = True

                    CurrentColName = ""
                    colNumToPaste = 3
                    stColPlanNum = 3
                    endColPlanNum = 3
                    AddNewRow = True
                    CopyStrtColNum = IntFormulaStartcol

                    Dim StrColName2 As String = ""
                    Dim ColMainNum As Integer = 3
                    Dim wrkShtNum As Integer = 0

                    Dim addStartColName As String
                    Dim AddEndColName As String

                    Dim addStartColNameMain As String
                    Dim AddEndColNameMain As String

                    Dim StartPlanNum As Integer = 3
                    Dim EndPlanNum As Integer = 3
                    Dim Num As Long
                    Dim strNum As String
                    Dim colInputData As Long = 0
                    For index As Integer = startWrkSht + 1 To usedWrkShts
                        xlWorkSheet = xlWorkBook.Worksheets(index)
                        ColMainNum = 3
                        CopyStrtColNum = 3
                        colNumToPaste = 3 + wrkShtNum
                        xlWorkSheet.Range("1:1").Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

                        'Copy Scheme Total 
                        If wrkShtNum = 0 Then
                            strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                            Num = 2 + objEqualizeTotalSchemeInputCol
                            strNum = objclsExcel.ExcelColName(Num)
                            xlWorkSheet.Range("C:" & strNum).Copy()
                            xlWorkSheetCurrScheme.Activate()
                            xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
                            xlWorkSheetCurrScheme.Paste()
                            xlWorkSheetCurrScheme.Cells(1, 3).value = xlWorkSheetCurrScheme.Name
                            xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, "C1:" & strNum & "1", True)
                        End If
                        '=========================================
                        CopyStrtColNum = objEqualizeTotalSchemeInputCol + 3 '11
                        ColMainNum = objEqualizeTotalSchemeInputCol + 3 '11
                        colNumToPaste = objEqualizeTotalSchemeInputCol + 3 '11
                        If wrkShtNum > 0 Then

                            colNumToPaste = colNumToPaste + (objEqualizeTotalInputCol * wrkShtNum) - 2           'Added by Tushar as on 17082012

                            addStartColName = objclsExcel.ExcelColName(colNumToPaste)
                            addStartColNameMain = objclsExcel.ExcelColName(colNumToPaste + (objEqualizeTotalInputCol - 3))      'Replace - 1 to - 3
                            xlWorkSheetCurrScheme.Range(addStartColName & ":" & addStartColNameMain).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                            addStartColName = objclsExcel.ExcelColName(ColMainNum)
                            addStartColNameMain = objclsExcel.ExcelColName(ColMainNum - 1 + (objEqualizeTotalInputCol * wrkShtNum) - 2)         'Added by Tushar as on 17082012
                            xlWorkSheet.Range(addStartColName & ":" & addStartColNameMain).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                            CopyStrtColNum = ColMainNum - 1 + (objEqualizeTotalInputCol * wrkShtNum) + 1
                            ColMainNum = CopyStrtColNum - 2
                            colNumToPaste = CopyStrtColNum - 2
                        End If

                        'To Copy INPUT DATA
                        strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                        strColName = objclsExcel.ExcelColName(ColMainNum)
                        CopyStrtColNum = ColMainNum + objEqualizeTotalInputCol - 3  'Replace - 1 - 3 by Tushar as on 17082012 
                        StrColName2 = objclsExcel.ExcelColName(CopyStrtColNum)

                        xlWorkSheet.Range(strColName & ":" & StrColName2).Copy()
                        xlWorkSheetCurrScheme.Activate()
                        xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
                        xlWorkSheetCurrScheme.Paste()
                        xlWorkSheetCurrScheme.Cells(1, colNumToPaste).value = xlWorkSheet.Name
                        xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strColName & "1:" & StrColName2 & "1", True)
                        '==============================================

                        CopyStrtColNum = CopyStrtColNum + 1
                        ColMainNum = CopyStrtColNum
                        colNumToPaste = colNumToPaste + objEqualizeTotalInputCol + (wrkShtNum) - 2   '- 2 added by Tushar as on 17082012

                        Dim oldColNumToPaste As Long = colNumToPaste
                        For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
                            Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
                            If wrkShtNum > 0 And (val = "" Or val = "TRUE" Or val = "1") Then
                                ColMainNum = ColMainNum + 1
                                colNumToPaste = colNumToPaste + 1
                                Continue For
                            End If
                            If wrkShtNum > 0 Then
                                addStartColName = objclsExcel.ExcelColName(ColMainNum)
                                addStartColNameMain = objclsExcel.ExcelColName(ColMainNum + 1)
                                ColMainNum = ColMainNum + wrkShtNum
                                AddEndColName = objclsExcel.ExcelColName(ColMainNum - 1)
                                AddEndColNameMain = objclsExcel.ExcelColName(ColMainNum)
                                xlWorkSheet.Range(addStartColName & ":" & AddEndColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
                                strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                                xlWorkSheetCurrScheme.Range(strPastColName & ":" & strPastColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
                            End If

                            'strPastColName = objExcel.ExcelColName(colNumToPaste)
                            'strEndPlan = strPastColName

                            'strColName = objExcel.ExcelColName(ColMainNum)

                            ''Changed
                            'xlWorkSheet.Range(strColName & ":" & strColName).Copy()
                            'xlWorkSheetCurrScheme.Activate()
                            'xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
                            'xlWorkSheetCurrScheme.Paste()
                            'xlWorkSheetCurrScheme.Cells(2, colNumToPaste).Value = xlWorkSheet.Name.ToString
                            'xlWorkSheetCurrScheme.Cells(1, colNumToPaste).Value = xlWorkSheet.Cells(2, ColMainNum).value

                            ColMainNum = ColMainNum + 1
                            colNumToPaste = colNumToPaste + wrkShtNum + 1
                        Next

                        'CopyStrtColNum = CopyStrtColNum + 1
                        ColMainNum = CopyStrtColNum
                        colNumToPaste = oldColNumToPaste 'colNumToPaste + objEqualizeTotalInputCol + (wrkShtNum)
                        For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
                            Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
                            If wrkShtNum > 0 And (val = "" Or val = "TRUE" Or val = "1") Then
                                ColMainNum = ColMainNum + 1
                                colNumToPaste = colNumToPaste + 1
                                Continue For
                            End If

                            If wrkShtNum > 0 Then
                                '    addStartColName = objExcel.ExcelColName(ColMainNum)
                                '    addStartColNameMain = objExcel.ExcelColName(ColMainNum + 1)
                                ColMainNum = ColMainNum + wrkShtNum
                                '    AddEndColName = objExcel.ExcelColName(ColMainNum - 1)
                                '    AddEndColNameMain = objExcel.ExcelColName(ColMainNum)
                                '    xlWorkSheet.Range(addStartColName & ":" & AddEndColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
                                strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                                '    xlWorkSheetCurrScheme.Range(strPastColName & ":" & strPastColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
                            End If

                            strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                            strEndPlan = strPastColName

                            strColName = objclsExcel.ExcelColName(ColMainNum)

                            'Changed
                            xlWorkSheet.Range(strColName & ":" & strColName).Copy()
                            xlWorkSheetCurrScheme.Activate()
                            xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
                            xlWorkSheetCurrScheme.Paste()
                            xlWorkSheetCurrScheme.Cells(2, colNumToPaste).Value = xlWorkSheet.Name.ToString
                            xlWorkSheetCurrScheme.Cells(1, colNumToPaste).Value = xlWorkSheet.Cells(2, ColMainNum).value

                            ColMainNum = ColMainNum + 1
                            colNumToPaste = colNumToPaste + wrkShtNum + 1
                        Next

                        wrkShtNum = wrkShtNum + 1
                    Next

                    ''Added By Shweta'
                    ''To avoid ms calculation due to formula contains data for next column
                    'wrkShtNum = 0

                    'For index As Integer = startWrkSht + 1 To usedWrkShts
                    '    xlWorkSheet = xlWorkBook.Worksheets(index)
                    '    ColMainNum = 3
                    '    CopyStrtColNum = 3
                    '    colNumToPaste = 3 + wrkShtNum

                    '    CopyStrtColNum = objEqualizeTotalSchemeInputCol + 3 '11
                    '    ColMainNum = objEqualizeTotalSchemeInputCol + 3 '11
                    '    colNumToPaste = objEqualizeTotalSchemeInputCol + 3 '11

                    '    If wrkShtNum > 0 Then
                    '        colNumToPaste = colNumToPaste + (objEqualizeTotalInputCol * wrkShtNum)
                    '        CopyStrtColNum = ColMainNum - 1 + (objEqualizeTotalInputCol * wrkShtNum) + 1
                    '        ColMainNum = CopyStrtColNum
                    '        colNumToPaste = CopyStrtColNum
                    '    End If


                    '    CopyStrtColNum = CopyStrtColNum + 1
                    '    ColMainNum = CopyStrtColNum
                    '    colNumToPaste = colNumToPaste + objEqualizeTotalInputCol + (wrkShtNum)

                    '    For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
                    '        Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
                    '        If wrkShtNum > 0 And (val = "" Or val = "TRUE" Or val = "1") Then
                    '            ColMainNum = ColMainNum + 1
                    '            colNumToPaste = colNumToPaste + 1
                    '            Continue For
                    '        End If

                    '        If wrkShtNum > 0 Then
                    '            addStartColName = objExcel.ExcelColName(ColMainNum)
                    '            addStartColNameMain = objExcel.ExcelColName(ColMainNum + 1)

                    '            ColMainNum = ColMainNum + wrkShtNum

                    '            AddEndColName = objExcel.ExcelColName(ColMainNum - 1)
                    '            AddEndColNameMain = objExcel.ExcelColName(ColMainNum)

                    '            'xlWorkSheet.Range(addStartColName & ":" & AddEndColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                    '            strPastColName = objExcel.ExcelColName(colNumToPaste)
                    '            'xlWorkSheetCurrScheme.Range(strPastColName & ":" & strPastColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
                    '        End If

                    '        strPastColName = objExcel.ExcelColName(colNumToPaste)
                    '        strEndPlan = strPastColName

                    '        strColName = objExcel.ExcelColName(ColMainNum)

                    '        'Changed
                    '        xlWorkSheet.Range(strColName & ":" & strColName).Copy()
                    '        xlWorkSheetCurrScheme.Activate()
                    '        xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
                    '        xlWorkSheetCurrScheme.Paste()
                    '        xlWorkSheetCurrScheme.Cells(2, colNumToPaste).Value = xlWorkSheet.Name.ToString
                    '        xlWorkSheetCurrScheme.Cells(1, colNumToPaste).Value = xlWorkSheet.Cells(2, ColMainNum).value

                    '        ColMainNum = ColMainNum + 1
                    '        colNumToPaste = colNumToPaste + wrkShtNum + 1
                    '    Next

                    '    wrkShtNum = wrkShtNum + 1

                    'Next



                    'Changed By Shweta(21 Jun 2012)
                    ' StartPlanNum = 11 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol) + 1
                    StartPlanNum = objEqualizeTotalInputCol + 1 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol) - 3  'Replace +1 to -3 by Tushar as on 17082012
                    '===========================================
                    'StartPlanNum = 11 + ((usedWrkShts - startWrkSht) * 8)
                    EndPlanNum = StartPlanNum
                    For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
                        Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
                        If val = "" Or val = "TRUE" Or val = "1" Then
                            EndPlanNum = StartPlanNum
                        Else
                            EndPlanNum = StartPlanNum + (usedWrkShts - startWrkSht - 1)
                        End If

                        strStartPlan = objclsExcel.ExcelColName(StartPlanNum)
                        strEndPlan = objclsExcel.ExcelColName(EndPlanNum)


                        'To add Total column In Worksheet 
                        If val <> "" And val <> "TRUE" And val <> "1" Then
                            val = dtSampleColData.Rows(index1)("ColShowTotal").ToString.Trim.ToUpper()
                            If val = "" Or val = "TRUE" Or val = "1" Then
                                TotalPlanColNum = EndPlanNum + 1
                                strTotalPlanCol = objclsExcel.ExcelColName(TotalPlanColNum)
                                xlWorkSheetCurrScheme.Range(strTotalPlanCol & ":" & strTotalPlanCol).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
                                xlWorkSheetCurrScheme.Cells(2, strTotalPlanCol).Value = "Total"
                                xlWorkSheetCurrScheme.Cells(2, strTotalPlanCol).Font.Bold = True
                                xlWorkSheetCurrScheme.Cells(3, strTotalPlanCol).Value = "=SUM(" & strStartPlan & "3:" & strEndPlan & "3)" '=SUM(C3:C3)
                                xlWorkSheetCurrScheme.Range(strTotalPlanCol & 3).AutoFill(xlWorkSheetCurrScheme.Range(strTotalPlanCol & "3:" & strTotalPlanCol & LastRowCnt), Excel.XlAutoFillType.xlFillDefault)
                                xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strTotalPlanCol & "2:" & strTotalPlanCol & LastRowCnt, False)

                                strEndPlan = strTotalPlanCol
                                EndPlanNum = TotalPlanColNum
                            End If
                        End If
                        '--------------------------------    

                        If FirstColorIndex Then
                            xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & LastRowCnt).Interior.ColorIndex = 35
                            FirstColorIndex = False
                        Else
                            xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & LastRowCnt).Interior.ColorIndex = 36
                            FirstColorIndex = True
                        End If

                        xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & "1").Select()
                        xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strStartPlan & "1:" & strEndPlan & "1", True)
                        StartPlanNum = EndPlanNum + 1
                    Next


                    xlWorkSheetCurrScheme.Columns.AutoFit()
                    'To hide the Row Data
                    StartPlanNum = 3
                    'Changed By Shweta(21 Jun 2012)
                    'EndPlanNum = 11 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol)
                    EndPlanNum = objEqualizeTotalInputCol + 1 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol) - 4 'Addedd by Tushar as on 22082012 (-4)
                    '======================================================
                    'strStartPlan = objExcel.ExcelColName(StartPlanNum)
                    strEndPlan = objclsExcel.ExcelColName(EndPlanNum)
                    xlWorkSheetCurrScheme.Range("C:" & strEndPlan).EntireColumn.Hidden = True

                    For index As Integer = startWrkSht + 1 To usedWrkShts
                        xlWorkSheet = xlWorkBook.Worksheets(startWrkSht + 1)
                        xlWorkSheet.Delete()
                    Next
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub FormatPlanWiseExcelSheet(ByVal xlSht As Excel.Worksheet, ByVal colCnt As Long, ByVal RowCnt As Long, ByVal FormulaStartColName As Long, ByVal strEqualizationDate As String) '
        Try
            xlSht.Activate()
            If RowCnt > 2 Then
                Dim ChkDt As String
                ChkDt = Convert.ToDateTime(xlSht.Cells(3, 2).Value.ToString).ToString("dd-MMM-yyyy")
                While Convert.ToDateTime(strEqualizationDate) > Convert.ToDateTime(ChkDt)
                    xlSht.Range("3:3").Copy()
                    xlSht.Range("A3").PasteSpecial(Excel.XlPasteType.xlPasteValues)
                    xlSht.Range("2:2").Delete()
                    RowCnt = RowCnt - 1
                    ChkDt = Convert.ToDateTime(xlSht.Cells(3, 2).Value.ToString).ToString("dd-MMM-yyyy")
                End While
            End If
            xlSht.Range("A:A").Delete()
            ''Insert Day Column
            xlSht.Cells(2, colCnt).Value = "=IF(A2="""","""",TEXT(A2,""dddd""))"
            Dim colNm As String = objclsExcel.ExcelColName(colCnt)
            If RowCnt > 2 Then
                xlSht.Range(colNm & "2").AutoFill(xlSht.Range(colNm & "2:" & colNm & RowCnt), Excel.XlAutoFillType.xlFillValues)
            End If
            xlSht.Cells(1, colCnt).Value = "DAY"
            xlSht.Columns(colNm & ":" & colNm).Select()
            xlSht.Application.CutCopyMode = False
            xlSht.Application.Selection.Cut()
            xlSht.Columns("B:B").Select()
            xlSht.Application.Selection.insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight)
            With xlSht
                .Range("B:B").Copy()
                .Range("B1").PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, False)
            End With
            xlSht.Calculate()
            xlSht.Range("A:A").NumberFormat = "dd-mmm-yyyy"

            'For Formatting Excel Sheet
            xlSht.Columns.AutoFit()
            With xlSht.Range("A1:" & colNm & RowCnt)
                .Select()
                With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
            End With
            'xlSht.Cells.NumberFormat = "#,##0.00"
            xlSht.Columns.AutoFit()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Function FormatRange(ByVal xlSht As Excel.Worksheet, ByVal rng As String, ByVal setAllign As Boolean) As Excel.Worksheet
        Try
            With xlSht.Range(rng)
                If setAllign Then
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .MergeCells = False
                    .Font.Bold = True
                    .Merge()
                End If
                With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With
                With .Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                End With


                'With .Borders(Excel.XlBordersIndex.xlEdgeLeft)
                '    .LineStyle = Excel.XlLineStyle.xlContinuous
                '    .ColorIndex = 0
                '    .TintAndShade = 0
                '    .Weight = Excel.XlBorderWeight.xlThin
                'End With
                'With .Borders(Excel.XlBordersIndex.xlEdgeTop)
                '    .LineStyle = Excel.XlLineStyle.xlContinuous
                '    .ColorIndex = 0
                '    .TintAndShade = 0
                '    .Weight = Excel.XlBorderWeight.xlThin
                'End With
                'With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                '    .LineStyle = Excel.XlLineStyle.xlContinuous
                '    .ColorIndex = 0
                '    .TintAndShade = 0
                '    .Weight = Excel.XlBorderWeight.xlThin
                'End With
                'With .Borders(Excel.XlBordersIndex.xlEdgeRight)
                '    .LineStyle = Excel.XlLineStyle.xlContinuous
                '    .ColorIndex = 0
                '    .TintAndShade = 0
                '    .Weight = Excel.XlBorderWeight.xlThin
                'End With
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
        Return xlSht
    End Function

    Private Sub BtnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBrowse.Click
        Try
            If FolderBrowserDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                TxtReportPath.Text = FolderBrowserDialog1.SelectedPath
                objEqualizeDefaultReportPath = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub rBtnReport_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rBtnReport.CheckedChanged
        Try
            If rBtnReport.Checked Then
                cmbRptFormat.Enabled = True
            Else
                cmbRptFormat.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    'Private Function OpenFileDlg() As String
    '    Dim sFileName As String = ""
    '    Try
    '        OpenFileDialog1.Filter = "Excel Files|*.xls"
    '        If OpenFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then

    '            sFileName = OpenFileDialog1.FileName
    '        Else
    '            MsgBox("Path is Not Selected", MsgBoxStyle.Information, "DBBilling")
    '            'Exit Function
    '        End If

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '    End Try
    '    Return sFileName
    'End Function    'Private Function PrintRowDataToExcel(ByVal xlSht As Excel.Worksheet) As Excel.Worksheet
    '    Try
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '    End Try
    'End Function

    'Private Sub SelectDataToCalculate()
    '    Try
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '    End Try
    'End Sub 'Private Function FormatExcelSheet(ByVal xlSht As Excel.Worksheet, ByVal startColNum As Integer, ByVal LastRowNum As Long, ByVal dtTemplateData As DataTable, ByVal dtTemplateColData As DataTable) As Excel.Worksheet
    '    Try
    '        startColNum = 9
    '        Dim FormulaRowNum As Long = LastRowNum
    '        'Dim colCnt As Integer = dtData.Columns.Count - startColNum + 1
    '        'Dim strcolName As String

    '        For index1 As Integer = 0 To dtTemplateData.Rows.Count - 1
    '            If index1 < dtTemplateData.Rows.Count - 1 Then

    '            End If
    '        Next

    '        'For index1 As Integer = 0 To dtTemplateColData.Rows.Count - 1
    '        '    xlWorkSheet.Cells(1, index1 + colCnt).value = dtTemplateColData.Rows(index1)("Column_Header").ToString
    '        '    xlWorkSheet.Cells(3, index1 + colCnt).value = dtTemplateColData.Rows(index1)("Column_Formula").ToString
    '        '    strcolName = objExcel.ExcelColName(index1 + colCnt)
    '        '    xlWorkSheet.Range(xlWorkSheet.Cells(3, index1 + colCnt)).AutoFill(xlWorkSheet.Range(xlWorkSheet.Cells(3, index1 + colCnt), xlWorkSheet.Cells(FormulaRowNum, index1 + colCnt)), Excel.XlAutoFillType.xlFillDefault)
    '        'Next


    '        'If dtBtweenFromTo.Rows.Count > 0 Then

    '        'End If

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '    End Try
    'End Function

    Private Sub chkSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectAll.CheckedChanged
        Try
            If userChngData = True Then
                For index As Integer = 0 To dgvSchemes.Rows.Count - 1
                    dgvSchemes.Rows(index).Cells("colChk").Value = chkSelectAll.Checked
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub btnEqualize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEqualize.Click
        Try
            strYesNoFlag = ""
            lstEqulizeStatus.Items.Clear()
            Dim MFundId As Long
            ' If rBtnEqualize.Checked Then
            If cmbMFund.SelectedIndex < 0 Then
                MessageBox.Show("Please Select the Mutual Fund", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            If chkEquReport.Checked Then
                MessageBox.Show("Equalization Report will not be generate during Equalization process. Please do not select Equalization Report.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            strFundCode = cmbMFund.GetItemText(cmbMFund.SelectedItem)
            drSelect = dtFundData.Select("Mutual_Fund_Code='" & strFundCode & "'")
            MFundId = 0
            If drSelect.Length > 0 Then
                MFundId = drSelect(0)("AutoId")
            End If
            Dim dataInUsed As Long = objEqualize_clsDAL.LockFund(MFundId)
            If dataInUsed <= 0 Then
                MessageBox.Show("Selected Fund Is in Use. Please Try Again later.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                'objExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
                Exit Sub
            End If
            'End If
            EqualizeData("Equalization")
            'CheckEqualizedSchemes()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Finally
            CheckEqualizedSchemes()
        End Try
    End Sub

    Private Sub btnEqulizeStatusOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEqulizeStatusOK.Click
        Try
            lstEqulizeStatus.Items.Clear()
            pnlEqualizeStatus.Enabled = False
            pnlEqualizeStatus.SendToBack()
            pnlMainData.Enabled = True
            pnlMainData.BringToFront()
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

#Region "For GL Report"
    Private Sub GetGLReportData(ByVal dtRptTempateData As DataTable, ByVal strFromDate As String, ByVal strToDate As String, ByVal lngSchemeId As Long)
        Try
            Dim strRptFromDate As String
            If Convert.ToDateTime(strFromDate) < Convert.ToDateTime(DTPFromDate.Text.ToString) Then
                strRptFromDate = DTPFromDate.Text.ToString
            Else
                strRptFromDate = strFromDate
            End If
            'Dim dtBeforeFromTemplate As DataTable
            'dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Get_Template_Before_From_Date", SchemeId, strFromDate)
            'dtRptTempateData = objEqualize_clsDAL.GetApprovedData("Get_Template_Between_From_Date", SchemeId, strFromDate, strToDate)
            ''merge first template to start the calculation in the one datatable
            'drAddRow = dtRptTempateData.NewRow
            'For index1 As Integer = 0 To dtBeforeFromTemplate.Columns.Count - 1
            '    drAddRow(index1) = dtBeforeFromTemplate.Rows(0)(index1)
            'Next
            'dtRptTempateData.Rows.InsertAt(drAddRow, 0)
            'dtRptTempateData.AcceptChanges()
            dtRptTempateData.Columns.Add("ExpiryDate")
            Dim strDtExpDate As String = ""
            For index As Integer = 0 To dtRptTempateData.Rows.Count - 1
                If index = dtRptTempateData.Rows.Count - 1 Then
                    strDtExpDate = strToDate
                Else
                    If Convert.ToDateTime(strRptFromDate) > dtRptTempateData.Rows(index)("EffectiveDate") And Convert.ToDateTime(strRptFromDate) > dtRptTempateData.Rows(index + 1)("EffectiveDate") Then
                        strDtExpDate = "Delete"
                    Else
                        strDtExpDate = Convert.ToDateTime(dtRptTempateData.Rows(index + 1)("EffectiveDate")).AddDays(-1)
                    End If
                End If
                'Changed by shweta (03Jan2011)
                'To Convert date in correct format
                dtRptTempateData.Rows(index)("ExpiryDate") = Convert.ToDateTime(strDtExpDate).ToString("dd-MMM-yyyy")
            Next
            Dim drselect() As DataRow
            drselect = dtRptTempateData.Select("ExpiryDate='Delete'")
            For Each dr As DataRow In drselect
                dr.Delete()
            Next
            'Dim rowNum As Long = 0
            For index As Integer = 0 To dtRptTempateData.Rows.Count - 1
                Dim fromDt As String = dtRptTempateData.Rows(index)("EffectiveDate")
                If Convert.ToDateTime(strRptFromDate) > dtRptTempateData.Rows(index)("EffectiveDate") Then
                    fromDt = strRptFromDate
                End If
                Dim Todt As String = dtRptTempateData.Rows(index)("ExpiryDate")
                Dim dtTbleTemap As DataSet = objEqualize_clsDAL.GetReportData("GL Report", lngSchemeId, fromDt, Todt, lngGLRptRowNum)
                If Not IsNothing(dtTbleTemap) Then
                    If dtTbleTemap.Tables.Count > 0 Then
                        If lngGLRptRowNum > 0 Then
                            dtGLReportData.Merge(dtTbleTemap.Tables(0), True, MissingSchemaAction.Ignore)
                        Else
                            dtGLReportData = dtTbleTemap.Tables(0).Copy()
                        End If
                        If dtTbleTemap.Tables(0).Rows.Count > 0 Then lngGLRptRowNum = dtTbleTemap.Tables(0).Compute("max(RowNumber)", "")
                    End If
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub GenerateGLReport(ByVal strFundName As String, ByVal strRptDate As String)
        Try
            If Not IsNothing(dtGLReportData) Then
                If dtGLReportData.Rows.Count > 0 Then
                    Dim rptPath As String = TxtReportPath.Text.Trim
                    If rptPath.Substring(rptPath.Length - 1, 1) = "\" Then
                        ' rptPath = rptPath & "GL Report\" & strRptDate
                        rptPath = rptPath & "GL Report"
                    Else
                        'rptPath = rptPath & "\" & "GL Report\" & strRptDate
                        rptPath = rptPath & "\" & "GL Report"
                    End If
                    If dtGLReportData.Columns.Contains("RowNumber") Then dtGLReportData.Columns.Remove("RowNumber")
                    objEqualize_clsBal.NewFolderCheckOrCreate(rptPath)
                    rptPath = rptPath & "\" & strFundName & ".xlsx"
                    'xlWorkSheet.Name = strFundName
                    dtGLReportData.TableName = strFundName
                    objEqualize_clsExcel.CreateGLReportFile(dtGLReportData, rptPath, 0, 0, False, , "GL Report")
                    'Dim Path As String = ""
                    'Path = objEqualize_clsBal.NewFolderCheckOrCreate(Path)
                    'Path = Path & "\FundData_Log.xlsx"
                    'If File.Exists(path) Then
                    '    File.Delete(path)
                    'End If                    
                    Dim lstItem As New ListViewItem
                    lstItem.Text = strFundName
                    lstItem.SubItems.Add("GL Report Generated successfully. At Path " & rptPath)
                    lstEqulizeStatus.Items.Add(lstItem)
                    'MessageBox.Show("Data Exported Successfully,At path " & rptPath, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Else
                    Dim lstItem As New ListViewItem
                    lstItem.Text = strFundName
                    lstItem.SubItems.Add("Data not exists to Generate the GL Report(Define Mapping column for GL Report in Template)")
                    lstEqulizeStatus.Items.Add(lstItem)
                End If
            Else
                Dim lstItem As New ListViewItem
                lstItem.Text = strFundName
                lstItem.SubItems.Add("Data not exists to Generate the GL Report(Define Mapping column for GL Report in Template)")
                lstEqulizeStatus.Items.Add(lstItem)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

#End Region


    Private Sub CreateDivDistTable()
        Try
            If dtDivDistData.Columns.Count <= 0 Then
                dtDivDistData = objEqualize_clsDAL.GetApprovedData("Get Table Structure", "Equalize_MstDividendDistData")
                dtDivDistData.Columns.Add("SchemeName")
                dtDivDistData.Columns.Add("SchemeCode")
                dtDivDistData.Columns.Add("PlanName")
                dtDivDistData.Columns.Add("PlanCode")
                dtDivDistData.Columns.Add("PlanRnTCode")
                'dtDivDistData.Columns.Add("SheetHeader")
                'dtDivDistData.Columns.Add("BeforeNAVDate")
                'dtDivDistData.Columns.Add("BeforeNAV")
                'dtDivDistData.Columns.Add("NAVDate")
                'dtDivDistData.Columns.Add("NAV")
                'dtDivDistData.Columns.Add("Distributable_Surplus")
                'dtDivDistData.Columns.Add("No_of_Units")
                'dtDivDistData.Columns.Add("Adhoc_Rate")
                'dtDivDistData.Columns.Add("Individual_Rate")
                'dtDivDistData.Columns.Add("Corporate_Rate")
                'dtDivDistData.Columns.Add("Individual_Unit")
                'dtDivDistData.Columns.Add("Corporate_Unit")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub GetDivDistReportEqualize(ByVal dtPlanData As DataTable, ByVal strSchemeCode As String, ByVal strToDate As String, ByVal lngSchemeId As Long)
        Try
            ', ByVal lngSchemeId As Long
            Dim savepath As String = ""
            Dim dtTbleTemap As DataSet = objEqualize_clsDAL.GetReportData("Dividend Distribution Report", lngSchemeId, strToDate)
            'objEqualize_clsExcel.CreateDistributionRptFile(dtTbleTemap)
            If IsNothing(dtDivDistData) Then
                dtDivDistData = New DataTable
            End If
            If dtDivDistData.Rows.Count > 0 Then
                dtDivDistData.Merge(dtTbleTemap.Tables(0), True, MissingSchemaAction.Ignore)
            Else
                dtDivDistData = dtTbleTemap.Tables(0).Copy()
            End If


            'If dtDivDistData.Columns.Count <= 0 Then CreateDivDistTable()
            'For index As Integer = 0 To dtTbleTemap.Tables(1).Rows.Count - 1
            '    Dim strPlanId As String = dtTbleTemap.Tables(1).Rows(index)("PlanID")
            '    Dim strSheetHeader As String = strFundCode.ToUpper & " FUND - PLAN " & dtTbleTemap.Tables(1).Rows(index)("PlanCode").ToString.ToUpper & " " & dtTbleTemap.Tables(1).Rows(index)("PlanName").ToString.ToUpper
            '    Dim drUnit() As DataRow
            '    drUnit = dtTbleTemap.Tables(2).Select("PlanID = '" & strPlanId & "'")
            '    Dim drDivData() As DataRow
            '    drDivData = dtTbleTemap.Tables(0).Select("PlanID = '" & strPlanId & "'")

            '    Dim drDiv As DataRow
            '    drDiv = dtDivDistData.NewRow
            '    drDiv("SchemeID") = lngSchemeId
            '    'drDiv("SchemeName") = ""
            '    drDiv("SchemeCode") = strSchemeCode
            '    drDiv("PlanID") = strPlanId
            '    drDiv("PlanName") = dtTbleTemap.Tables(1).Rows(index)("PlanName")
            '    drDiv("PlanCode") = dtTbleTemap.Tables(1).Rows(index)("PlanCode")
            '    drDiv("PlanRnTCode") = dtTbleTemap.Tables(1).Rows(index)("PlanRnTCode")

            '    drDiv("SheetHeader") = strSheetHeader
            '    If drDivData.Length > 0 Then
            '        drDiv("Distributable_Surplus") = drDivData(0)("Distributable_Surplus").ToString
            '        drDiv("BeforeNAVDate") = drDivData(0)("BEFORE_NAV DATE").ToString
            '        drDiv("BeforeNAV") = drDivData(0)("Before NAV Data").ToString
            '        drDiv("NAVDate") = drDivData(0)("NAV DATE").ToString
            '        drDiv("NAV") = drDivData(0)("NAV Data").ToString
            '        drDiv("No_of_Units") = drDivData(0)("Num_Units").ToString

            '        If drDivData(0)("Adhoc_Rate").ToString <> "" Then
            '            If IsNumeric(drDivData(0)("Adhoc_Rate").ToString) Then
            '                drDiv("Adhoc_Rate") = drDivData(0)("Adhoc_Rate").ToString
            '            End If
            '        End If
            '    End If
            '    If dtTbleTemap.Tables(3).Rows.Count > 0 Then
            '        drDiv("Individual_Rate") = dtTbleTemap.Tables(3).Rows(0)("IndividualRate")
            '        drDiv("Corporate_Rate") = dtTbleTemap.Tables(3).Rows(0)("CorporateRate")
            '    End If
            '    If drUnit.Length > 0 Then
            '        drDiv("Individual_Unit") = drUnit(0)("IndividualUnit")
            '        drDiv("Corporate_Unit") = drUnit(0)("CorporateUnit")
            '    End If
            '    dtDivDistData.Rows.Add(drDiv)
            '    'objEqualize_clsExcel.CreateDistributionRptFile(drDiv, ClsCommon.userName.Trim, savepath)
            'Next

            Dim drSelect() As DataRow
            drSelect = dtDivDistData.Select("SchemeID = '" & lngSchemeId & "'")

            If drSelect.Length > 0 Then
                ' CreateDistributionRptFile(lngSchemeId, drSelect, strToDate, strSchemeCode)
                objEqualize_clsExcel.CreateDistributionRptFile(lngSchemeId, drSelect, strToDate, strSchemeCode, ClsCommon.userName.Trim, TxtReportPath.Text.Trim, True, chkPrint.Checked, PrintDialog1, chk_WithFormula.Checked) REM 'chk_WithFormula.Checked' added by sameer on 18-Jan-2013

                Dim lstItem As New ListViewItem
                lstItem.Text = strSchemeCode
                lstItem.SubItems.Add("Dividend Distribution Report generate successfully.")
                lstEqulizeStatus.Items.Add(lstItem)
            Else
                Dim lstItem As New ListViewItem
                lstItem.Text = strSchemeCode
                lstItem.SubItems.Add("Plan not available to generate Dividend Distribution Report")
                lstEqulizeStatus.Items.Add(lstItem)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub GetDivDataFromDataBase(ByVal strSchemeCode As String, ByVal lngSchemeId As Long, ByVal strToDate As String)
        Try
            Dim dtTbleTemap As DataTable = objEqualize_clsDAL.GetApprovedData("Approved Dividend Dist Data", ClsCommon.userName.Trim, lngSchemeId, strToDate)
            Dim drSelect() As DataRow
            drSelect = dtTbleTemap.Select("SchemeID = '" & lngSchemeId & "'")
            If drSelect.Length > 0 Then
                'CreateDistributionRptFile(lngSchemeId, drSelect, strToDate, strSchemeCode)
                objEqualize_clsExcel.CreateDistributionRptFile(lngSchemeId, drSelect, strToDate, strSchemeCode, ClsCommon.userName.Trim, TxtReportPath.Text.Trim, False, chkPrint.Checked, PrintDialog1, chk_WithFormula.Checked) REM 'chk_WithFormula.Checked' added by sameer on 18-Jan-2013
                Dim lstItem As New ListViewItem
                lstItem.Text = strSchemeCode
                lstItem.SubItems.Add("Dividend Distribution Report generate successfully.")
                lstEqulizeStatus.Items.Add(lstItem)
            Else
                Dim lstItem As New ListViewItem
                lstItem.Text = strSchemeCode
                lstItem.SubItems.Add("Data not available to generate Dividend Distribution Report")
                lstEqulizeStatus.Items.Add(lstItem)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Public Sub CreateDistributionRptFile(ByVal lngSchemeId As Long, ByVal dtDivDistData() As DataRow, ByVal strToDate As String, ByVal strSchemeName As String)
        'Dim XlProcessIdDivDist As Integer
        'Dim xlAppDivDist As Excel.Application
        'Dim xlWorkBookDivDist As Excel.Workbook
        'Dim xlWorkSheetDivDist As Excel.Worksheet
        'Dim xlWorkSheetDivDistSample As Excel.Worksheet


        'Try
        '    Dim strTemplatePath As String = ConfigurationManager.AppSettings("TemplatePath")
        '    strTemplatePath = strTemplatePath & "\Dividend Distribution Report"

        '    'dtTbl = objFeeSch.ImportUploadedData(tblName, UpMonth, UpYear)
        '    'Dim ReportSavePath As String
        '    'ReportSavePath = objFeeSch.SchSysLogPath & "\" & Name & "_" & UpMonth & "-" & UpYear & ".xls"
        '    Dim strSavePath As String = TxtReportPath.Text.Trim
        '    If strSavePath.Substring(strSavePath.Length - 1, 1) = "\" Then
        '        strSavePath = strSavePath & "Dividend Dist Report\" & strToDate
        '    Else
        '        strSavePath = strSavePath & "\" & "Dividend Dist Report\" & strToDate
        '    End If
        '    objEqualize_clsBal.NewFolderCheckOrCreate(strSavePath)
        '    strSavePath = strSavePath & "\" & strSchemeName & ".xlsx"


        '    objEqualize_clsExcel.Initialise_ExcelObj(xlAppDivDist, XlProcessIdDivDist) 'xlWorkBook,
        '    xlWorkBookDivDist = xlAppDivDist.Workbooks.Open(strTemplatePath)
        '    xlAppDivDist.Visible = False
        '    'While xlWorkBook.Worksheets.Count > 1
        '    '    xlWorkSheet = xlWorkBook.Worksheets(2)
        '    '    xlWorkSheet.Delete()
        '    'End While
        '    xlWorkSheetDivDistSample = xlWorkBookDivDist.Worksheets("SampleFile")
        '    For index As Integer = 0 To dtDivDistData.Length - 1
        '        xlWorkSheetDivDistSample.Copy(, After:=xlWorkBookDivDist.Sheets(xlWorkBookDivDist.Sheets.Count)) ', Count:=1, Type:=Excel.XlSheetType.xlWorksheet
        '        xlWorkSheetDivDist = xlWorkBookDivDist.Sheets(xlWorkBookDivDist.Sheets.Count)
        '        xlWorkSheetDivDist.Name = "PLAN " & dtDivDistData(index)("Plan Code").ToString.ToUpper & " " & dtDivDistData(index)("Plan Name").ToString.ToUpper
        '        xlWorkSheetDivDist = objEqualize_clsExcel.CreateDistributionRptFile(xlWorkSheetDivDist, dtDivDistData(index), ClsCommon.userName)
        '    Next

        '    'xlWorkSheetDivDist = xlWorkBookDivDist.Worksheets("SampleFile")
        '    xlWorkSheetDivDistSample.Delete()
        '    xlAppDivDist.DisplayAlerts = False
        '    ' If ReportSavePath <> "" Then
        '    xlWorkBookDivDist.SaveAs(strSavePath)
        '    xlWorkBookDivDist.Close()
        '    'Else
        '    'xlApp.Visible = True
        '    'End If
        '    xlAppDivDist.Quit()
        '    GC.Collect()
        '    GC.WaitForPendingFinalizers()
        '    objEqualize_clsExcel.ReleaseObject(xlWorkSheetDivDist)
        '    GC.Collect()
        '    GC.WaitForPendingFinalizers()
        '    objEqualize_clsExcel.ReleaseObject(xlWorkBookDivDist)
        '    GC.Collect()
        '    GC.WaitForPendingFinalizers()
        '    objEqualize_clsExcel.ReleaseObject(xlAppDivDist)
        '    'Destroy_ExcelObj(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
        '    Dim lstItem As New ListViewItem
        '    lstItem.Text = strSchemeName
        '    lstItem.SubItems.Add("Dividend Distribution Report generate successfully. At Path " & strSavePath)
        '    lstEqulizeStatus.Items.Add(lstItem)
        'Catch ex As Exception
        '    If Not IsNothing(xlApp) Then
        '        GC.Collect()
        '        GC.WaitForPendingFinalizers()
        '        objEqualize_clsExcel.ReleaseObject(xlWorkSheetDivDist)
        '        GC.Collect()
        '        GC.WaitForPendingFinalizers()
        '        xlWorkBook.Close()
        '        objEqualize_clsExcel.ReleaseObject(xlWorkBookDivDist)
        '        GC.Collect()
        '        GC.WaitForPendingFinalizers()
        '        xlApp.Quit()
        '        objEqualize_clsExcel.ReleaseObject(xlAppDivDist)
        '        'Destroy_ExcelObj(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
        '    End If
        '    MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        'End Try
    End Sub

    Private Sub btnSaveDivData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveDivData.Click
        Try
            If IsNothing(dtDivDistData) Then
                MessageBox.Show("Data not available to save", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            If dtDivDistData.Rows.Count <= 0 Then
                MessageBox.Show("Data not available to save", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            For index As Integer = 0 To dtDivDistData.Rows.Count - 1
                Dim lngPlanID As Long = dtDivDistData.Rows(index)("PlanID").ToString
                Dim strNAVDate As String = dtDivDistData.Rows(index)("NAV Date").ToString
                Dim strStatus As String = ""
                strStatus = objEqualize_clsDAL.AddDividendDistData("Check Data", lngPlanID, strNAVDate)

                Dim boolAddData As Boolean = False
                If strStatus <> "" Then
                    Dim ans = MessageBox.Show(strStatus & " data already exist for Scheme " & dtDivDistData.Rows(index)("Scheme Code").ToString & vbCrLf & " Do you want to Replace data?", "DBEqualization", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    If ans = vbYes Then
                        boolAddData = True
                    End If
                Else
                    boolAddData = True
                End If
                If boolAddData Then
                    strStatus = objEqualize_clsDAL.AddDividendDistData("Add", lngPlanID, strNAVDate, dtDivDistData.Rows(index), 0, ClsCommon.userName)
                End If
            Next
            dtDivDistData = Nothing
            MessageBox.Show("Data added successfully.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    'Public Sub EqualizeData()
    '    Try
    '        Dim chkBool As Boolean
    '        Dim MFundId As Long
    '        Dim boolGetTempData As Boolean

    '        Dim boolRprGenrate As Boolean = False
    '        Dim usedWrkShts As Integer = 0
    '        Dim strRptPath As String = TxtReportPath.Text.Trim
    '        Dim SavePath As String = strRptPath
    '        Dim fileDetail As IO.FileInfo
    '        Dim extc As String = ""
    '        Dim FileNamelen As Integer
    '        Dim RptFileName As String = ""
    '        Dim strStartEqualizationDate As String
    '        Dim boolEqDataNotFound As Boolean = False




    '        If cmbMFund.SelectedIndex < 0 Then
    '            MessageBox.Show("Please Select the Mutual Fund", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '            Exit Sub
    '        End If
    '        If strRptPath = "" And rBtnReport.Checked Then
    '            MessageBox.Show("Report Save path is not assigned. Please assign the report save path first", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '            Exit Sub
    '        End If

    '        If rBtnReport.Checked And cmbRptFormat.SelectedIndex < 0 Then
    '            MessageBox.Show("Please select type of Report ", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '            Exit Sub
    '        End If
    '        rptType = cmbRptFormat.GetItemText(cmbRptFormat.SelectedItem)


    '        strFromDate = Convert.ToDateTime(DTPFromDate.Text).ToString("dd-MMM-yyyy") 'Convert.ToDateTime(DTPFromDate.Text).AddDays(-1).ToString("dd-MMM-yyyy")
    '        strStartEqualizationDate = strFromDate
    '        strToDate = dtpToDate.Text

    '        If Convert.ToDateTime(strFromDate) > Convert.ToDateTime(strToDate) Then
    '            MessageBox.Show("From date must be less than To Date", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '            Exit Sub
    '        End If

    '        dtDgvSchemeData = dgvSchemes.DataSource
    '        If IsNothing(dtDgvSchemeData) Then
    '            MessageBox.Show("Scheme not exist to Equalize data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '            Exit Sub
    '        End If
    '        If dtDgvSchemeData.Rows.Count < 0 Then
    '            MessageBox.Show("Scheme not exist to Equalize data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '            Exit Sub
    '        End If
    '        drSelectedScheme = dtDgvSchemeData.Select("ColChk =True")
    '        'If drSelectedScheme.Length <= 0 Then
    '        '    MessageBox.Show("Please select atleast 1 scheme to Calculate Equalization", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '        '    Exit Sub
    '        'End If

    '        Dim boolSchemeSelected As Boolean = False
    '        For index As Integer = 0 To dtDgvSchemeData.Rows.Count - 1
    '            If dtDgvSchemeData.Rows(index)("ColChk") Then
    '                boolSchemeSelected = True
    '                Exit For
    '            End If
    '        Next
    '        If boolSchemeSelected = False Then
    '            MessageBox.Show("Please select atleast 1 scheme to Calculate Equalization", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '            Exit Sub
    '        End If




    '        strFundCode = cmbMFund.GetItemText(cmbMFund.SelectedItem)
    '        drSelect = dtFundData.Select("Mutual_Fund_Code='" & strFundCode & "'")
    '        MFundId = 0
    '        If drSelect.Length > 0 Then
    '            MFundId = drSelect(0)("AutoId")
    '        End If

    '        strFromDate = Convert.ToDateTime(strFromDate).ToString("dd-MMM-yyyy")

    '        If rBtnReport.Checked Then
    '            fileDetail = My.Computer.FileSystem.GetFileInfo(strRptPath)
    '            extc = fileDetail.Extension
    '            FileNamelen = strRptPath.LastIndexOf(extc)
    '            RptFileName = strRptPath.Substring(0, FileNamelen)
    '            chkBool = System.IO.Directory.Exists(fileDetail.DirectoryName)

    '            If chkBool = False Then
    '                MessageBox.Show("Selected Path not exist. Please Select the Proper Path.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                Exit Sub
    '            End If
    '            objExcel = New clsEqualization_Excel
    '            objExcel.Initialise_ExcelObj(xlApp, xlWorkBook, XlProcessId)
    '            xlApp.Visible = False
    '            If rptType = "Plan on Single Worksheet" Then
    '                CreateNewWorkbook()
    '                usedWrkShts = 1
    '            End If
    '        Else
    '            objExcel = New clsEqualization_Excel
    '            objExcel.Initialise_ExcelObj(xlApp, xlWorkBook, XlProcessId)
    '            xlApp.Visible = False
    '            CreateNewWorkbook()
    '            usedWrkShts = 1
    '        End If

    '        'Dim dtChkData As New DataTable
    '        'Dim dtPlanData As New DataTable
    '        boolGetTempData = False
    '        For index As Integer = 0 To dtDgvSchemeData.Rows.Count - 1 'drSelectedScheme.Length - 1
    '            If dtDgvSchemeData.Rows(index)("ColChk") = False Then
    '                Continue For
    '            End If
    '            boolEqDataNotFound = False
    '            SchemeCode = dtDgvSchemeData.Rows(index)("Scheme_Code").ToString
    '            SchemeId = Convert.ToInt64(dtDgvSchemeData.Rows(index)("AutoId").ToString)

    '            'SchemeCode = drSelectedScheme(index)("Scheme_Code").ToString
    '            'SchemeId = Convert.ToInt64(drSelectedScheme(index)("AutoId").ToString)

    '            dtPlanData = objEqualize_clsDAL.GetApprovedData("PlanWithDivCal", "", SchemeId, strToDate) 
    '            If dtPlanData.Rows.Count <= 0 Then
    '                MessageBox.Show("Plan not Exist for Scheme " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                Continue For
    '            End If
    '            drSelect = dtPlanData.Select("Dividend_Frequency = 'Not Set'")
    '            If drSelect.Length > 0 Then
    '                MessageBox.Show("Dividend Frequency is not set for some plans of " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                Continue For
    '            End If

    '            newStrFromDate = strFromDate

    '            'if Data is never equalized for selectd Scheme then start the equaliztion from Open Date
    '            dtLastEqualizeDate = New DataTable
    '            dtLastEqualizeDate = objEqualize_clsDAL.GetApprovedData("Approved Input Data Open Date", SchemeId)
    '            If dtLastEqualizeDate.Rows.Count > 0 Then
    '                SchemeOpenDate = dtLastEqualizeDate.Rows(0)("DataDate").ToString
    '            End If
    '            If SchemeOpenDate = "" Then
    '                MessageBox.Show("Input Data Not exist Scheme code " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                Continue For
    '            End If


    '            If rBtnEqualize.Checked Then
    '                'Get the Last Equalized date 
    '                dtLastEqualizeDate = New DataTable
    '                dtLastEqualizeDate = objEqualize_clsDAL.GetApprovedData("Equalize_Data_LastEqualizeDate", "", SchemeId)
    '                newStrFromDate = ""
    '                If dtLastEqualizeDate.Rows.Count > 0 Then
    '                    newStrFromDate = dtLastEqualizeDate.Rows(0)("dataDate").ToString
    '                End If
    '                If newStrFromDate = "" Then
    '                    boolEqDataNotFound = False
    '                End If

    '                If newStrFromDate = "" Then
    '                    'if Equalization is never done for Selected Scheme
    '                    strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy") 'strFromDate 'Convert.ToDateTime(SchemeOpenDate).AddDays(2).ToString("dd-MMM-yyyy")
    '                    strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")

    '                Else
    '                    If Convert.ToDateTime(newStrFromDate) < Convert.ToDateTime(strFromDate) Then
    '                        'If Start date is greater than last equalization date then start equalization after Last equalized date
    '                        strStartEqualizationDate = Convert.ToDateTime(newStrFromDate).AddDays(1).ToString("dd-MMM-yyyy")
    '                        strFromDate = Convert.ToDateTime(strStartEqualizationDate).AddDays(-1).ToString("dd-MMM-yyyy")
    '                    Else
    '                        If Convert.ToDateTime(SchemeOpenDate) > Convert.ToDateTime(strFromDate) Then
    '                            'If Start date is less than open date then start equalization then start equalization from open date
    '                            strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy") 'Convert.ToDateTime(SchemeOpenDate).AddDays(2).ToString("dd-MMM-yyyy")
    '                            strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
    '                        Else
    '                            'If Start date is greater than open date then start equalization from From date 
    '                            strStartEqualizationDate = Convert.ToDateTime(strFromDate).ToString("dd-MMM-yyyy")
    '                            'Select From date less than 2 day from equalization date or Scheme open date
    '                            If Convert.ToDateTime(strStartEqualizationDate).AddDays(-1) < Convert.ToDateTime(SchemeOpenDate) Then
    '                                strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
    '                            Else
    '                                strFromDate = Convert.ToDateTime(strStartEqualizationDate).AddDays(-1).ToString("dd-MMM-yyyy")
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            Else
    '                'Check Data already Equalize or Not
    '                dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalize_Data_From_To_To", SchemeId.ToString, DTPFromDate.Text, strToDate) 
    '                If dtEqualizeData.Rows.Count <= 0 Then
    '                    MessageBox.Show("Data is not Equalized for selected period of Scheme " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                    Continue For
    '                Else
    '                    'Get the Min Equalize date as Start Date
    '                    'And Max Equalize date as end date between the selected Period
    '                    strStartEqualizationDate = Convert.ToDateTime(dtEqualizeData.Rows(0)("dataDate").ToString).ToString("dd-MMM-yyyy")

    '                    If Convert.ToDateTime(strStartEqualizationDate).AddDays(-1) < Convert.ToDateTime(SchemeOpenDate) Then
    '                        strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
    '                        'strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
    '                    Else
    '                        strFromDate = Convert.ToDateTime(dtEqualizeData.Rows(0)("dataDate").ToString).AddDays(-1).ToString("dd-MMM-yyyy")
    '                    End If

    '                    strToDate = Convert.ToDateTime(dtEqualizeData.Rows(dtEqualizeData.Rows.Count - 1)("dataDate").ToString).ToString("dd-MMM-yyyy")
    '                End If
    '            End If

    '            'If boolGetTempData = False Or newStrFromDate <> strFromDate Then
    '            'To check Template exist or not to create report
    '            dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Get_Template_Before_From_Date", SchemeId, strFromDate)
    '            If dtBeforeFromTemplate.Rows.Count <= 0 Then
    '                MessageBox.Show("Template Not exist for Selected period of Scheme Code = " & SchemeCode & vbCrLf & ". Please create Template to calculate Data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                Continue For
    '            End If
    '            dtBtweenFromToTemplate = objEqualize_clsDAL.GetApprovedData("Get_Template_Between_From_Date", SchemeId, strFromDate, strToDate) 

    '            'merge first template to start the calculation in the one datatable
    '            drAddRow = dtBtweenFromToTemplate.NewRow
    '            For index1 As Integer = 0 To dtBeforeFromTemplate.Columns.Count - 1
    '                drAddRow(index1) = dtBeforeFromTemplate.Rows(0)(index1)
    '            Next
    '            dtBtweenFromToTemplate.Rows.InsertAt(drAddRow, 0)
    '            dtBtweenFromToTemplate.AcceptChanges()
    '            dtTemplateColData = New DataTable
    '            dtTemplateColData = objEqualize_clsDAL.GetApprovedData("Get_All_Selected_Template_Columns", SchemeId, strFromDate, strToDate) 
    '            boolGetTempData = True
    '            'End If

    '            'Check Input data exist or not for selected From to To date.
    '            dtInputData = objEqualize_clsDAL.GetApprovedData("Input_Data_From_To_To", SchemeId.ToString, strFromDate, strToDate) 
    '            If dtInputData.Rows.Count <= 0 Then
    '                MessageBox.Show("Input Data Not exist for selected period for Scheme code " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                Continue For
    '            End If

    '            'Dim isValid As Boolean = True
    '            If rBtnEqualize.Checked Then
    '                While xlWorkBook.Worksheets.Count > 1
    '                    xlWorkSheet = xlWorkBook.Worksheets(2)
    '                    xlWorkSheet.Delete()
    '                End While
    '                Dim datauploaded As Boolean = SaveEqualData(MFundId, SchemeId, strFundCode, SchemeCode, strStartEqualizationDate, strFromDate, strToDate, dtPlanData, dtInputData) 'strStartEqualizationDate
    '                'If datauploaded Then
    '                '    MessageBox.Show("Data Equalization completed successfully Of the Scheme code " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                'End If
    '            End If
    '            If rBtnReport.Checked Then
    '                strDateNotFound = CheckAllDataExistsOrNot(SchemeCode, dtInputData, dtPlanData, strFromDate, strToDate)
    '                If strDateNotFound <> "" Then
    '                    MessageBox.Show("Input Data Not exist of the day " & strDateNotFound & " For Scheme code " & SchemeCode & " . Please Upload the data First", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '                    Continue For
    '                End If
    '                If rptType = "Plan on Single Worksheet" Then
    '                    xlWorkSheet = xlWorkBook.Worksheets(1)
    '                    xlWorkSheet.Activate()
    '                    If usedWrkShts > 0 Then xlWorkBook.Worksheets.Add()
    '                    xlWorkSheetCurrScheme = xlWorkBook.Worksheets(1)
    '                    If usedWrkShts > 0 Then xlWorkSheetCurrScheme.Move(, xlWorkBook.Worksheets(usedWrkShts + 1))
    '                    xlWorkSheetCurrScheme.Name = "Scheme-" & SchemeCode
    '                    usedWrkShts = usedWrkShts + 1
    '                    SavePath = strRptPath
    '                Else
    '                    CreateNewWorkbook()
    '                    usedWrkShts = 1
    '                    SavePath = RptFileName & "-" & SchemeCode & extc
    '                End If
    '                'xlApp.Visible = True
    '                GenerateRpt(usedWrkShts, dtInputData, dtPlanData, dtBtweenFromToTemplate, dtTemplateColData, SavePath, strStartEqualizationDate)
    '                boolRprGenrate = True
    '            End If
    '        Next
    '        If rBtnReport.Checked Then
    '            If boolRprGenrate Then
    '                If rptType = "Plan on Single Worksheet" Then
    '                    xlApp.DisplayAlerts = False
    '                    xlWorkSheetSample.Delete()
    '                    xlWorkBook.SaveAs(strRptPath, Excel.XlFileFormat.xlExcel7)
    '                    xlWorkBook.Close()
    '                End If
    '                objExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
    '                MessageBox.Show("Report Generated successfully.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '            Else
    '                objExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
    '            End If
    '        Else
    '            objExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
    '        End If

    '        For i As Integer = 0 To dgvSchemes.Rows.Count - 1
    '            dgvSchemes.Rows(i).Cells("colChk").Value = False
    '        Next
    '        CheckedSchemeToEqualize(True)

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    '        objExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
    '    End Try
    'End Sub

    Public Sub EqualizeData_OLD()
        Dim MFundId As Long
        Try
            Dim chkBool As Boolean

            Dim boolGetTempData As Boolean

            Dim boolRprGenrate As Boolean = False
            Dim usedWrkShts As Integer = 0
            Dim strRptPath As String = TxtReportPath.Text.Trim
            Dim SavePath As String = strRptPath
            Dim fileDetail As IO.FileInfo
            Dim extc As String = ""
            Dim FileNamelen As Integer
            Dim RptFileName As String = ""
            Dim strStartEqualizationDate As String
            Dim boolEqDataNotFound As Boolean = False

            strFundCode = cmbMFund.GetItemText(cmbMFund.SelectedItem)
            drSelect = dtFundData.Select("Mutual_Fund_Code='" & strFundCode & "'")
            MFundId = 0
            If drSelect.Length > 0 Then
                MFundId = drSelect(0)("AutoId")
            End If


            If strRptPath = "" And rBtnReport.Checked Then
                MessageBox.Show("Report Save path is not assigned. Please assign the report save path first", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If



            If rBtnReport.Checked And cmbRptFormat.SelectedIndex < 0 Then
                MessageBox.Show("Please select type of Report ", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            rptType = cmbRptFormat.GetItemText(cmbRptFormat.SelectedItem)


            strFromDate = Convert.ToDateTime(DTPFromDate.Text).ToString("dd-MMM-yyyy")
            'Convert.ToDateTime(DTPFromDate.Text).AddDays(-1).ToString("dd-MMM-yyyy")
            strStartEqualizationDate = strFromDate
            strToDate = dtpToDate.Text

            If Convert.ToDateTime(strFromDate) > Convert.ToDateTime(strToDate) Then
                MessageBox.Show("From date must be less than To Date", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If

            dtDgvSchemeData = dgvSchemes.DataSource
            If IsNothing(dtDgvSchemeData) Then
                MessageBox.Show("Scheme not exist to Equalize data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            If dtDgvSchemeData.Rows.Count < 0 Then
                MessageBox.Show("Scheme not exist to Equalize data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            drSelectedScheme = dtDgvSchemeData.Select("ColChk =True")
            'If drSelectedScheme.Length <= 0 Then
            '    MessageBox.Show("Please select atleast 1 scheme to Calculate Equalization", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            '    Exit Sub
            'End If

            Dim boolSchemeSelected As Boolean = False
            For index As Integer = 0 To dtDgvSchemeData.Rows.Count - 1
                If dtDgvSchemeData.Rows(index)("ColChk") Then
                    boolSchemeSelected = True
                    Exit For
                End If
            Next
            If boolSchemeSelected = False Then
                MessageBox.Show("Please select atleast 1 scheme to Calculate Equalization", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If





            strFromDate = Convert.ToDateTime(strFromDate).ToString("dd-MMM-yyyy")
            If rBtnReport.Checked Then
                'fileDetail = My.Computer.FileSystem.GetFileInfo(strRptPath)
                'extc = fileDetail.Extension
                'FileNamelen = strRptPath.LastIndexOf(extc)
                'RptFileName = strRptPath.Substring(0, FileNamelen)
                'chkBool = System.IO.Directory.Exists(fileDetail.DirectoryName)

                'If chkBool = False Then
                '    MessageBox.Show("Selected Path not exist. Please Select the Proper Path.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                '    Exit Sub
                'End If
                'objExcel = New clsEqualization_Excel
                'objExcel.Initialise_ExcelObj(xlApp, XlProcessId) ' xlWorkBook,
                'xlApp.Visible = False
                'If rptType = "Plan on Single Worksheet" Then
                '    CreateNewWorkbook()
                '    usedWrkShts = 1
                'End If


                fileDetail = My.Computer.FileSystem.GetFileInfo(strRptPath)
                extc = fileDetail.Extension
                If extc.ToLower <> ".xls" And extc.ToLower <> ".xlsx" Then
                    MessageBox.Show("Selected File Path  extension is invalid", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Exit Sub
                End If
                FileNamelen = strRptPath.LastIndexOf(extc)
                RptFileName = strRptPath.Substring(0, FileNamelen)
                Dim DirName As String = "" '= strRptPath.Substring(0, strRptPath.LastIndexOf("\"))
                If strRptPath.LastIndexOf("\") >= 0 Then
                    DirName = strRptPath.Substring(0, strRptPath.LastIndexOf("\"))
                End If


                If Not Directory.Exists(DirName) Then
                    MessageBox.Show("Selected Path not exist. Please Select the Proper Path.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Exit Sub
                End If

                'chkBool = System.IO.Directory.Exists(fileDetail.DirectoryName)
                'If chkBool = False Then
                '    MessageBox.Show("Selected Path not exist. Please Select the Proper Path.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                '    Exit Sub
                'End If

                objclsExcel = New clsEqualization_Excel
                objclsExcel.Initialise_ExcelObj(xlApp, XlProcessId) ' xlWorkBook,
                xlApp.Visible = False
                If rptType = "Plan on Single Worksheet" Then
                    CreateNewWorkbook()
                    usedWrkShts = 1
                End If
            Else
                objclsExcel = New clsEqualization_Excel
                objclsExcel.Initialise_ExcelObj(xlApp, XlProcessId) ' xlWorkBook,
                xlApp.Visible = False
                CreateNewWorkbook()
                usedWrkShts = 1
            End If

            Me.Cursor = Cursors.WaitCursor
            boolGetTempData = False
            For index As Integer = 0 To dtDgvSchemeData.Rows.Count - 1 'drSelectedScheme.Length - 1
                'MessageBox.Show("In For Loop", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)


                If dtDgvSchemeData.Rows(index)("ColChk") = False Then
                    Continue For
                End If
                boolEqDataNotFound = False
                SchemeCode = dtDgvSchemeData.Rows(index)("Scheme_Code").ToString
                'MessageBox.Show("SchemeCode=" & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                SchemeId = Convert.ToInt64(dtDgvSchemeData.Rows(index)("AutoId").ToString)
                'MessageBox.Show("SchemeId=" & SchemeId, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                'MessageBox.Show("Before PlanWithDivCal", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                dtPlanData = objEqualize_clsDAL.GetApprovedData("PlanWithDivCal", "", SchemeId, strToDate, strFromDate)
                'MessageBox.Show("After PlanWithDivCal", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                If dtPlanData.Rows.Count <= 0 Then
                    Me.Cursor = Cursors.Arrow
                    MessageBox.Show("Plan not Exist for Scheme " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Continue For
                End If
                'drSelect = dtPlanData.Select("Dividend_Frequency = 'Not Set'")
                'If drSelect.Length > 0 Then
                '    Me.Cursor = Cursors.Arrow
                '    MessageBox.Show("Dividend Frequency is not set for some plans of " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                '    Continue For
                'End If

                newStrFromDate = strFromDate

                'if Data is never equalized for selectd Scheme then start the equaliztion from Open Date
                dtLastEqualizeDate = New DataTable
                dtLastEqualizeDate = objEqualize_clsDAL.GetApprovedData("Approved Input Data Open Date", SchemeId)
                'MessageBox.Show("After Approved Input Data Open Date", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                If dtLastEqualizeDate.Rows.Count > 0 Then
                    SchemeOpenDate = dtLastEqualizeDate.Rows(0)("DataDate").ToString
                End If
                If SchemeOpenDate = "" Then
                    Me.Cursor = Cursors.Arrow
                    MessageBox.Show("Input Data Not exist Scheme code " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Continue For
                End If


                If rBtnEqualize.Checked Then
                    'Get the Last Equalized date 
                    dtLastEqualizeDate = New DataTable
                    dtLastEqualizeDate = objEqualize_clsDAL.GetApprovedData("Equalize_Data_LastEqualizeDate", "", SchemeId)
                    'MessageBox.Show("Equalize_Data_LastEqualizeDatee", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                    newStrFromDate = ""
                    If dtLastEqualizeDate.Rows.Count > 0 Then
                        newStrFromDate = dtLastEqualizeDate.Rows(0)("dataDate").ToString
                    End If
                    If newStrFromDate = "" Then
                        boolEqDataNotFound = False
                    End If

                    If newStrFromDate = "" Then
                        'if Equalization is never done for Selected Scheme
                        strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                        strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                    Else
                        If Convert.ToDateTime(newStrFromDate) < Convert.ToDateTime(strFromDate) Then
                            'If Start date is greater than last equalization date then start equalization after Last equalized date
                            strStartEqualizationDate = Convert.ToDateTime(newStrFromDate).AddDays(1).ToString("dd-MMM-yyyy")
                            strFromDate = Convert.ToDateTime(strStartEqualizationDate).AddDays(-1).ToString("dd-MMM-yyyy")
                        Else
                            If Convert.ToDateTime(SchemeOpenDate) > Convert.ToDateTime(strFromDate) Then
                                'If Start date is less than open date then start equalization then start equalization from open date
                                strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                                'Convert.ToDateTime(SchemeOpenDate).AddDays(2).ToString("dd-MMM-yyyy")
                                strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                            Else
                                'If Start date is greater than open date then start equalization from From date 
                                strStartEqualizationDate = Convert.ToDateTime(strFromDate).ToString("dd-MMM-yyyy")
                                'Select From date less than 2 day from equalization date or Scheme open date
                                If Convert.ToDateTime(strStartEqualizationDate).AddDays(-1) < Convert.ToDateTime(SchemeOpenDate) Then
                                    strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                                Else
                                    strFromDate = Convert.ToDateTime(strStartEqualizationDate).AddDays(-1).ToString("dd-MMM-yyyy")
                                End If
                            End If
                        End If
                    End If
                Else
                    'Check Data already Equalize or Not
                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalize_Data_From_To_To", SchemeId.ToString, DTPFromDate.Text, strToDate)
                    If dtEqualizeData.Rows.Count <= 0 Then
                        Me.Cursor = Cursors.Arrow
                        MessageBox.Show("Data is not Equalized for selected period of Scheme " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Continue For
                    Else
                        'Get the Min Equalize date as Start Date
                        'And Max Equalize date as end date between the selected Period
                        strStartEqualizationDate = Convert.ToDateTime(dtEqualizeData.Rows(0)("dataDate").ToString).ToString("dd-MMM-yyyy")

                        If Convert.ToDateTime(strStartEqualizationDate).AddDays(-1) < Convert.ToDateTime(SchemeOpenDate) Then
                            strFromDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                            'strStartEqualizationDate = Convert.ToDateTime(SchemeOpenDate).ToString("dd-MMM-yyyy")
                        Else
                            strFromDate = Convert.ToDateTime(dtEqualizeData.Rows(0)("dataDate").ToString).AddDays(-1).ToString("dd-MMM-yyyy")
                        End If

                        strToDate = Convert.ToDateTime(dtEqualizeData.Rows(dtEqualizeData.Rows.Count - 1)("dataDate").ToString).ToString("dd-MMM-yyyy")
                    End If
                End If
                'MessageBox.Show("Before template", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                'If boolGetTempData = False Or newStrFromDate <> strFromDate Then
                'To check Template exist or not to create report
                dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Get_Template_Before_From_Date", SchemeId, strFromDate)
                'MessageBox.Show("After : Get_Template_Before_From_Date", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                If dtBeforeFromTemplate.Rows.Count <= 0 Then
                    Me.Cursor = Cursors.Arrow
                    MessageBox.Show("Template Not exist for Selected period of Scheme Code = " & SchemeCode & vbCrLf & ". Please create Template to calculate Data.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Continue For
                End If
                dtBtweenFromToTemplate = objEqualize_clsDAL.GetApprovedData("Get_Template_Between_From_Date", SchemeId, strFromDate, strToDate)
                'MessageBox.Show("After : Get_Template_Between_From_Date", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)


                'merge first template to start the calculation in the one datatable
                drAddRow = dtBtweenFromToTemplate.NewRow
                For index1 As Integer = 0 To dtBeforeFromTemplate.Columns.Count - 1
                    drAddRow(index1) = dtBeforeFromTemplate.Rows(0)(index1)
                Next
                dtBtweenFromToTemplate.Rows.InsertAt(drAddRow, 0)
                dtBtweenFromToTemplate.AcceptChanges()
                dtTemplateColData = New DataTable
                dtTemplateColData = objEqualize_clsDAL.GetApprovedData("Get_All_Selected_Template_Columns", SchemeId, strFromDate, strToDate)

                boolGetTempData = True


                'Check Input data exist or not for selected From to To date.
                dtInputData = objEqualize_clsDAL.GetApprovedData("Input_Data_From_To_To", SchemeId.ToString, strFromDate, strToDate)
                'MessageBox.Show("After : Input_Data_From_To_To", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                If dtInputData.Rows.Count <= 0 Then
                    Me.Cursor = Cursors.Arrow
                    MessageBox.Show("Input Data Not exist for selected period for Scheme code " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Continue For
                End If

                'Dim isValid As Boolean = True
                If rBtnEqualize.Checked Then
                    While xlWorkBook.Worksheets.Count > 1
                        xlWorkSheet = xlWorkBook.Worksheets(2)
                        xlWorkSheet.Delete()
                    End While
                    'MessageBox.Show("Before Equalize", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

                    Dim datauploaded As Boolean = SaveEqualData(MFundId, SchemeId, strFundCode, SchemeCode, strStartEqualizationDate, strFromDate, strToDate, dtPlanData, dtInputData) 'strStartEqualizationDate
                    'If datauploaded Then
                    '    MessageBox.Show("Data Equalization completed successfully Of the Scheme code " & SchemeCode, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    'End If
                End If
                If rBtnReport.Checked Then
                    strDateNotFound = CheckAllDataExistsOrNot(SchemeCode, dtInputData, dtPlanData, strFromDate, strToDate)
                    If strDateNotFound <> "" Then
                        Me.Cursor = Cursors.Arrow
                        MessageBox.Show("Input Data Not exist of the day " & strDateNotFound & " For Scheme code " & SchemeCode & " . Please Upload the data First", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Continue For
                    End If
                    If rptType = "Plan on Single Worksheet" Then
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                        xlWorkSheet.Activate()
                        If usedWrkShts > 0 Then xlWorkBook.Worksheets.Add()
                        xlWorkSheetCurrScheme = xlWorkBook.Worksheets(1)
                        If usedWrkShts > 0 Then xlWorkSheetCurrScheme.Move(, xlWorkBook.Worksheets(usedWrkShts + 1))
                        'strRptPath = strRptPath & "\" & SchemeCode
                        'NewFolderCheckOrCreate
                        xlWorkSheetCurrScheme.Name = "Scheme-" & SchemeCode
                        usedWrkShts = usedWrkShts + 1
                        'SavePath = RptFileName & "-" & SchemeCode & extc
                        SavePath = SavePath
                    Else
                        'objBAL.NewFolderCheckOrCreate(strRptPath)
                        CreateNewWorkbook()
                        usedWrkShts = 1
                        SavePath = RptFileName & "-" & SchemeCode & extc
                    End If
                    'xlApp.Visible = True
                    'GenerateRpt(usedWrkShts, dtInputData, dtPlanData, dtBtweenFromToTemplate, dtTemplateColData, SavePath, strStartEqualizationDate)
                    boolRprGenrate = True
                End If
            Next
            If rBtnReport.Checked Then
                If boolRprGenrate Then
                    If rptType = "Plan on Single Worksheet" Then
                        xlApp.DisplayAlerts = False
                        xlWorkSheetSample.Delete()
                        xlWorkBook.SaveAs(strRptPath, Excel.XlFileFormat.xlExcel7)
                        xlWorkBook.Close()
                    End If
                    objclsExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
                    Me.Cursor = Cursors.Arrow
                    MessageBox.Show("Report Generated successfully.", "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Else
                    objclsExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
                End If
            Else
                objclsExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
            End If

            For i As Integer = 0 To dgvSchemes.Rows.Count - 1
                dgvSchemes.Rows(i).Cells("colChk").Value = False
            Next
            CheckedSchemeToEqualize(True)

        Catch ex As Exception
            Me.Cursor = Cursors.Arrow
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            objclsExcel.ExitExcel(xlApp, xlWorkBook, xlWorkSheet, XlProcessId)
        Finally
            Me.Cursor = Cursors.Arrow
            If rBtnEqualize.Checked Then
                objEqualize_clsDAL.UnLockFund(MFundId)
            End If
        End Try
        Me.Cursor = Cursors.Arrow
    End Sub
    Private Sub GenerateRpt_OLD(ByVal startWrkSht As Integer, ByVal dtData As DataTable, ByVal dtPlanData As DataTable, ByVal dtTemplateData As DataTable, ByVal dtTemplateColData As DataTable, ByVal strRptPath As String, ByVal strEqualizationDate As String)
        'Dim XlProcessId As Integer
        'Dim xlApp As Excel.Application
        'Dim xlWorkBook As Excel.Workbook
        'Dim xlWorkSheet As Excel.Worksheet
        'Dim xlWorkSheetSample As Excel.Worksheet
        Try
            Dim xlShtLastColName As String
            Dim xlShtLastColNum As String

            Dim StartColName As String
            Dim startColNum As Integer = 11
            Dim strColName As String = ""

            Dim RowCnt As Long = 1
            Dim usedWrkShts As Integer = startWrkSht
            Dim PlanCode As String = ""
            Dim PlanId As String = ""
            Dim MaxColCnt As Long = 1

            Dim drSelect() As DataRow
            Dim col, row As Integer


            Dim CurrenTtemplateID As Long
            Dim FormulaEndRow As Long
            Dim FormulaStartRow As Long
            Dim formulaRowCnt As Long
            Dim colCnt As Integer = dtData.Columns.Count - startColNum + 1
            MaxColCnt = dtData.Columns.Count - startColNum
            Dim ChkDt As String
            Dim AddHeader As Boolean
            Dim TempColCnt As Long

            Dim CurrentEffDt As String
            Dim EffectiveDt As String
            Dim showFormula As Boolean


            Dim FirstColorIndex As Boolean
            Dim MaxColName As String = ""
            Dim drSelectCol() As DataRow
            Dim IntFormulaStartcol As Integer
            Dim FormulaStartColName As String



            Dim CurrentColName As String = ""
            Dim strStartPlan As String
            Dim strEndPlan As String

            Dim TotalPlanColNum As Long
            Dim strTotalPlanCol As String

            Dim CopyStrtColNum As Integer
            Dim strPastColName As String
            Dim colNumToPaste As Long
            Dim stColPlanNum As Long
            Dim endColPlanNum As Long
            Dim AddNewRow As Boolean

            dtSampleColData = New DataTable
            dtSampleColData = dtTemplateColData.Copy
            dtSampleColData.Rows.Clear()

            'xlWorkSheetSample.Cells.Clear()
            '
            'xlWorkSheetSample.Activate()
            'xlWorkSheetSample.Range("A1").Select()
            'xlWorkSheetSample.Range("A1").End(Excel.XlDirection.xlToRight).Select()
            'xlShtLastColName = xlWorkSheetSample.

            'xlApp.Visible = True
            Dim PlanClosedDate As String
            If dtData.Rows.Count > 0 Then
                For index As Integer = 0 To dtPlanData.Rows.Count - 1
                    PlanCode = dtPlanData.Rows(index)("Plan_Code").ToString
                    PlanId = dtPlanData.Rows(index)("PlanId").ToString
                    planStartDate = dtPlanData.Rows(index)("Start_Date").ToString
                    PlanClosedDate = dtPlanData.Rows(index)("ClosedDate").ToString
                    drSelect = dtData.Select("PlanId =" & PlanId)
                    If drSelect.Length > 0 Then
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                        xlWorkSheet.Activate()
                        If usedWrkShts > 0 Then xlWorkBook.Worksheets.Add()
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                        If usedWrkShts > 0 Then xlWorkSheet.Move(, xlWorkBook.Worksheets(usedWrkShts + 1))
                        xlWorkSheet.Name = "Plan-" & PlanCode

                        xlWorkSheet.Activate()

                        'Paste Row  Data 
                        Dim rawData(drSelect.Length, dtData.Columns.Count - startColNum) As Object
                        Dim intlen As Integer = drSelect.Rank

                        ' Copy the column names to the first row of the object array
                        For col = 0 To dtData.Columns.Count - startColNum - 1
                            rawData(0, col) = dtData.Columns(col + startColNum).ColumnName.ToUpper
                        Next

                        ' Copy the values to the object array
                        For col = 0 To dtData.Columns.Count - startColNum - 1
                            For row = 0 To drSelect.Length - 1
                                rawData(row + 1, col) = drSelect(row)(col + startColNum)
                            Next
                        Next

                        'Calculate the final column letter
                        Dim finalColLetter As String = String.Empty
                        finalColLetter = objclsExcel.ExcelColName(dtData.Columns.Count - startColNum)
                        Dim excelRange As String = String.Format("A" & RowCnt & ":{0}{1}", finalColLetter, drSelect.Length + RowCnt)
                        xlWorkSheet.Range(excelRange, Type.Missing).NumberFormat = "@"
                        xlWorkSheet.Range(excelRange, Type.Missing).Select()
                        xlWorkSheet.Range(excelRange, Type.Missing).Value2 = rawData
                        xlWorkSheet.Columns.AutoFit()
                        xlWorkSheet.Range("1:1").Font.Bold = True
                        xlWorkSheet.Range("A:" & finalColLetter).Cells.NumberFormat = "#,##0.00"

                        IntFormulaStartcol = dtData.Columns.Count - startColNum
                        FormulaStartColName = objclsExcel.ExcelColName(IntFormulaStartcol)
                        MaxColName = ""


                        'Dim CurrenTtemplateID As Long
                        FormulaEndRow = 3
                        FormulaStartRow = 3
                        colCnt = dtData.Columns.Count - startColNum + 1
                        MaxColCnt = dtData.Columns.Count - startColNum
                        AddHeader = True
                        TempColCnt = 0

                        StartColName = objclsExcel.ExcelColName(colCnt)
                        xlShtLastColNum = colCnt
                        'Get previous Day  Data
                        Dim ColNum As Integer = 1
                        'ChkDt = xlWorkSheet.Cells(2, 2).value

                        'If IsDate(planStartDate) Then
                        '    If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
                        '        dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, ChkDt)

                        '        If dtEqualizeData.Rows.Count > 0 Then
                        '            RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
                        '            dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)
                        '        End If

                        '        For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
                        '            xlWorkSheet.Cells(2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                        '            ColNum = ColNum + 1
                        '            xlShtLastColNum = i + colCnt
                        '        Next
                        '    Else
                        '        dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, Convert.ToDateTime(planStartDate).ToString("dd-MMM-yyyy"))

                        '        If dtEqualizeData.Rows.Count > 0 Then
                        '            RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
                        '            dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)
                        '        End If
                        '        For index2 As Integer = 2 To drSelect.Length + 1
                        '            ChkDt = xlWorkSheet.Cells(index2, 2).value
                        '            If Convert.ToDateTime(ChkDt) = Convert.ToDateTime(planStartDate) Then
                        '                For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
                        '                    xlWorkSheet.Cells(index2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                        '                    ColNum = ColNum + 1
                        '                    xlShtLastColNum = i + colCnt
                        '                Next
                        '                Exit For
                        '            End If
                        '        Next
                        '    End If
                        'End If
                        'xlApp.Visible = True

                        'If drSelect.Length = 1 Then
                        '    FormulaStartRow = 3
                        'End If


                        For index1 As Integer = 0 To dtTemplateData.Rows.Count - 1
                            formulaRowCnt = 2
                            CurrentEffDt = dtTemplateData.Rows(index1)("EffectiveDate").ToString
                            If index1 < dtTemplateData.Rows.Count - 1 Then
                                EffectiveDt = dtTemplateData.Rows(index1 + 1)("EffectiveDate").ToString
                                ChkDt = xlWorkSheet.Cells(FormulaStartRow, 2).value
                                If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(EffectiveDt) Then
                                    Continue For
                                End If
                                For index2 As Integer = FormulaStartRow To drSelect.Length + RowCnt
                                    ChkDt = xlWorkSheet.Cells(index2, 2).value
                                    If Convert.ToDateTime(ChkDt) < Convert.ToDateTime(EffectiveDt) Then
                                        FormulaEndRow = index2
                                        formulaRowCnt = formulaRowCnt + 1
                                    Else
                                        Exit For
                                    End If
                                Next
                            Else
                                FormulaEndRow = drSelect.Length + RowCnt
                                formulaRowCnt = FormulaEndRow - FormulaStartRow + 3
                            End If

                            CurrenTtemplateID = dtTemplateData.Rows(index1)("AutoID")
                            drSelectCol = dtTemplateColData.Select("TemplateID ='" & CurrenTtemplateID & "'", "ColSequenceNum")
                            xlWorkSheetSample.Cells.Clear()
                            xlWorkSheet.Range("A" & FormulaStartRow - 1 & ":" & finalColLetter & FormulaEndRow).Copy()
                            xlWorkSheetSample.Activate()
                            xlWorkSheetSample.Range("A2").Select()
                            xlWorkSheetSample.Paste()

                            xlShtLastColName = objclsExcel.ExcelColName(xlShtLastColNum)
                            xlWorkSheet.Range(StartColName & FormulaStartRow - 1 & ":" & xlShtLastColName & FormulaStartRow - 1).Copy()
                            xlWorkSheetSample.Activate()
                            xlWorkSheetSample.Range(StartColName & 2).PasteSpecial(Excel.XlPasteType.xlPasteValues)

                            'xlApp.Visible = True
                            For i As Integer = 0 To drSelectCol.Length - 1
                                If AddHeader Then
                                    xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                Else
                                    If IsNothing(xlWorkSheet.Cells(1, i + colCnt).value) Then
                                        xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                    ElseIf xlWorkSheet.Cells(1, i + colCnt).value.ToString = "" Then
                                        xlWorkSheet.Cells(1, i + colCnt).value = drSelectCol(i)("ColHeader").ToString
                                    End If
                                End If

                                If drSelect.Length > 1 Then
                                    xlWorkSheetSample.Cells(3, i + colCnt).Value = drSelectCol(i)("ColFormula").ToString
                                    strColName = objclsExcel.ExcelColName(i + colCnt)
                                    If formulaRowCnt > 3 Then

                                        xlWorkSheetSample.Range(strColName & 3).AutoFill(xlWorkSheetSample.Range(strColName & 3 & ":" & strColName & (formulaRowCnt)), Excel.XlAutoFillType.xlFillDefault)
                                    End If

                                End If

                                'to Show Decimal number
                                xlRange = strColName & 3 & ":" & strColName & (formulaRowCnt)
                                Dim Val As Integer = 0
                                If drSelectCol(i)("ColDecimalNum").ToString() <> "" Then
                                    Val = drSelectCol(i)("ColDecimalNum")
                                End If
                                objclsExcel.SetNumberFormatToColumn(Val, xlRange, xlWorkSheetSample)

                                If dtSampleColData.Rows.Count - 1 < i Then
                                    MaxColCnt = MaxColCnt + 1
                                    drAddRow = dtSampleColData.NewRow
                                    drAddRow("ColHeader") = drSelectCol(i)("ColHeader").ToString
                                    drAddRow("ColFormula") = drSelectCol(i)("ColFormula").ToString
                                    drAddRow("ColShowFormula") = drSelectCol(i)("ColShowFormula").ToString
                                    drAddRow("ColIsSchemeWise") = drSelectCol(i)("ColIsSchemeWise").ToString
                                    drAddRow("ColShowTotal") = drSelectCol(i)("ColShowTotal")
                                    drAddRow("ColDecimalNum") = drSelectCol(i)("ColDecimalNum")
                                    dtSampleColData.Rows.Add(drAddRow)
                                    xlShtLastColNum = xlShtLastColNum + 1
                                Else
                                    MaxColCnt = dtData.Columns.Count - startColNum + dtSampleColData.Rows.Count
                                End If
                            Next


                            If drSelect.Length > 1 Then
                                xlWorkSheetSample.Range("A3:" & strColName & (formulaRowCnt)).Copy()
                                xlWorkSheet.Activate()
                                xlWorkSheet.Range("A" & FormulaStartRow).Select()
                                xlWorkSheet.Paste()
                            End If

                            If AddHeader Then
                                CopyStrtColNum = FormulaStartRow - 1
                            Else
                                CopyStrtColNum = FormulaStartRow
                            End If

                            ChkDt = xlWorkSheet.Cells(CopyStrtColNum, 2).value
                            ColNum = 1
                            If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
                                dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, ChkDt)

                                'If dtEqualizeData.Rows.Count > 0 Then
                                '    RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
                                '    dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)
                                'End If
                                'For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
                                '    xlWorkSheet.Cells(CopyStrtColNum, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                '    ColNum = ColNum + 1
                                '    xlShtLastColNum = i + colCnt
                                'Next

                                If dtEqualizeData.Rows.Count > 0 Then
                                    For i As Integer = 0 To 50
                                        If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
                                            If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
                                                xlWorkSheet.Cells(CopyStrtColNum, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                ColNum = ColNum + 1
                                                xlShtLastColNum = i + colCnt
                                            End If
                                        End If
                                    Next
                                End If
                            Else
                                'Get Last date of selecte Region
                                ChkDt = xlWorkSheet.Cells(FormulaStartRow + formulaRowCnt - 3, 2).value
                                If Convert.ToDateTime(ChkDt) >= Convert.ToDateTime(planStartDate) Then
                                    dtEqualizeData = objEqualize_clsDAL.GetApprovedData("Equalization_Data", PlanId, Convert.ToDateTime(planStartDate).ToString("dd-MMM-yyyy"))

                                    'If dtEqualizeData.Rows.Count > 0 Then
                                    '    RptTemplateId = Convert.ToInt64(dtEqualizeData.Rows(0)("RptTemplateId").ToString)
                                    '    dtBeforeFromTemplate = objEqualize_clsDAL.GetApprovedData("Template Column", RptTemplateId)
                                    'End If
                                    'For index2 As Integer = CopyStrtColNum To CopyStrtColNum + formulaRowCnt - 3
                                    '    ChkDt = xlWorkSheet.Cells(index2, 2).value
                                    '    If Convert.ToDateTime(ChkDt) = Convert.ToDateTime(planStartDate) Then
                                    '        For i As Integer = 0 To dtBeforeFromTemplate.Rows.Count - 1
                                    '            xlWorkSheet.Cells(index2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                    '            ColNum = ColNum + 1
                                    '            xlShtLastColNum = i + colCnt
                                    '        Next
                                    '        Exit For
                                    '    End If
                                    'Next

                                    'Changed on 18 Aug 2011
                                    For index2 As Integer = CopyStrtColNum To CopyStrtColNum + formulaRowCnt - 3
                                        ChkDt = xlWorkSheet.Cells(index2, 2).value
                                        If Convert.ToDateTime(ChkDt) = Convert.ToDateTime(planStartDate) Then
                                            If dtEqualizeData.Rows.Count > 0 Then
                                                For i As Integer = 0 To 50
                                                    If dtEqualizeData.Columns.Contains("ColumnValue" & ColNum) Then
                                                        If dtEqualizeData.Rows(0)("ColumnValue" & ColNum).ToString <> "" Then
                                                            xlWorkSheet.Cells(index2, i + colCnt).value = dtEqualizeData.Rows(0)("ColumnValue" & ColNum)
                                                            ColNum = ColNum + 1
                                                            xlShtLastColNum = i + colCnt
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                    '=====================================
                                End If
                            End If
                            If Not AddHeader Then xlWorkSheet.Range(FormulaStartRow & ":" & FormulaStartRow).Interior.ColorIndex = 47.6
                            FormulaStartRow = FormulaEndRow + 1
                            AddHeader = False
                        Next
                        MaxColName = objclsExcel.ExcelColName(MaxColCnt)

                        FormatPlanWiseExcelSheet(xlWorkSheet, MaxColCnt, drSelect.Length + 1, IntFormulaStartcol, strEqualizationDate)
                        usedWrkShts = usedWrkShts + 1
                    End If
                Next

                If rptType = "Plan on Different Worksheet" Then
                    xlApp.DisplayAlerts = False
                    xlWorkSheetSample.Delete()
                    xlWorkBook.SaveAs(strRptPath, Excel.XlFileFormat.xlExcel7)
                    xlWorkBook.Close()
                Else
                    'If All Plan in one sheet
                    If usedWrkShts > 1 Then
                        'xlApp.Visible = True
                        xlWorkSheetCurrScheme.Cells.Clear()
                        xlWorkSheetCurrScheme.Activate()
                        xlWorkSheet.Range("A:B").Copy()
                        xlWorkSheetCurrScheme.Range("A1").Select()
                        xlWorkSheetCurrScheme.Paste()
                        xlWorkSheetCurrScheme.Range("1:1").Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

                        Dim LastRowCnt As Long
                        LastRowCnt = xlWorkSheetCurrScheme.Cells.Find(What:="*", After:=xlWorkSheetCurrScheme.Cells(2, 1),
                                                 SearchOrder:=Excel.XlSearchOrder.xlByRows,
                                                 SearchDirection:=Excel.XlSearchDirection.xlPrevious).Row

                        FirstColorIndex = True

                        CurrentColName = ""
                        colNumToPaste = 3
                        stColPlanNum = 3
                        endColPlanNum = 3
                        AddNewRow = True
                        CopyStrtColNum = IntFormulaStartcol

                        Dim StrColName2 As String = ""
                        Dim ColMainNum As Integer = 3
                        Dim wrkShtNum As Integer = 0

                        Dim addStartColName As String
                        Dim AddEndColName As String

                        Dim addStartColNameMain As String
                        Dim AddEndColNameMain As String

                        Dim StartPlanNum As Integer = 3
                        Dim EndPlanNum As Integer = 3
                        Dim Num As Long
                        Dim strNum As String
                        'xlApp.Visible = True
                        For index As Integer = startWrkSht + 1 To usedWrkShts
                            xlWorkSheet = xlWorkBook.Worksheets(index)
                            ColMainNum = 3
                            CopyStrtColNum = 3
                            colNumToPaste = 3 + wrkShtNum
                            xlWorkSheet.Range("1:1").Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)

                            'Copy Scheme Total 
                            If wrkShtNum = 0 Then
                                strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                                Num = 2 + objEqualizeTotalSchemeInputCol
                                strNum = objclsExcel.ExcelColName(Num)
                                xlWorkSheet.Range("C:" & strNum).Copy()
                                xlWorkSheetCurrScheme.Activate()
                                xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
                                xlWorkSheetCurrScheme.Paste()
                                xlWorkSheetCurrScheme.Cells(1, 3).value = xlWorkSheetCurrScheme.Name
                                xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, "C1:" & strNum & "1", True)
                            End If
                            '=========================================

                            CopyStrtColNum = objEqualizeTotalSchemeInputCol + 3 '11
                            ColMainNum = objEqualizeTotalSchemeInputCol + 3 '11
                            colNumToPaste = objEqualizeTotalSchemeInputCol + 3 '11


                            If wrkShtNum > 0 Then
                                'colNumToPaste = colNumToPaste + (8 * wrkShtNum)
                                colNumToPaste = colNumToPaste + (objEqualizeTotalInputCol * wrkShtNum)
                                addStartColName = objclsExcel.ExcelColName(colNumToPaste)
                                'addStartColNameMain = objExcel.ExcelColName(colNumToPaste + (7))
                                addStartColNameMain = objclsExcel.ExcelColName(colNumToPaste + (objEqualizeTotalInputCol - 1))
                                xlWorkSheetCurrScheme.Range(addStartColName & ":" & addStartColNameMain).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                                addStartColName = objclsExcel.ExcelColName(ColMainNum)
                                'addStartColNameMain = objExcel.ExcelColName(ColMainNum - 1 + (8 * wrkShtNum))
                                addStartColNameMain = objclsExcel.ExcelColName(ColMainNum - 1 + (objEqualizeTotalInputCol * wrkShtNum))
                                xlWorkSheet.Range(addStartColName & ":" & addStartColNameMain).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                                'CopyStrtColNum = ColMainNum - 1 + (8 * wrkShtNum) + 1
                                CopyStrtColNum = ColMainNum - 1 + (objEqualizeTotalInputCol * wrkShtNum) + 1
                                ColMainNum = CopyStrtColNum
                                colNumToPaste = CopyStrtColNum
                            End If

                            'To Copy INPUT DATA
                            strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                            strColName = objclsExcel.ExcelColName(ColMainNum)
                            CopyStrtColNum = ColMainNum + objEqualizeTotalInputCol - 1
                            StrColName2 = objclsExcel.ExcelColName(CopyStrtColNum)

                            xlWorkSheet.Range(strColName & ":" & StrColName2).Copy()
                            xlWorkSheetCurrScheme.Activate()
                            xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
                            xlWorkSheetCurrScheme.Paste()
                            xlWorkSheetCurrScheme.Cells(1, colNumToPaste).value = xlWorkSheet.Name
                            xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strColName & "1:" & StrColName2 & "1", True)
                            '==============================================


                            CopyStrtColNum = CopyStrtColNum + 1
                            ColMainNum = CopyStrtColNum
                            'colNumToPaste = colNumToPaste + 8 + (wrkShtNum)
                            colNumToPaste = colNumToPaste + objEqualizeTotalInputCol + (wrkShtNum)



                            For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
                                Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
                                If wrkShtNum > 0 And (val = "" Or val = "TRUE" Or val = "1") Then
                                    ColMainNum = ColMainNum + 1
                                    colNumToPaste = colNumToPaste + 1
                                    Continue For
                                End If

                                If wrkShtNum > 0 Then
                                    addStartColName = objclsExcel.ExcelColName(ColMainNum)
                                    addStartColNameMain = objclsExcel.ExcelColName(ColMainNum + 1)

                                    ColMainNum = ColMainNum + wrkShtNum

                                    AddEndColName = objclsExcel.ExcelColName(ColMainNum - 1)
                                    AddEndColNameMain = objclsExcel.ExcelColName(ColMainNum)

                                    xlWorkSheet.Range(addStartColName & ":" & AddEndColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)

                                    strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                                    xlWorkSheetCurrScheme.Range(strPastColName & ":" & strPastColName).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
                                End If

                                strPastColName = objclsExcel.ExcelColName(colNumToPaste)
                                strEndPlan = strPastColName

                                strColName = objclsExcel.ExcelColName(ColMainNum)

                                'Changed
                                xlWorkSheet.Range(strColName & ":" & strColName).Copy()
                                xlWorkSheetCurrScheme.Activate()
                                xlWorkSheetCurrScheme.Range(strPastColName & "1").Select()
                                xlWorkSheetCurrScheme.Paste()
                                xlWorkSheetCurrScheme.Cells(2, colNumToPaste).Value = xlWorkSheet.Name.ToString
                                xlWorkSheetCurrScheme.Cells(1, colNumToPaste).Value = xlWorkSheet.Cells(2, ColMainNum).value




                                ColMainNum = ColMainNum + 1
                                colNumToPaste = colNumToPaste + wrkShtNum + 1
                            Next

                            wrkShtNum = wrkShtNum + 1
                        Next

                        StartPlanNum = 11 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol) + 1
                        'StartPlanNum = 11 + ((usedWrkShts - startWrkSht) * 8)
                        EndPlanNum = StartPlanNum
                        For index1 As Integer = 0 To dtSampleColData.Rows.Count - 1
                            Dim val As String = dtSampleColData.Rows(index1)("ColIsSchemeWise").ToString.Trim.ToUpper
                            If val = "" Or val = "TRUE" Or val = "1" Then
                                EndPlanNum = StartPlanNum
                            Else
                                EndPlanNum = StartPlanNum + (usedWrkShts - startWrkSht - 1)
                            End If

                            strStartPlan = objclsExcel.ExcelColName(StartPlanNum)
                            strEndPlan = objclsExcel.ExcelColName(EndPlanNum)


                            'To add Total column In Worksheet 
                            If val <> "" And val <> "TRUE" And val <> "1" Then
                                val = dtSampleColData.Rows(index1)("ColShowTotal").ToString.Trim.ToUpper()
                                If val = "" Or val = "TRUE" Or val = "1" Then
                                    TotalPlanColNum = EndPlanNum + 1
                                    strTotalPlanCol = objclsExcel.ExcelColName(TotalPlanColNum)
                                    xlWorkSheetCurrScheme.Range(strTotalPlanCol & ":" & strTotalPlanCol).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
                                    xlWorkSheetCurrScheme.Cells(2, strTotalPlanCol).Value = "Total"
                                    xlWorkSheetCurrScheme.Cells(2, strTotalPlanCol).Font.Bold = True
                                    xlWorkSheetCurrScheme.Cells(3, strTotalPlanCol).Value = "=SUM(" & strStartPlan & "3:" & strEndPlan & "3)" '=SUM(C3:C3)
                                    xlWorkSheetCurrScheme.Range(strTotalPlanCol & 3).AutoFill(xlWorkSheetCurrScheme.Range(strTotalPlanCol & "3:" & strTotalPlanCol & LastRowCnt), Excel.XlAutoFillType.xlFillDefault)
                                    xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strTotalPlanCol & "2:" & strTotalPlanCol & LastRowCnt, False)

                                    strEndPlan = strTotalPlanCol
                                    EndPlanNum = TotalPlanColNum
                                End If
                            End If
                            '--------------------------------    

                            If FirstColorIndex Then
                                xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & LastRowCnt).Interior.ColorIndex = 35
                                FirstColorIndex = False
                            Else
                                xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & LastRowCnt).Interior.ColorIndex = 36
                                FirstColorIndex = True
                            End If

                            xlWorkSheetCurrScheme.Range(strStartPlan & "1:" & strEndPlan & "1").Select()
                            xlWorkSheetCurrScheme = FormatRange(xlWorkSheetCurrScheme, strStartPlan & "1:" & strEndPlan & "1", True)
                            StartPlanNum = EndPlanNum + 1
                        Next


                        xlWorkSheetCurrScheme.Columns.AutoFit()
                        'To hide the Row Data
                        StartPlanNum = 3
                        EndPlanNum = 11 + ((usedWrkShts - startWrkSht) * objEqualizeTotalInputCol)
                        'strStartPlan = objExcel.ExcelColName(StartPlanNum)
                        strEndPlan = objclsExcel.ExcelColName(EndPlanNum)
                        xlWorkSheetCurrScheme.Range("C:" & strEndPlan).EntireColumn.Hidden = True

                        For index As Integer = startWrkSht + 1 To usedWrkShts
                            xlWorkSheet = xlWorkBook.Worksheets(startWrkSht + 1)
                            xlWorkSheet.Delete()
                        Next
                    End If

                End If
            End If



        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub
#Region "Vijay"

    Private Sub chkDistributablerpt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDistributablerpt.CheckedChanged
        Try
            If chkDistributablerpt.Checked Then
                pnldistrpt.Visible = True
            Else
                pnldistrpt.Visible = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & ex.Source & ex.StackTrace, "DBEqualization", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub CheckEqualizedSchemes()
        Try
            Dim sb As New System.Text.StringBuilder()
            Dim dtNotEqualize As DataTable

            ' Build Scheme ID list and clear ColChk in one pass
            For Each row As DataGridViewRow In dtDgvSchemeData.Rows
                If Convert.ToBoolean(row.Cells("ColChk").Value) Then
                    sb.Append("'").Append(row.Cells("Scheme_Code").Value.ToString()).Append("',")
                End If
                row.Cells("ColChk").Value = False
            Next

            Dim strSchemeId As String = If(sb.Length > 0, sb.ToString().TrimEnd(","c), "")
            If strSchemeId = "" Then Return

            ' Adjust date range safely
            Dim fromDate As String = DTPFromDate.Value.ToString("dd-MMM-yyyy")
            Dim toDate As Date = dtpToDate.Value
            If DTPFromDate.Value.Day = Date.DaysInMonth(DTPFromDate.Value.Year, DTPFromDate.Value.Month) _
               OrElse toDate.Day = Date.DaysInMonth(toDate.Year, toDate.Month) Then
                toDate = toDate.AddDays(-1)
            End If
            Dim toDateStr As String = toDate.ToString("dd-MMM-yyyy")

            ' Fetch once
            dtNotEqualize = objEqualize_clsDAL.GetApprovedData("Get Not Equalized Scheme", fromDate, toDateStr, cmbMFund.SelectedValue, strSchemeId)

            If dtNotEqualize IsNot Nothing AndAlso dtNotEqualize.Rows.Count > 0 Then
                ' Build a fast lookup dictionary
                Dim map = dtNotEqualize.AsEnumerable().ToDictionary(Function(r) r.Field(Of String)("Scheme Code"),
                                  Function(r) r.Field(Of Object)("Not Equalized ON").ToString())

                ' Update the list-view items
                For i As Integer = 0 To lstEqulizeStatus.Items.Count - 1
                    Dim code = lstEqulizeStatus.Items(i).SubItems(0).Text
                    If map.ContainsKey(code) Then
                        If lstEqulizeStatus.Items(i).SubItems(1).Text.Contains("completed successfully") Then
                            lstEqulizeStatus.Items(i).SubItems(1).Text =
                                $"Data Equalization Not completed – date: {map(code)}"
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            MsgBox($"Error Source: {ex.Source}{vbCrLf}Message: {ex.Message}", MsgBoxStyle.Information, Me.Text)
        End Try
    End Sub
    'Private Sub CheckEqualizedSchemes()
    '    Dim StrSchemeId As String = ""
    '    Dim DtNotEqualize As New DataTable
    '    Dim dv As DataView
    '    Try
    '        For index As Integer = 0 To dtDgvSchemeData.Rows.Count - 1
    '            If dtDgvSchemeData.Rows(index)("ColChk") Then
    '                StrSchemeId = StrSchemeId & "'" & dtDgvSchemeData.Rows(index)("Scheme_Code").ToString & "',"
    '            End If
    '        Next
    '        If StrSchemeId <> "" Then
    '            StrSchemeId = StrSchemeId.Substring(0, Len(StrSchemeId) - 1)
    '            'DtNotEqualize = objEqualize_clsDAL.GetApprovedData("Get Not Equalized Scheme", DTPFromDate.Text, dtpToDate.Text, cmbMFund.SelectedValue, StrSchemeId)

    '            'START Added by Dattatray
    '            Dim strFromDate As String = DTPFromDate.Text
    '            Dim strToDate As String = dtpToDate.Text

    '            If (DTPFromDate.Value.Day = Date.DaysInMonth(DTPFromDate.Value.Year, DTPFromDate.Value.Month) And dtpToDate.Value.Day = Date.DaysInMonth(dtpToDate.Value.Year, dtpToDate.Value.Month)) _
    '                Or dtpToDate.Value.Day = Date.DaysInMonth(dtpToDate.Value.Year, dtpToDate.Value.Month) Then

    '                strToDate = Format(DateAdd(DateInterval.Day, -1, dtpToDate.Value), "dd-MMM-yyyy")
    '            End If
    '            DtNotEqualize = objEqualize_clsDAL.GetApprovedData("Get Not Equalized Scheme", strFromDate, strToDate, cmbMFund.SelectedValue, StrSchemeId)
    '            'END Added by Dattatray
    '        End If

    '        If DtNotEqualize.Rows.Count > 0 Then
    '            For i As Integer = 0 To lstEqulizeStatus.Items.Count - 1
    '                dv = New DataView(DtNotEqualize)
    '                dv.RowFilter = "[Scheme Code]='" & lstEqulizeStatus.Items(i).SubItems(0).Text & "'"
    '                If dv.ToTable.Rows.Count > 0 Then
    '                    If lstEqulizeStatus.Items(i).SubItems(1).Text.Contains("Data Equalization completed successfully") Then
    '                        lstEqulizeStatus.Items(i).SubItems(1).Text = "Data Equalization Not completed - date: " & dv.ToTable.Rows(0)("Not Equalized ON").ToString()
    '                    End If
    '                End If
    '            Next
    '        End If

    '        For index As Integer = 0 To dtDgvSchemeData.Rows.Count - 1
    '            dtDgvSchemeData.Rows(index)("ColChk") = False
    '        Next
    '    Catch ex As Exception
    '        MsgBox("Error Source : " & ex.Source & vbCrLf & "Error Message : " & ex.Message & vbCrLf & "Error Occured in Method:-" & vbCrLf & ex.StackTrace, MsgBoxStyle.Information, Me.Text)
    '    End Try
    'End Sub
#End Region

End Class
