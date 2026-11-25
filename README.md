
## SOLIDWORKS API & VBA Resources

A curated collection of open-source SOLIDWORKS automation projects, VBA macros, API examples, and engineering tools.

The primary target environment is **SOLIDWORKS 2019**, with a focus on reusable VBA code and backward-compatible API implementations.

### 1. VBA Macro Libraries

| Project | Main Features | Repository |
|---|---|---|
| CodeStack | SOLIDWORKS API, VBA macros, Drawing, Assembly, Properties | https://github.com/xarial/codestack |
| SOLIDWORKS Macro Library | PDF, DXF, STEP, BOM, batch export | https://github.com/SandBasher/solidworks-automation |
| Custom Properties Manager | Excel-based batch property management | https://github.com/ramugopal92/solidworks-custom-properties-manager |
| AutoDrawing | Automatic drawing views, dimensions and scale | https://github.com/kilwizac/solidworks-automate-drawing |
| Assembly Export | Recursive assembly traversal and batch export | https://github.com/mmmlinux/solidworks-export-macro |
Dim swApp As SldWorks.SldWorks
Dim swModel As ModelDoc2
Dim swDraw As DrawingDoc
Dim swSheet As Sheet
Dim swDrawCustPropMgr As CustomPropertyManager
Dim swModelConfigMgr As CustomPropertyManager
Dim swConfigMgr As ConfigurationManager
Dim swConfig As Configuration
Dim swOpenModel As ModelDoc2
Dim swModelView As ModelView
Dim swView As View
Dim swTable As TableAnnotation
Dim swBOMTable As IBomTableAnnotation
Dim swAnn As Annotation
Dim swNote As Note
Dim swFeat As Feature
Dim swFeatFolder As FeatureFolder
Dim swFtrFolder As Feature
Dim swComp As Component2
Dim swAsm As AssemblyDoc
Dim TheMouse As Mouse
Dim swDispDim As DisplayDimension
Dim swDim As Dimension
Dim swCalloutStringVariable As CalloutStringVariable
Dim swCalloutLengthVariable As CalloutLengthVariable
Dim swCalloutVariable As CalloutVariable
Dim swDisplayData As DisplayData
Dim swRootComp As Component2
Dim swTextFormat As TextFormat
Dim swSelMgr As SelectionMgr
Dim swSketchMgr As SketchManager
Dim swSketchSegment As SketchSegment
Dim swMassProperty As MassProperty
Dim xlApp As Object
Dim xlWorkbook As Object
Dim xlWorkSheet As Object
Dim MousePosition As New Class_ClickMouse
Dim EndProgram As New Class_ClickESC
Dim swExportPDFData As ExportPdfData
Dim swExportSTPData As ExportPdfData
Dim swRevTable As SldWorks.RevisionTableAnnotation
Dim strThreadSize As String
Dim strValue() As String
Dim strPoCode As String
Dim strSoMaySoCum As String
Dim strDrawingNo As String
Dim strExtension As String
Dim strPartName As String
Dim strPathName As String
Dim strMaterial As String
Dim strPartNameTopLevel_Temp As String
Dim strPartNameTopLevel_Last As String
Dim strStandardDevice() As String
Dim strSemiStandardDevice() As String
Dim strMechanicalStandardDevice() As String
Dim strElectricalStandardDevice() As String
Dim strHydraulics_PneumaticsStandardDevice() As String
Dim strBoltStandardDevice() As String
Dim strCheckPartname() As String
Dim strConfigurationName As String
Dim strCheckConfigurationName() As String
Dim strFilePath As String
Dim strFileName As String
Dim strFilter As String
Dim strThread As String
Dim strGetPrefixTextDefinition As String
Dim strGetSuffixTextDefinition As String
Dim strGetCalloutAboveDefinition As String
Dim strGetCalloutBelow As String
Dim strConfigName As String
Dim strExport() As String
Dim strPartNameRecovery() As String
Dim strDrawingNoRecovery() As String
Dim strMaterialRecovery() As String
Dim strTreatmentRecovery() As String
Dim strRemarkRecovery() As String
Dim strFinishRecovery() As String
Dim strQtyRecovery() As String
Dim strClassRecovery() As String
Dim strRoughnessRecovery() As String
Dim strConfigurationNameRecovery() As String
Dim strFileNameRecovery() As String
Dim strGetSubFeatFolder() As String
Dim strGetParentCompSubFeat() As String
Dim strGetMass As String
Dim strFinish As String
Dim strCheckExtension() As String
Dim strGetCurrentPath As String
Dim strGetAddNoteInsertThread As String
Dim strRevision As String
Dim pathSTP3 As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim n As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim intSheetNumber As Integer
Dim intCountDuplicateSheet As Integer
Dim intDuplicatePosition() As Integer
Dim intInitializedPosition() As Integer
Dim intActiveSheetPosition() As Integer
Dim intCountEmptySheet As Integer
Dim intHorizontalMarkPosition As Integer
Dim intCheckExistPartName As Integer
Dim intFrequenceUnequal As Integer
Dim intAddSheet As Integer
Dim iMemory2 As Integer
Dim intInsertThreadPosition As Integer
Dim intHoleDepthPosition As Integer
Dim intRow As Integer
Dim intFirstSlashesPosition As Integer
Dim intLastSlashesPosition As Integer
Dim intCountConfigNameDefault As Integer
Dim intCountStandardDevice As Integer
Dim intCountSemiStandardDevice As Integer
Dim intCountMechanicalStandardDevice As Integer
Dim intCountElectricalStandardDevice As Integer
Dim intCountHydraulics_PneumaticsStandardDevice As Integer
Dim intCountBoltStandardDevice As Integer
Dim intCountChildrenStandardDevice As Integer
Dim intCountAllComponent As Integer
Dim intCheckDuplicateSheet As Integer
Dim intGetRange() As Integer
Dim vSheetNames As Variant
Dim vDrawCustPropNames As Variant
Dim vModelConfigMgr As Variant
Dim vModelCustPropMgr As Variant
Dim dbNotePosition As Variant
Dim vNote As Variant
Dim vComps As Variant
Dim vGetScaleRatio As Variant
Dim vFeatures As Variant
Dim vconfigNames As Variant
Dim vconfigName As Variant
Dim vGetSheetProperties As Variant
Dim blnStatusSetDrawingNo As Boolean
Dim blnCheckWeldment_1 As Boolean
Dim blnCheckWeldment_2 As Boolean
Dim blnCheckWeldment_3 As Boolean
Dim dictWeldmentChildParts As Object
Dim blnCheckWeldmentChild As Boolean
Dim blnCheckOpenModel3D As Boolean
Dim blnCheckEditingDrawing As Boolean
Dim blnCheckTBTC As Boolean
Dim blnStatusNameFirstSheet As Boolean
Dim blnCheckLastHorizontalMark As Boolean
Dim blnCheckFirstSlashes As Boolean
Dim blnStatusExportExcel As Boolean
Dim blnStatusBackup As Boolean
Dim blnStatusAutoSaved As Boolean
Dim flag As Boolean
Dim statusExportCAD2D As Boolean
Dim blnStatusMainLoop As Boolean
Dim blnStatusExportPDF As Boolean
Dim blnInsertThreadMaterial As Boolean
Dim blnCheckDepthThru As Boolean
Dim blnAutoRemoveExistDrawingNo As Boolean
Dim blnHoleTable As Boolean
Dim blnExportExcel As Boolean
Dim blnExportCAD2D As Boolean
Dim blnStatusEmptyRow As Boolean
Dim blnCheckMissingBalloon As Boolean
Dim singleDensity As Single
Dim strThreadDepth As Single
Dim singleInsertThread As Single
Dim dblMassTotal As Double
Dim dblVolume As Double
Dim dblMass As Double
Dim dblSurfaceArea As Double
Dim DisplayWidth As Long
Dim DisplayHeight As Long
Dim Xposition As Long

Private Sub chbAllInOne_Click()
    If chbAllInOne = True Then
        opbtDrawingNoExportExcel = True
        opbtAllSheetsExportExcel = True
    Else
        opbtDrawingNoExportExcel = True
    End If
End Sub

Private Sub CheckBox1_Click()
End Sub

Private Sub cmdAUTOPARTLIST2D_Click()
    Call CheckRun
    blnStatusMainLoop = True
    intCountDuplicateSheet = 0
    intCountEmptySheet = 0
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Call SetNameFirstSheet
    If blnStatusNameFirstSheet = True Then Exit Sub
    flag = False
    If opbtCurrentSheetAutoPartlist = False Then
        Call GetRange(strRangeAutoPartlist, CInt(UBound(vSheetNames) + txtStartRenameSheet - 1))
    End If
    AUTOPARTLIST.Hide
    If flag = False Then
        Call MainLoop
        If opbtCurrentSheetAutoPartlist = False Then
            Call DetermineLoop2
        Else
            If opbtAutoSort = True Or opbtAutoNote = True Then
                If swApp.SendMsgToUser2("Have you sure exported all drawings of this assembly?", swMbWarning, swMbYesNo) = swMbHitYes Then
                    Call Loop2
                Else
                    End
                End If
            End If
            blnStatusMainLoop = False
            If blnStatusSetDrawingNo Then swApp.SendMsgToUser2 "AUTO PARTLIST COMPLETED!", swMbInformation, swMbOk
        End If
    End If
    If opbtQuit_wd = True Then
        Call Quit_word_Click
    End If
    Unload Me
End Sub

Private Sub MainLoop()
    ReDim strCheckPartname(UBound(vSheetNames)) As String
    ReDim strCheckConfigurationName(UBound(vSheetNames)) As String
    ReDim strCheckExtension(UBound(vSheetNames)) As String
    ReDim intDuplicatePosition(UBound(vSheetNames)) As Integer
    ReDim intInitializedPosition(UBound(vSheetNames)) As Integer
    ReDim intActiveSheetPosition(UBound(vSheetNames)) As Integer
    swApp.FrameState = swWindowMaximized
    Set EndProgram = New Class_ClickESC
    On Error Resume Next
    swModel.Extension.LoadDraftingStandard swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsDraftingStandard) & "\ISO-TCI.sldstd"
    swModel.Extension.SetUserPreferenceToggle swUserPreferenceIntegerValue_e.swDetailingDimensionStandard, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swDetailingStandard_e.swDetailingStandardUserDefined
    intCheckDuplicateSheet = 0
    Call ScanWeldedAssemblies
    If opbtCurrentSheetAutoPartlist = True Then
        Dim strCurrentTargetSheet As String
        strCurrentTargetSheet = swDraw.GetCurrentSheet.GetName
        For intSheetNumber = 0 To UBound(vSheetNames)
            If vSheetNames(intSheetNumber) = strCurrentTargetSheet Then
                Call ProcessAutoPartlist(intSheetNumber)
                Exit For
            End If
        Next intSheetNumber
    ElseIf opbtAllSheetsAutoPartlist = True Then
        For intSheetNumber = 0 To UBound(vSheetNames)
            Call ProcessAutoPartlist(intSheetNumber)
        Next intSheetNumber
    ElseIf opbtRangeAutoPartlist = True Then
        Call GetRange(strRangeAutoPartlist, CInt(UBound(vSheetNames) + txtStartRenameSheet - 1))
        For a = 0 To iMemory2 - 1
            intSheetNumber = CInt(intGetRange(a) - txtStartRenameSheet)
            Call ProcessAutoPartlist(intSheetNumber)
        Next a
    End If
    'Delete duplicate sheets
    If blnStatusMainLoop = True And opbtCurrentSheetAutoPartlist = False Then
        If intCountDuplicateSheet > 0 Then
            'Delete excess message box
            For i = 0 To intCountDuplicateSheet - 1
                For j = 0 To intCountDuplicateSheet - 1
                    If intDuplicatePosition(i) = intInitializedPosition(j) Then
                        For k = j To intCountDuplicateSheet - 1
                            intDuplicatePosition(k) = intDuplicatePosition(k + 1)
                            intInitializedPosition(k) = intInitializedPosition(k + 1)
                        Next k
                        intCountDuplicateSheet = intCountDuplicateSheet - 1
                    End If
                Next j
            Next i
            For i = 0 To intCountDuplicateSheet - 1
                MsgBox "Error: Duplicate sheet " & (intDuplicatePosition(i) + txtStartRenameSheet) & " = sheet " & (intInitializedPosition(i) + txtStartRenameSheet) & " !", vbCritical, "WARNING!"
            Next i
            vSheetNames = swDraw.GetSheetNames()
            If swApp.SendMsgToUser2("Do you want to delete duplicate sheets?", swMbStop, swMbYesNo) = swMbHitYes Then
                swApp.SendMsgToUser2 "This program will delete duplicate sheet automatically and auto run AUTO PARTLIST 2D again!", swMbWarning, swMbOk
                swDraw.ActivateSheet (vSheetNames(0))
                For i = 0 To intCountDuplicateSheet - 1
                    swDraw.Extension.SelectByID2 vSheetNames(intDuplicatePosition(i)), "SHEET", 0, 0, 0, False, 0, Nothing, swSelectOptionDefault
                    swDraw.Extension.DeleteSelection2 (0)
                Next i
            Else
                End
            End If
        End If
    End If
    'Write info total sheet in drawing custom properties
    If opbtTotalSheets = True Then
        ReDim strValue(swDrawCustPropMgr.Count) As String
        For i = 0 To UBound(vDrawCustPropNames)
            swDrawCustPropMgr.Get3 vDrawCustPropNames(i), False, strValue(i), Empty
            If UCase(vDrawCustPropNames(i)) = "TOTAL SHEET" Then swDrawCustPropMgr.Set "TOTAL SHEET", UBound(vSheetNames) + txtStartRenameSheet
        Next i
    End If
    'Rename first sheet after remove duplicate sheet
    If opbtRangeAutoPartlist = True And (intCountDuplicateSheet > 0 Or intCountEmptySheet > 0) And blnStatusMainLoop = False Then
        swDraw.ActivateSheet (vSheetNames(0))
        swDraw.GetCurrentSheet.SetName txtStartRenameSheet
    End If
    swModel.EditRebuild3
    Call TurnOnAutoSaved
    Exit Sub
End Sub

Private Sub ProcessAutoPartlist(intSheetNumber As Integer)
    EndProgram.EndByESCKey
    swDraw.ActivateSheet (vSheetNames(intSheetNumber))
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    If Not swView Is Nothing Then
        strConfigurationName = swView.ReferencedConfiguration
        strPathName = swView.GetReferencedModelName
LoopClickMouseAutoPartlist:
        Call OpenModel(strPathName)
        blnStatusSetDrawingNo = False
        If Not swModelConfigMgr Is Nothing Then
            For i = 0 To swModelConfigMgr.Count
                On Error Resume Next
                Select Case UCase(vModelConfigMgr(i))
                Case "PART NAME"
                    Call InfoPartname(swOpenModel)
                Case "DRAWING NO."
                    Call InfoDrawingNo
                    swModelConfigMgr.Set "DRAWING NO.", strDrawingNo
                    blnStatusSetDrawingNo = True
                    If opbtQRAdd = True Then 'Neu tick tao QR thì
                        If opbtCheckInternet = True Then
                            Call DelQR_offline
                            Call AddQR
                            Call AddQR2
                        ElseIf opbtCheckInternet = False Then
                            Call Del_QR_RS
                            Call QR_offline_Click 'Call AddQR
                            End If
                    End If
'"""
                Case "REMARK"
                    swModelConfigMgr.Set "REMARK", GetWeldingRemark(strPathName, strExtension)
                End Select
            Next i
        End If
        Call InfoMass
        'Call RemoveUnsuitFinish
        EndProgram.EndByESCKey
        'Check position duplicate sheet
        If blnStatusMainLoop = True Then
            strCheckPartname(intCheckDuplicateSheet) = strPartName
            strCheckConfigurationName(intCheckDuplicateSheet) = strConfigurationName
            strCheckExtension(intCheckDuplicateSheet) = strExtension
            intActiveSheetPosition(intCheckDuplicateSheet) = intSheetNumber
            If intCheckDuplicateSheet > 0 Then
                For i = 0 To intCheckDuplicateSheet - 1
                    If strPartName <> "" And strPartName = strCheckPartname(i) And strConfigurationName = strCheckConfigurationName(i) And strExtension = strCheckExtension(i) Then
                        intDuplicatePosition(intCountDuplicateSheet) = intSheetNumber
                        intInitializedPosition(intCountDuplicateSheet) = intActiveSheetPosition(i)
                        intCountDuplicateSheet = intCountDuplicateSheet + 1
                    End If
                Next i
            End If
            intCheckDuplicateSheet = intCheckDuplicateSheet + 1
        End If
        If blnStatusSetDrawingNo = False And blnCheckOpenModel3D = True Then
            If Xposition <= DisplayWidth Then
                TheMouse.Move DisplayWidth * 2 - 260, 160, swMouse_MouseMove
                GoTo LoopClickMouseAutoPartlist
            Else
                TheMouse.Move DisplayWidth - 260, 160, swMouse_MouseMove
                GoTo LoopClickMouseAutoPartlist
            End If
        End If
        'Close model 3D if having more one configuration
        Call SaveAndCloseModel3D(strPathName)
    Else
    'Delete empty sheet
        swDraw.Extension.SelectByID2 vSheetNames(intSheetNumber), "SHEET", 0, 0, 0, False, 0, Nothing, 0
        swDraw.Extension.DeleteSelection2 (0)
        intCountEmptySheet = intCountEmptySheet + 1
    End If
    'Export info for flat-pattern view
    Dim strTreatment As String
    Dim strQty As String
    Dim strClass As String
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    If Not swView Is Nothing Then
        If swView.IsFlatPatternView() Then
            Set swView = swView.GetNextView
            Do While Not swView Is Nothing
                If swView.IsFlatPatternView() = False Then
                    Call OpenModel(strPathName)
                    If Not swModelConfigMgr Is Nothing Then
                        For i = 0 To swModelConfigMgr.Count
                            On Error Resume Next
                            Select Case UCase(vModelConfigMgr(i))
                            Case "PART NAME"
                                swModelConfigMgr.Set "PART NAME", UCase(strPartName)
                            Case "DRAWING NO."
                                swModelConfigMgr.Set "DRAWING NO.", strDrawingNo
                            Case "MATERIAL"
                                swModelConfigMgr.Get3 "MATERIAL", False, Empty, strMaterial
                            Case "TREATMENT"
                                swModelConfigMgr.Get3 "TREATMENT", False, Empty, strTreatment
                            Case "FINISH"
                                swModelConfigMgr.Get3 "FINISH", False, Empty, strFinish
                            Case "QTY"
                                swModelConfigMgr.Get3 "QTY", False, Empty, strQty
                            Case "CLASS"
                                swModelConfigMgr.Get3 "CLASS", False, Empty, strClass
                            End Select
                        Next i
                    End If
                    Exit Do
                End If
                Set swView = swView.GetNextView
            Loop
        End If
    End If
    Call SaveAndCloseModel3D(strPathName)
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    If Not swView Is Nothing Then
        If swView.IsFlatPatternView() Then
            Call OpenModel(strPathName)
            If Not swModelConfigMgr Is Nothing Then
                For i = 0 To swModelConfigMgr.Count
                    On Error Resume Next
                    Select Case UCase(vModelConfigMgr(i))
                    Case "MATERIAL"
                        If strMaterial <> "" Then swModelConfigMgr.Set "MATERIAL", strMaterial
                    Case "TREATMENT"
                        If strTreatment <> "" Then swModelConfigMgr.Set "TREATMENT", strTreatment
                    Case "FINISH"
                        If strFinish <> "" Then swModelConfigMgr.Set "FINISH", strFinish
                    Case "QTY"
                        If strQty <> "" And strQty <> "0" Then swModelConfigMgr.Set "QTY", strQty
                    Case "CLASS"
                        If strClass <> "" Then swModelConfigMgr.Set "CLASS", strClass
                    End Select
                Next i
            End If
            Call SaveAndCloseModel3D(strPathName)
        End If
    End If
    'Check status set Drawing No
    If blnStatusSetDrawingNo = False And intCountEmptySheet = 0 Then
        swApp.SendMsgToUser2 "This program recognizes that you haven't set APPLY button in Custom Property of the model 3D yet!", swMbWarning, swMbOk
        End
    End If
    If intCountEmptySheet = 0 Then
        Call SetSheetScaleCAD2D
        'Edit BOM cut list
        If strExtension = "SLDPRT" Then Call EditBOMCutList
        'Auto Insert thread
        If opbtAutoInsertThread = True Then Call INSERT_THREAD
        'Auto Po Code for external component
        If opbtAutoPoCodeForExternalComponent = True Then Call AutoPoCodeForExternalComponent
        'Rename sheet
        If InStr(1, txtStartRenameSheet, "0", vbTextCompare) = 1 And intSheetNumber + txtStartRenameSheet < 10 Then
            swDraw.GetCurrentSheet.SetName "0" & intSheetNumber + txtStartRenameSheet & " "
            swDraw.GetCurrentSheet.SetName "0" & CStr(intSheetNumber + txtStartRenameSheet)
        Else
            swDraw.GetCurrentSheet.SetName "" & intSheetNumber + txtStartRenameSheet & " "
            swDraw.GetCurrentSheet.SetName CStr(intSheetNumber + txtStartRenameSheet)
        End If
        'Delete Table
        If strExtension = "SLDASM" Then
            'Call DeleteHoleTableAndRoughnessBox
        End If
        'Refresh partlist
        If strExtension = "SLDASM" Then
            Call RefreshPartlist
            If blnStatusEmptyRow = False Then
                Call AutoSortPartlist
                Call DeleteEmptyRowAndEditItemNumber
            End If
        End If
    End If
End Sub

Private Sub InfoDrawingNo()
    strRevision = ""
    Set swRevTable = swDraw.GetCurrentSheet.RevisionTable
    If Not swRevTable Is Nothing Then
        strRevision = swRevTable.CurrentRevision
    End If
    If InStr(1, txtStartRenameSheet, "0", vbTextCompare) = 1 And intSheetNumber + txtStartRenameSheet < 10 Then
        If strRevision = "" Then
            strDrawingNo = UCase(strPoCode) & "-" & strSoMaySoCum & "-0" & intSheetNumber + txtStartRenameSheet
        Else
            strDrawingNo = UCase(strPoCode) & "-" & strSoMaySoCum & "-0" & intSheetNumber + txtStartRenameSheet & "-R" & strRevision
        End If
    Else
        If strRevision = "" Then
            strDrawingNo = UCase(strPoCode) & "-" & strSoMaySoCum & "-" & intSheetNumber + txtStartRenameSheet
        Else
            strDrawingNo = UCase(strPoCode) & "-" & strSoMaySoCum & "-" & intSheetNumber + txtStartRenameSheet & "-R" & strRevision
        End If
    End If
End Sub

Private Sub RefreshPartlist()
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotation_BillOfMaterials Then
            'Set header on bottom
            swTable.SetHeader swTableHeaderPosition_e.swTableHeader_Bottom, 1
            'Check empty row
            blnStatusEmptyRow = False
            For i = 0 To swTable.RowCount - 3
                If swTable.Text(i, 2) = Empty Then blnStatusEmptyRow = True
            Next i
            If chbAutoRefreshPartlist = True Then
                'Delete current BOM table
                Set swAnn = swTable.GetAnnotation
                swAnn.Select3 False, Nothing
                swModel.EditDelete
                'Insert new BOM table
                Set swView = swView.GetNextView
                Set swSheet = swDraw.GetCurrentSheet()
                vGetSheetProperties = swSheet.GetProperties2()
                If vGetSheetProperties(0) = 8 Then 'Size A3
                    On Error Resume Next
                    Set swBOMTable = swView.InsertBomTable4(False, 0.415, 0.04, swBOMConfigurationAnchor_BottomRight, swBomType_TopLevelOnly, swView.ReferencedConfiguration, swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsBOMTemplates) & "\PARTLIST TCI.sldbomtbt", False, Empty, Empty)
                ElseIf vGetSheetProperties(0) = 9 Then 'Size A2
                    On Error Resume Next
                    Set swBOMTable = swView.InsertBomTable4(False, 0.584, 0.045, swBOMConfigurationAnchor_BottomRight, swBomType_TopLevelOnly, swView.ReferencedConfiguration, swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsBOMTemplates) & "\PARTLIST TCI.sldbomtbt", False, Empty, Empty)
                ElseIf vGetSheetProperties(0) = 10 Then 'Size A1
                    On Error Resume Next
                    Set swBOMTable = swView.InsertBomTable4(False, 0.836, 0.04, swBOMConfigurationAnchor_BottomRight, swBomType_TopLevelOnly, swView.ReferencedConfiguration, swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsBOMTemplates) & "\PARTLIST TCI.sldbomtbt", False, Empty, Empty)
                ElseIf vGetSheetProperties(0) = 11 Then 'Size A0
                    On Error Resume Next
                    Set swBOMTable = swView.InsertBomTable4(False, 1.184, 0.04, swBOMConfigurationAnchor_BottomRight, swBomType_TopLevelOnly, swView.ReferencedConfiguration, swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsBOMTemplates) & "\PARTLIST TCI.sldbomtbt", False, Empty, Empty)
                End If
                'Edit a title of new BOM
                Set swView = swDraw.GetFirstView
                Set swTable = swView.GetFirstTableAnnotation
                While Not swTable Is Nothing
                    If swTable.Type = swTableAnnotation_BillOfMaterials Then
                        If swBOMTable.BomFeature.Configuration <> "Default" Then
                            swTable.Text(swTable.RowCount - 2, 3) = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG"
                            Set swTextFormat = swTable.GetCellTextFormat(0, 0)
                            swTable.SetCellTextFormat swTable.RowCount - 2, 3, True, swTextFormat
                            swTable.SetColumnWidth 3, 16.5 / 1000, 0
                        End If
                    End If
                    Set swTable = swTable.GetNext
                Wend
            End If
            'Get name of current BOM
            Set swView = swDraw.GetFirstView
            Set swView = swView.GetNextView
            Dim strGetKeepLinkedtoBOMName As String
            While Not swView Is Nothing
                If swView.GetKeepLinkedToBOMName <> "" Then strGetKeepLinkedtoBOMName = swView.GetKeepLinkedToBOMName
                Set swView = swView.GetNextView
            Wend
            'Link all views with current BOM
            Set swView = swDraw.GetFirstView
            Set swView = swView.GetNextView
            While Not swView Is Nothing
                swView.SetKeepLinkedToBOM True, strGetKeepLinkedtoBOMName
                Set swView = swView.GetNextView
            Wend
        End If
        Set swTable = swTable.GetNext
    Wend
End Sub

Private Sub DeleteEmptyRowAndEditItemNumber()
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotation_BillOfMaterials Then
            'Delete empty row
            Set swAnn = swTable.GetAnnotation
            swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationHidden
            For i = 0 To swTable.RowCount - 3
               For j = 0 To swTable.RowCount - 3
                    On Error Resume Next
                    If swTable.Text(j, 2) = Empty Then
                        swTable.DeleteRow (j)
                    End If
               Next j
            Next i
            'Edit item number ascending in BOM
            On Error Resume Next
            Set swTextFormat = swTable.GetCellTextFormat(0, swTable.RowCount - 2)
            For i = 0 To swTable.RowCount - 3
                swTable.Text(i, 0) = ""
                On Error Resume Next
                swTable.Text(i, 0) = swTable.RowCount - 2 - i
                swTable.SetCellTextFormat i, 0, True, swTextFormat
            Next i
            swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationVisible
        End If
        Set swTable = swTable.GetNext
    Wend
End Sub

Private Sub DeleteHoleTableAndRoughnessBox()
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotation_BillOfMaterials Then
            swModel.EditTemplate
            Set swSheet = swDraw.GetCurrentSheet()
            vGetSheetProperties = swSheet.GetProperties2()
            Call SelectRoughnessBox
            swModel.Extension.DeleteSelection2 (0) 'Delete Roughness Box
            Call SelectHoleTableBox
            swModel.Extension.DeleteSelection2 (0) 'Delete HoleTable Box
            swModel.EditSheet
        End If
        Set swTable = swTable.GetNext
    Wend
End Sub

Private Sub EditBOMCutList()
    swModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swWeldmentEnableAutomaticCutList, True
    swModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swWeldmentEnableAutomaticUpdate, True
    swModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swDisableDerivedConfigurations, True
    swModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swDisableWeldmentConfigStrings, True
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotation_WeldmentCutList Then
            If chbAutoRefreshPartlist = True Then
                'Delete current BOM table
                Set swAnn = swTable.GetAnnotation
                swAnn.Select3 False, Nothing
                swModel.EditDelete
                'Insert new BOM table
                Set swView = swView.GetNextView
                Set swSheet = swDraw.GetCurrentSheet()
                vGetSheetProperties = swSheet.GetProperties2()
                If vGetSheetProperties(0) = 8 Then 'Size A3
                    On Error Resume Next
                    swView.InsertWeldmentTable False, 0.415, 0.04, swBOMConfigurationAnchor_BottomRight, "Default<As Welded>", swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsWeldmentCutListTemplates) & "\WELDMENT CUT LIST TCI.sldwldtbt"
                ElseIf vGetSheetProperties(0) = 9 Then 'Size A2
                    On Error Resume Next
                    swView.InsertWeldmentTable False, 0.584, 0.045, swBOMConfigurationAnchor_BottomRight, "Default<As Welded>", swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsWeldmentCutListTemplates) & "\WELDMENT CUT LIST TCI.sldwldtbt"
                ElseIf vGetSheetProperties(0) = 10 Then 'Size A1
                    On Error Resume Next
                    swView.InsertWeldmentTable False, 0.836, 0.04, swBOMConfigurationAnchor_BottomRight, "Default<As Welded>", swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsWeldmentCutListTemplates) & "\WELDMENT CUT LIST TCI.sldwldtbt"
                ElseIf vGetSheetProperties(0) = 11 Then 'Size A0
                    On Error Resume Next
                    swView.InsertWeldmentTable False, 1.184, 0.04, swBOMConfigurationAnchor_BottomRight, "Default<As Welded>", swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsWeldmentCutListTemplates) & "\WELDMENT CUT LIST TCI.sldwldtbt"
                End If
            End If
            'Edit table
            Set swView = swDraw.GetFirstView
            Set swTable = swView.GetFirstTableAnnotation
            While Not swTable Is Nothing
                If swTable.Type = swTableAnnotation_WeldmentCutList Then
                    Set swTextFormat = swTable.GetCellTextFormat(0, swTable.RowCount - 2)
                    Set swAnn = swTable.GetAnnotation
                    swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationHidden
                    For i = 0 To swTable.RowCount - 3
                        'Info of second column
                        If InStr(1, swTable.Text(i, 4), "AL", vbTextCompare) Then
                            swTable.Text(i, 1) = "THANH NH" & ChrW(212) & "M " & (swTable.RowCount - 2 - i)
                        Else
                            swTable.Text(i, 1) = "THANH " & (swTable.RowCount - 2 - i)
                        End If
                        swTable.SetCellTextFormat i, 1, False, swTextFormat
                    Next i
                    swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationVisible
                End If
                Set swTable = swTable.GetNext
            Wend
        ElseIf swTable.Type = swTableAnnotation_BillOfMaterials Then
            If chbAutoRefreshPartlist = True Then
                'Delete current BOM table
                Set swAnn = swTable.GetAnnotation
                swAnn.Select3 False, Nothing
                swModel.EditDelete
                'Insert new BOM table
                Set swView = swView.GetNextView
                Set swSheet = swDraw.GetCurrentSheet()
                vGetSheetProperties = swSheet.GetProperties2()
                If vGetSheetProperties(0) = 8 Then 'Size A3
                    On Error Resume Next
                    swView.InsertBomTable4 False, 0.415, 0.04, swBOMConfigurationAnchor_BottomRight, swBomType_e.swBomType_Indented, swView.ReferencedConfiguration, swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsBOMTemplates) & "\BOM CUT LIST TCI.sldbomtbt", False, swNumberingType_Flat, True
                ElseIf vGetSheetProperties(0) = 9 Then 'Size A2
                    On Error Resume Next
                    swView.InsertBomTable4 False, 0.584, 0.045, swBOMConfigurationAnchor_BottomRight, swBomType_e.swBomType_Indented, swView.ReferencedConfiguration, swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsBOMTemplates) & "\BOM CUT LIST TCI.sldbomtbt", False, swNumberingType_Flat, True
                ElseIf vGetSheetProperties(0) = 10 Then 'Size A1
                    On Error Resume Next
                    swView.InsertBomTable4 False, 0.836, 0.04, swBOMConfigurationAnchor_BottomRight, swBomType_e.swBomType_Indented, swView.ReferencedConfiguration, swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsBOMTemplates) & "\BOM CUT LIST TCI.sldbomtbt", False, swNumberingType_Flat, True
                ElseIf vGetSheetProperties(0) = 11 Then 'Size A0
                    On Error Resume Next
                    swView.InsertBomTable4 False, 1.184, 0.04, swBOMConfigurationAnchor_BottomRight, swBomType_e.swBomType_Indented, swView.ReferencedConfiguration, swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsBOMTemplates) & "\BOM CUT LIST TCI.sldbomtbt", False, swNumberingType_Flat, True
                End If
            End If
            'Get name of current BOM
            Set swView = swDraw.GetFirstView
            Set swView = swView.GetNextView
            While Not swView Is Nothing
                Dim strGetKeepLinkedtoBOMName As String
                If swView.GetKeepLinkedToBOMName <> "" Then strGetKeepLinkedtoBOMName = swView.GetKeepLinkedToBOMName
                Set swView = swView.GetNextView
            Wend
            'Link all views with current BOM
            Set swView = swDraw.GetFirstView
            Set swView = swView.GetNextView
            While Not swView Is Nothing
                swView.SetKeepLinkedToBOM True, strGetKeepLinkedtoBOMName
                Set swView = swView.GetNextView
            Wend
            'Edit table
            Set swView = swDraw.GetFirstView
            Set swTable = swView.GetFirstTableAnnotation
            While Not swTable Is Nothing
                If swTable.Type = swTableAnnotation_BillOfMaterials Then
                    'Delete first row of table
                    If chbAutoRefreshPartlist = True Then swTable.DeleteRow swTable.RowCount - 3
                    'Activate new first row of table
                    swApp.TaskPaneIsPinned = False
                    swModel.Extension.HideFeatureManager True
                    swModel.ViewZoomtofit2
                    Set swModelView = swModel.GetFirstModelView
                    Set TheMouse = swModelView.GetMouse
                    Set MousePosition = New Class_ClickMouse
                    MousePosition.GetCursorPosition Xposition
                    MousePosition.InfoDisplayMonitor DisplayWidth, DisplayHeight
                    If Xposition <= DisplayWidth Then
                        Call ClickMouseBOMCutList(CInt(DisplayWidth * 0.63))
                    Else
                        Call ClickMouseBOMCutList(DisplayWidth + CInt(DisplayWidth * 0.63))
                    End If
                    'Edit table
                    On Error Resume Next
                    Set swTextFormat = swTable.GetCellTextFormat(0, swTable.RowCount - 2)
                    Set swAnn = swTable.GetAnnotation
                    For i = 0 To swTable.RowCount - 3
                        Dim intLoopClickMouse As Integer
                        intLoopClickMouse = 0
                        'Info of second column
LoopClickMouseBOMCutList:
                        EndProgram.EndByESCKey
                        If InStr(1, swTable.Text(swTable.RowCount - 3 - i, 4), "AL", vbTextCompare) Then
                            swTable.Text(swTable.RowCount - 3 - i, 1) = "THANH NH" & ChrW(212) & "M " & (i + 1)
                        Else
                            swTable.Text(swTable.RowCount - 3 - i, 1) = "THANH " & (i + 1)
                        End If
                        swTable.SetCellTextFormat swTable.RowCount - 3 - i, 1, True, swTextFormat
                        swTable.CellTextHorizontalJustification(swTable.RowCount - 3 - i, 1) = swTextJustification_e.swTextJustificationCenter
                        If Len(swTable.Text(swTable.RowCount - 3, 1)) < 5 And intLoopClickMouse <= 2 Then
                            If Xposition <= DisplayWidth Then
                                Call ClickMouseBOMCutList(DisplayWidth + CInt(DisplayWidth * 0.63))
                                intLoopClickMouse = intLoopClickMouse + 1
                                GoTo LoopClickMouseBOMCutList
                            Else
                                Call ClickMouseBOMCutList(CInt(DisplayWidth * 0.63))
                                intLoopClickMouse = intLoopClickMouse + 1
                                GoTo LoopClickMouseBOMCutList
                            End If
                        End If
                        swModel.Extension.HideFeatureManager False
                        swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationHidden
                        'Info of third column
                        swTable.Text(swTable.RowCount - 3 - i, 2) = RoundString(swTable.Text(swTable.RowCount - 3 - i, 2))
                        swTable.SetCellTextFormat swTable.RowCount - 3 - i, 2, False, swTextFormat
                    Next i
                    swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationVisible
                End If
                Set swTable = swTable.GetNext
            Wend
            Set swView = swView.GetNextView
            Set swOpenModel = swApp.ActivateDoc3(swView.GetReferencedModelName, False, 2, 0)
            swOpenModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swWeldmentEnableAutomaticCutList, True
            swOpenModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swWeldmentEnableAutomaticUpdate, True
            swOpenModel.ShowConfiguration2 swView.ReferencedConfiguration
            Call SetConfigModel
            swConfig.BOMPartNoSource = swBOMPartNumberSource_e.swBOMPartNumber_DocumentName
            swOpenModel.Save3 swSaveAsOptions_Silent, Empty, Empty
            swApp.CloseDoc swView.GetReferencedModelName
        End If
        Set swTable = swTable.GetNext
    Wend
End Sub

Private Sub ClickMouseBOMCutList(longWidthScreen As Long)
    If vGetSheetProperties(0) = 8 Then 'Size A3
        TheMouse.Move longWidthScreen - 10, CInt(DisplayHeight * 0.78) + 10, swMouse_MouseMove & swMouse_Click
    ElseIf vGetSheetProperties(0) = 9 Then 'Size A2
        TheMouse.Move longWidthScreen + 100, CInt(DisplayHeight * 0.78) + 45, swMouse_MouseMove & swMouse_Click
    ElseIf vGetSheetProperties(0) = 10 Then 'Size A1
        TheMouse.Move longWidthScreen + 180, CInt(DisplayHeight * 0.78) + 85, swMouse_MouseMove & swMouse_Click
    ElseIf vGetSheetProperties(0) = 11 Then 'Size A0
        TheMouse.Move longWidthScreen + 230, CInt(DisplayHeight * 0.78) + 110, swMouse_MouseMove & swMouse_Click
    End If
End Sub

Private Function RoundString(strRound As String) As String
    If InStr(1, strRound, ".", vbTextCompare) <> 0 Then
        Dim intPositionDot As Integer
        intPositionDot = InStr(1, strRound, ".", vbTextCompare)
        If Mid(strRound, intPositionDot + 1, 1) >= 5 Then
            strRound = Left(strRound, intPositionDot - 1) + 1
        Else
            strRound = Left(strRound, intPositionDot - 1)
        End If
    End If
    RoundString = strRound
End Function

Private Sub OpenModel(strPathModel As String)
    strExtension = Right(UCase(strPathModel), 6)
    Call OpenModelSilent(strPathModel)
    blnCheckOpenModel3D = False
    If Not swOpenModel Is Nothing Then
        If swOpenModel.GetConfigurationCount > 1 Then
            Call TurnOffAutoSaved
            intCountConfigNameDefault = 0
            vconfigNames = swOpenModel.GetConfigurationNames()
            For Each vconfigName In vconfigNames
                strConfigName = vconfigName
                If InStr(1, strConfigName, "Default", vbTextCompare) <> 0 Then intCountConfigNameDefault = intCountConfigNameDefault + 1
            Next
            If intCountConfigNameDefault = UBound(vconfigNames) + 1 Then
                Call OpenModelSilent(strPathModel)
            Else
                Set swOpenModel = swApp.ActivateDoc3(strPathModel, True, 2, 0)
                swOpenModel.ShowConfiguration2 swView.ReferencedConfiguration
                Set swModelView = swOpenModel.GetFirstModelView
                swModelView.FrameState = swWindowMaximized
                Call SetClickMouse
                blnCheckOpenModel3D = True
            End If
        End If
        Call SetConfigModel
    End If
End Sub

Private Sub OpenModelSilent(strPathModel As String)
    If strExtension = "SLDPRT" Then
       Set swOpenModel = swApp.OpenDoc6(strPathModel, swDocPART, swOpenDocOptions_Silent, Empty, Empty, Empty)
    ElseIf strExtension = "SLDASM" Then
       Set swOpenModel = swApp.OpenDoc6(strPathModel, swDocASSEMBLY, swOpenDocOptions_Silent, Empty, Empty, Empty)
    End If
End Sub

Private Sub InfoPartname(swOpenModel As ModelDoc2)
    If swOpenModel.GetConfigurationCount > 1 Then
        If InStr(1, swView.ReferencedConfiguration, "Default", vbTextCompare) <> 0 Then
            Call FilterPartNameFromPathName(strPathName)
            swModelConfigMgr.Set "PART NAME", UCase(strPartName)
        Else
            If InStr(1, swView.ReferencedConfiguration, "<As Welded>", vbTextCompare) <> 0 Then
                Dim intPositionWelded As Integer
                intPositionWelded = InStr(1, swView.ReferencedConfiguration, "<As Welded>", vbTextCompare)
                strPartName = Left(swView.ReferencedConfiguration, intPositionWelded - 1)
                swModelConfigMgr.Set "PART NAME", UCase(strPartName)
            ElseIf InStr(1, swView.ReferencedConfiguration, "<As Machined>", vbTextCompare) <> 0 Then
                Dim intPositionMachined As Integer
                intPositionMachined = InStr(1, swView.ReferencedConfiguration, "<As Machined>", vbTextCompare)
                strPartName = Left(swView.ReferencedConfiguration, intPositionMachined - 1)
                swModelConfigMgr.Set "PART NAME", UCase(strPartName)
            Else
                strPartName = swView.ReferencedConfiguration
                swModelConfigMgr.Set "PART NAME", UCase(strPartName)
            End If
        End If
    Else
        Call FilterPartNameFromPathName(strPathName)
        swModelConfigMgr.Set "PART NAME", UCase(strPartName)
    End If
End Sub

Private Sub SetNameFirstSheet()
    blnStatusNameFirstSheet = False
    If IsNumeric(txtStartRenameSheet) = False Or Trim(txtStartRenameSheet) <= 0 Then
        swApp.SendMsgToUser2 "Please rename first sheet by a positive integer number!", swMbWarning, swMbOk
        blnStatusNameFirstSheet = True
    End If
End Sub

Private Sub ErrorHandling()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "Please open a drawing!", vbCritical
        End
    ElseIf swModel.GetType <> swDocDRAWING Then
        MsgBox "Please open a drawing!", vbCritical
        End
    End If
End Sub

Private Sub GetPoCodeAndSoMaySoCum()
    Set swDrawCustPropMgr = swModel.Extension.CustomPropertyManager(Empty)
    ReDim strValue(swDrawCustPropMgr.Count) As String
    vDrawCustPropNames = swDrawCustPropMgr.GetNames()
    If IsEmpty(vDrawCustPropNames) = False Then
        For i = 0 To UBound(vDrawCustPropNames)
            swDrawCustPropMgr.Get3 vDrawCustPropNames(i), False, Empty, strValue(i)
            If UCase(vDrawCustPropNames(i)) = "PO CODE" Then
                strPoCode = strValue(i)
            ElseIf UCase(vDrawCustPropNames(i)) = "DRAWING NO" Then
                strSoMaySoCum = strValue(i)
            End If
        Next i
    End If
    If strPoCode = Empty Or strPoCode = "-" Then
        MsgBox "Please input PO CODE information in Drawing Custom Properties!", vbCritical
        End
    ElseIf strSoMaySoCum = Empty Or strSoMaySoCum = "-" Then
        MsgBox "Please input So may-So cum information in Drawing Custom Properties!", vbCritical
        End
    End If
    Set swDraw = swModel
    vSheetNames = swDraw.GetSheetNames()
End Sub

Private Function StripVietnameseAccents(ByVal sInput As String) As String
    Dim i As Long
    Dim code As Long
    Dim ch As String
    Dim res As String
    res = ""
    For i = 1 To Len(sInput)
        ch = Mid(sInput, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case code
            Case 192, 193, 194, 195, 258, 7840, 7842, 7844, 7846, 7848, 7850, 7852, 7854, 7856, 7858, 7860, 7862, _
                 224, 225, 226, 227, 259, 7841, 7843, 7845, 7847, 7849, 7851, 7853, 7855, 7857, 7859, 7861, 7863
                res = res & "A"
            Case 200, 201, 202, 7864, 7866, 7868, 7870, 7872, 7874, 7876, 7878, _
                 232, 233, 234, 7865, 7867, 7869, 7871, 7873, 7875, 7877, 7879
                res = res & "E"
            Case 204, 205, 296, 7880, 7882, _
                 236, 237, 297, 7881, 7883
                res = res & "I"
            Case 210, 211, 212, 213, 416, 7884, 7886, 7888, 7890, 7892, 7894, 7896, 7898, 7900, 7902, 7904, 7906, _
                 242, 243, 244, 245, 417, 7885, 7887, 7889, 7891, 7893, 7895, 7897, 7899, 7901, 7903, 7905, 7907
                res = res & "O"
            Case 217, 218, 360, 431, 7908, 7910, 7912, 7914, 7916, 7918, 7920, _
                 249, 250, 361, 432, 7909, 7911, 7913, 7915, 7917, 7919, 7921
                res = res & "U"
            Case 221, 7922, 7924, 7926, 7928, _
                 253, 7923, 7925, 7927, 7929
                res = res & "Y"
            Case 272, 273
                res = res & "D"
            Case Else
                res = res & UCase(ch)
        End Select
    Next i
    StripVietnameseAccents = res
End Function

Private Function IsNonMetallicMaterial(ByVal sMaterial As String) As Boolean
    Dim sUpper As String
    sUpper = StripVietnameseAccents(Trim(sMaterial))
    IsNonMetallicMaterial = False
    If sUpper = "" Then Exit Function
    If InStr(sUpper, "NYLON") > 0 Or InStr(sUpper, "MC ") > 0 Or Left(sUpper, 3) = "MC9" Or _
       InStr(sUpper, "POM") > 0 Or InStr(sUpper, "ACRYLIC") > 0 Or _
       InStr(sUpper, "PET") > 0 Or InStr(sUpper, "MICA") > 0 Or _
       InStr(sUpper, "MIKA") > 0 Or InStr(sUpper, "RUBBER") > 0 Or _
       InStr(sUpper, "CAO SU") > 0 Or InStr(sUpper, "URETHANE") > 0 Or _
       InStr(sUpper, "POLYURETHANE") > 0 Or InStr(sUpper, "NHUA") > 0 Or _
       InStr(sUpper, "PLASTIC") > 0 Or InStr(sUpper, "TEFLON") > 0 Or _
       InStr(sUpper, "PTFE") > 0 Or InStr(sUpper, "BAKELITE") > 0 Or _
       InStr(sUpper, "EPDM") > 0 Or InStr(sUpper, "PVC") > 0 Or _
       InStr(sUpper, "PEEK") > 0 Or InStr(sUpper, "SILICON") > 0 Or _
       sUpper = "PU" Or Left(sUpper, 3) = "PU " Or InStr(sUpper, "PU-") > 0 Or InStr(sUpper, " PU") > 0 Then
        IsNonMetallicMaterial = True
    End If
End Function

Private Sub GetSheetBounds(swSht As Sheet, ByRef dW As Double, ByRef dH As Double)
    Dim vProps As Variant
    dW = 0.42
    dH = 0.297
    On Error Resume Next
    If Not swSht Is Nothing Then
        vProps = swSht.GetProperties2()
        If Not IsEmpty(vProps) Then
            If UBound(vProps) >= 6 Then
                If CDbl(vProps(5)) > 0.05 And CDbl(vProps(6)) > 0.05 Then
                    dW = CDbl(vProps(5))
                    dH = CDbl(vProps(6))
                End If
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Function NormalizeNoteText(ByVal sRaw As String) As String
    Dim s As String
    s = Replace(sRaw, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeNoteText = StripVietnameseAccents(Trim(s))
End Function

Private Function IsNoteInsideSheetAndManual(swN As Note, ByVal dSheetW As Double, ByVal dSheetH As Double) As Boolean
    Dim vPos As Variant
    Dim sLinkText As String
    Dim bIsBalloon As Boolean

    IsNoteInsideSheetAndManual = False
    If swN Is Nothing Then Exit Function

    ' 1. Bo qua BOM Balloon
    bIsBalloon = False
    On Error Resume Next
    bIsBalloon = swN.IsBomBalloon()
    On Error GoTo 0
    If bIsBalloon Then Exit Function

    ' 2. Bo qua moi Note lien ket thuoc tinh ($PRP / $PRPSHEET / $PRPVIEW) -> Chan doc nguoc Khung ten (REMARK, PART NAME...)
    sLinkText = ""
    On Error Resume Next
    sLinkText = swN.PropertyLinkedText
    On Error GoTo 0
    If InStr(1, sLinkText, "$PRP", vbTextCompare) > 0 Then Exit Function

    ' 3. Bat buoc nam BEN TRONG kho giay (Chan doc cac Note mau de ngoai ria khung giay A3/A2/A1/A0)
    On Error Resume Next
    vPos = swN.GetTextPoint2
    On Error GoTo 0
    If IsEmpty(vPos) Then Exit Function
    If vPos(0) > 0 And vPos(0) < dSheetW And vPos(1) > 0 And vPos(1) < dSheetH Then
        IsNoteInsideSheetAndManual = True
    End If
End Function

Private Function IsAnnotationInsideSheet(swA As Annotation, ByVal dSheetW As Double, ByVal dSheetH As Double) As Boolean
    Dim vPos As Variant
    IsAnnotationInsideSheet = False
    If swA Is Nothing Then Exit Function
    On Error Resume Next
    vPos = swA.GetPosition
    On Error GoTo 0
    If IsEmpty(vPos) Then
        IsAnnotationInsideSheet = True
        Exit Function
    End If
    If vPos(0) > 0 And vPos(0) < dSheetW And vPos(1) > 0 And vPos(1) < dSheetH Then
        IsAnnotationInsideSheet = True
    End If
End Function

Private Function IsViewWelded(swV As View, Optional ByVal dSheetW As Double = 0.42, Optional ByVal dSheetH As Double = 0.297) As Boolean
    Dim swN As Note
    Dim swA As Annotation
    Dim swT As TableAnnotation
    Dim sNoteText As String

    IsViewWelded = False
    If swV Is Nothing Then Exit Function

    ' 1. Kiem tra bang Weldment Cut-List
    Set swT = swV.GetFirstTableAnnotation
    While Not swT Is Nothing
        If swT.Type = swTableAnnotationType_e.swTableAnnotation_WeldmentCutList Then
            IsViewWelded = True
            Exit Function
        End If
        Set swT = swT.GetNext
    Wend

    ' 2. Kiem tra Note chu thich tinh nam TRONG kho giay (da chuan hoa khong dau -> bat duoc ca "Ban ve han" nua co dau)
    Set swN = swV.GetFirstNote
    While Not swN Is Nothing
        If IsNoteInsideSheetAndManual(swN, dSheetW, dSheetH) Then
            sNoteText = NormalizeNoteText(swN.GetText)
            If InStr(1, sNoteText, "BAN VE HAN", vbTextCompare) <> 0 Or _
               InStr(1, sNoteText, "BAN TONG HAN", vbTextCompare) <> 0 Or _
               InStr(1, sNoteText, "BV HAN TONG", vbTextCompare) <> 0 Or _
               InStr(1, sNoteText, "MOI HAN", vbTextCompare) <> 0 Then
                IsViewWelded = True
                Exit Function
            End If
            If InStr(1, sNoteText, "BV HAN", vbTextCompare) <> 0 And _
               InStr(1, sNoteText, "CON", vbTextCompare) = 0 Then
                IsViewWelded = True
                Exit Function
            End If
        End If
        Set swN = swN.GetNext
    Wend

    ' 3. Kiem tra ky hieu moi han nam TRONG kho giay
    Set swA = swV.GetFirstAnnotation
    While Not swA Is Nothing
        If swA.GetType = swAnnotationType_e.swWeldSymbol Or swA.GetType = swAnnotationType_e.swWeldBeadSymbol Then
            If IsAnnotationInsideSheet(swA, dSheetW, dSheetH) Then
                IsViewWelded = True
                Exit Function
            End If
        End If
        Set swA = swA.GetNext
    Wend
End Function

Private Sub CheckWeldmentPhrase()
    Dim swV As View
    Dim swN As Note
    Dim swA As Annotation
    Dim swT As TableAnnotation
    Dim swCurSht As Sheet
    Dim dSheetW As Double
    Dim dSheetH As Double
    Dim sNoteText As String

    blnCheckWeldment_1 = False
    blnCheckWeldment_2 = False
    blnCheckWeldment_3 = False
    blnCheckWeldmentChild = False

    Set swCurSht = swDraw.GetCurrentSheet
    Call GetSheetBounds(swCurSht, dSheetW, dSheetH)

    Set swV = swDraw.GetFirstView
    While Not swV Is Nothing
        ' 1. Kiem tra bang Cut-List
        Set swT = swV.GetFirstTableAnnotation
        While Not swT Is Nothing
            If swT.Type = swTableAnnotationType_e.swTableAnnotation_WeldmentCutList Then
                blnCheckWeldment_1 = True
            End If
            Set swT = swT.GetNext
        Wend

        ' 2. Kiem tra cac Note chu thich tinh nam TRONG kho giay (da chuan hoa khong dau)
        Set swN = swV.GetFirstNote
        While Not swN Is Nothing
            If IsNoteInsideSheetAndManual(swN, dSheetW, dSheetH) Then
                sNoteText = NormalizeNoteText(swN.GetText)

                If InStr(1, sNoteText, "HAN CON", vbTextCompare) <> 0 Then
                    blnCheckWeldmentChild = True
                End If

                If InStr(1, sNoteText, "BAN VE HAN", vbTextCompare) <> 0 Or _
                   InStr(1, sNoteText, "BAN TONG HAN", vbTextCompare) <> 0 Or _
                   InStr(1, sNoteText, "BV HAN TONG", vbTextCompare) <> 0 Then
                    blnCheckWeldment_1 = True
                End If

                If InStr(1, sNoteText, "BV HAN", vbTextCompare) <> 0 And _
                   InStr(1, sNoteText, "CON", vbTextCompare) = 0 Then
                    blnCheckWeldment_1 = True
                End If

                If InStr(1, sNoteText, "MOI HAN", vbTextCompare) <> 0 Then
                    blnCheckWeldment_2 = True
                End If
            End If
            Set swN = swN.GetNext
        Wend

        ' 3. Kiem tra ky hieu moi han nam TRONG kho giay
        Set swA = swV.GetFirstAnnotation
        While Not swA Is Nothing
            If swA.GetType = swAnnotationType_e.swWeldSymbol Or swA.GetType = swAnnotationType_e.swWeldBeadSymbol Then
                If IsAnnotationInsideSheet(swA, dSheetW, dSheetH) Then
                    blnCheckWeldment_3 = True
                End If
            End If
            Set swA = swA.GetNext
        Wend

        Set swV = swV.GetNextView
    Wend
End Sub

Private Function StripCompInstanceName(ByVal sCompName2 As String) As String
    Dim sLeaf As String
    Dim pSlash As Long
    Dim pDash As Long
    Dim k As Long
    Dim ch As String
    Dim bDigits As Boolean

    StripCompInstanceName = ""
    sLeaf = Trim(sCompName2)
    If sLeaf = "" Then Exit Function

    pSlash = InStrRev(sLeaf, "/")
    If pSlash > 0 And pSlash < Len(sLeaf) Then
        sLeaf = Trim(Mid(sLeaf, pSlash + 1))
    End If

    pDash = InStrRev(sLeaf, "-")
    If pDash > 1 And pDash < Len(sLeaf) Then
        bDigits = True
        For k = pDash + 1 To Len(sLeaf)
            ch = Mid(sLeaf, k, 1)
            If ch < "0" Or ch > "9" Then
                bDigits = False
                Exit For
            End If
        Next k
        If bDigits Then
            sLeaf = Trim(Left(sLeaf, pDash - 1))
        End If
    End If

    StripCompInstanceName = UCase(sLeaf)
End Function

Private Sub RegisterWeldedAssemblyComponents(swAsmDoc As ModelDoc2)
    Dim swAsmLocal As AssemblyDoc
    Dim vCompsLocal As Variant
    Dim swCompLocal As Component2
    Dim compPath As String
    Dim compFileName As String
    Dim compBaseName As String
    Dim compTreeBase As String
    Dim idxComp As Long

    On Error Resume Next
    If swAsmDoc Is Nothing Then Exit Sub
    If dictWeldmentChildParts Is Nothing Then
        Set dictWeldmentChildParts = CreateObject("Scripting.Dictionary")
    End If

    Set swAsmLocal = swAsmDoc
    If swAsmLocal Is Nothing Then Exit Sub

    vCompsLocal = swAsmLocal.GetComponents(False)
    If Not IsEmpty(vCompsLocal) Then
        For idxComp = LBound(vCompsLocal) To UBound(vCompsLocal)
            Set swCompLocal = vCompsLocal(idxComp)
            If Not swCompLocal Is Nothing Then
                If Not swCompLocal.IsEnvelope Then
                    compPath = UCase(Trim(swCompLocal.GetPathName))
                    If Right(compPath, 6) = "SLDPRT" Then
                        If Not dictWeldmentChildParts.Exists(compPath) Then
                            dictWeldmentChildParts.Add compPath, True
                        End If
                        compFileName = Mid(compPath, InStrRev(compPath, "\") + 1)
                        If compFileName <> "" And Not dictWeldmentChildParts.Exists(compFileName) Then
                            dictWeldmentChildParts.Add compFileName, True
                        End If
                        If InStrRev(compFileName, ".") > 1 Then
                            compBaseName = Left(compFileName, InStrRev(compFileName, ".") - 1)
                            If compBaseName <> "" And Not dictWeldmentChildParts.Exists(compBaseName) Then
                                dictWeldmentChildParts.Add compBaseName, True
                            End If
                        End If
                    End If

                    ' Bo sung lay ten tu cay lap rap Name2 (luon co ngay ca khi Lightweight chua resolve duong dan)
                    If Right(compPath, 6) <> "SLDASM" Then
                        compTreeBase = StripCompInstanceName(swCompLocal.Name2)
                        If compTreeBase <> "" Then
                            If Not dictWeldmentChildParts.Exists(compTreeBase) Then
                                dictWeldmentChildParts.Add compTreeBase, True
                            End If
                            If Not dictWeldmentChildParts.Exists(compTreeBase & ".SLDPRT") Then
                                dictWeldmentChildParts.Add compTreeBase & ".SLDPRT", True
                            End If
                        End If
                    End If
                End If
            End If
        Next idxComp
    End If
    On Error GoTo 0
End Sub

Private Sub ScanWeldedAssemblies()
    Dim vSheetsViews As Variant
    Dim vSheetArray As Variant
    Dim vShtNamesLocal As Variant
    Dim swShtLocal As Sheet
    Dim swV As View
    Dim swModelRefView As View
    Dim strModelPath As String
    Dim strExt As String
    Dim swAsmDoc As ModelDoc2
    Dim idxSheet As Long, idxView As Long
    Dim blnIsWeldAsm As Boolean
    Dim dSheetW As Double, dSheetH As Double
    Dim lErr As Long, lWarn As Long
    Dim sOrigActiveSheet As String

    On Error Resume Next
    Set dictWeldmentChildParts = CreateObject("Scripting.Dictionary")

    If swDraw Is Nothing Then Set swDraw = swApp.ActiveDoc
    If swDraw Is Nothing Then Exit Sub

    sOrigActiveSheet = ""
    If Not swDraw.GetCurrentSheet Is Nothing Then
        sOrigActiveSheet = swDraw.GetCurrentSheet.GetName
    End If

    vSheetsViews = swDraw.GetViews
    If IsEmpty(vSheetsViews) Then Exit Sub
    vShtNamesLocal = swDraw.GetSheetNames

    For idxSheet = LBound(vSheetsViews) To UBound(vSheetsViews)
        vSheetArray = vSheetsViews(idxSheet)
        If Not IsEmpty(vSheetArray) Then
            dSheetW = 0.42
            dSheetH = 0.297
            If Not IsEmpty(vShtNamesLocal) Then
                If idxSheet <= UBound(vShtNamesLocal) Then
                    Set swShtLocal = swDraw.Sheet(vShtNamesLocal(idxSheet))
                    Call GetSheetBounds(swShtLocal, dSheetW, dSheetH)
                End If
            End If

            strModelPath = ""
            Set swModelRefView = Nothing
            For idxView = 1 To UBound(vSheetArray)
                Set swV = vSheetArray(idxView)
                If Not swV Is Nothing Then
                    strModelPath = swV.GetReferencedModelName
                    If strModelPath <> "" Then
                        Set swModelRefView = swV
                        Exit For
                    End If
                End If
            Next idxView

            strExt = Right(UCase(Trim(strModelPath)), 6)
            If strExt = "SLDASM" Then
                blnIsWeldAsm = False
                For idxView = 0 To UBound(vSheetArray)
                    Set swV = vSheetArray(idxView)
                    If Not swV Is Nothing Then
                        If IsViewWelded(swV, dSheetW, dSheetH) Then
                            blnIsWeldAsm = True
                            Exit For
                        End If
                    End If
                Next idxView

                ' Neu sheet SLDASM dang inactive khien SolidWorks chua nap Note/Table, kich hoat sheet SLDASM de kiem tra chinh xac
                If Not blnIsWeldAsm And Not IsEmpty(vShtNamesLocal) Then
                    If idxSheet <= UBound(vShtNamesLocal) Then
                        swDraw.ActivateSheet vShtNamesLocal(idxSheet)
                        Set swV = swDraw.GetFirstView
                        While Not swV Is Nothing
                            If IsViewWelded(swV, dSheetW, dSheetH) Then
                                blnIsWeldAsm = True
                            End If
                            Set swV = swV.GetNextView
                        Wend
                    End If
                End If

                If blnIsWeldAsm Then
                    Set swAsmDoc = swApp.GetOpenDocumentByName(strModelPath)
                    If swAsmDoc Is Nothing And Not swModelRefView Is Nothing Then
                        Set swAsmDoc = swModelRefView.ReferencedDocument
                    End If
                    If swAsmDoc Is Nothing Then
                        lErr = 0
                        lWarn = 0
                        Set swAsmDoc = swApp.OpenDoc6(strModelPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", lErr, lWarn)
                    End If
                    If Not swAsmDoc Is Nothing Then
                        Call RegisterWeldedAssemblyComponents(swAsmDoc)
                    End If
                End If
            End If
        End If
    Next idxSheet

    If sOrigActiveSheet <> "" Then
        swDraw.ActivateSheet sOrigActiveSheet
    End If
    On Error GoTo 0
End Sub

Private Function IsChildOfWeldedAssembly(ByVal strPartPath As String) As Boolean
    Dim sFull As String
    Dim sFile As String
    Dim sBase As String

    IsChildOfWeldedAssembly = False
    If dictWeldmentChildParts Is Nothing Then Exit Function
    If dictWeldmentChildParts.Count = 0 Then Exit Function

    sFull = UCase(Trim(strPartPath))
    If sFull <> "" Then
        If dictWeldmentChildParts.Exists(sFull) Then
            IsChildOfWeldedAssembly = True
            Exit Function
        End If

        sFile = Mid(sFull, InStrRev(sFull, "\") + 1)
        If sFile <> "" And dictWeldmentChildParts.Exists(sFile) Then
            IsChildOfWeldedAssembly = True
            Exit Function
        End If

        If InStrRev(sFile, ".") > 1 Then
            sBase = Left(sFile, InStrRev(sFile, ".") - 1)
            If sBase <> "" And dictWeldmentChildParts.Exists(sBase) Then
                IsChildOfWeldedAssembly = True
                Exit Function
            End If
        End If
    End If

    If strPartName <> "" Then
        If dictWeldmentChildParts.Exists(UCase(Trim(strPartName))) Then
            IsChildOfWeldedAssembly = True
            Exit Function
        End If
    End If
End Function

Private Function GetWeldingRemark(ByVal strModelPath As String, ByVal strExt As String) As String
    Dim strRemark As String
    Dim sExt As String
    Dim swRevTblLocal As SldWorks.RevisionTableAnnotation
    Dim sRevLocal As String
    Dim sCheckMat As String
    strRemark = ""

    If dictWeldmentChildParts Is Nothing Then
        Call ScanWeldedAssemblies
    End If

    sExt = UCase(Trim(strExt))

    ' 1. Kiem tra ghi chu / ky hieu han thuc su nam TRONG kho giay 2D
    Call CheckWeldmentPhrase

    If sExt = "SLDASM" Then
        ' Ban ve Lap rap (Assembly):
        If blnCheckWeldmentChild = True Then
            strRemark = "BV H" & ChrW(192) & "N CON"
        ElseIf blnCheckWeldment_1 = True Or blnCheckWeldment_2 = True Or blnCheckWeldment_3 = True Then
            strRemark = "BV H" & ChrW(192) & "N T" & ChrW(7892) & "NG"
            If Not swOpenModel Is Nothing Then
                Call RegisterWeldedAssemblyComponents(swOpenModel)
            End If
        Else
            strRemark = "BV L" & ChrW(7854) & "P"
        End If
    ElseIf sExt = "SLDPRT" Then
        ' Ban ve Part:
        sCheckMat = ""
        On Error Resume Next
        If Not swModelConfigMgr Is Nothing Then
            sCheckMat = swModelConfigMgr.Get("MATERIAL")
        End If
        If sCheckMat = "" And Not swOpenModel Is Nothing Then
            sCheckMat = swOpenModel.CustomInfo2("", "MATERIAL")
        End If
        On Error GoTo 0

        If IsNonMetallicMaterial(sCheckMat) Then
            strRemark = ""
        Else
            If blnCheckWeldmentChild = True Then
                strRemark = "BV H" & ChrW(192) & "N CON"
            ElseIf IsChildOfWeldedAssembly(strModelPath) Then
                strRemark = "BV H" & ChrW(192) & "N CON"
            ElseIf blnCheckWeldment_1 = True Or blnCheckWeldment_2 = True Or blnCheckWeldment_3 = True Then
                strRemark = "BV H" & ChrW(192) & "N T" & ChrW(7892) & "NG"
            Else
                strRemark = ""
            End If
        End If
    End If

    ' Remark semi-standard (Gia cong tu TBTC)
    Call CheckTBTC
    If blnCheckTBTC = True Then
        strRemark = "GC T" & ChrW(7914) & " TBTC"
    End If

    ' Remark Editing Drawing (Ban ve sua)
    sRevLocal = strRevision
    On Error Resume Next
    Set swRevTblLocal = swDraw.GetCurrentSheet.RevisionTable
    If Not swRevTblLocal Is Nothing Then
        If swRevTblLocal.CurrentRevision <> "" Then sRevLocal = swRevTblLocal.CurrentRevision
    End If
    On Error GoTo 0
    If sRevLocal <> "" Then
        strRemark = "BV S" & ChrW(7916) & "A"
    End If

    GetWeldingRemark = strRemark
End Function

Private Sub CheckEditingPhrase()
    Dim swV As View
    Dim swN As Note
    Dim vPos As Variant
    Dim bInSheet As Boolean

    blnCheckEditingDrawing = False
    Set swV = swDraw.GetFirstView
    While Not swV Is Nothing
        Set swN = swV.GetFirstNote
        While Not swN Is Nothing
            vPos = swN.GetTextPoint2
            bInSheet = False
            If Not IsEmpty(vPos) Then
                If vPos(0) > 0 And vPos(0) < 0.42 And vPos(1) > 0 And vPos(1) < 0.297 Then bInSheet = True
            End If
            If (InStr(1, UCase(swN.GetText), "B" & ChrW(7842) & "N V" & ChrW(7868) & " S" & ChrW(7916) & "A", vbTextCompare) <> 0 Or InStr(1, UCase(swN.GetText), "BAN VE SUA", vbTextCompare) <> 0) And bInSheet = True Then
                blnCheckEditingDrawing = True
            End If
            Set swN = swN.GetNext
        Wend
        Set swV = swV.GetNextView
    Wend
End Sub

Private Sub SetConfigModel()
    Set swConfigMgr = swOpenModel.ConfigurationManager
    Set swConfig = swConfigMgr.ActiveConfiguration
    Set swModelConfigMgr = swConfig.CustomPropertyManager
    vModelConfigMgr = swModelConfigMgr.GetNames()
End Sub

Private Sub FilterPartNameFromPathName(strCallName As String)
    intLastSlashesPosition = InStrRev(strCallName, "\", -1, vbTextCompare)
    strPartName = Mid(strCallName, intLastSlashesPosition + 1, Len(strCallName) - intLastSlashesPosition - 7)
End Sub

Private Sub SaveAndCloseModel3D(strPath As String)
    If blnCheckOpenModel3D = True Then
       swOpenModel.Save3 swSaveAsOptions_Silent, Empty, Empty
       swApp.CloseDoc strPath
       If opbtQRAdd = True Then 'Neu tick tao QR thì
                        If opbtCheckInternet = True Then
                            Call DelQR_offline
                            Call AddQR
                        ElseIf opbtCheckInternet = False Then
                            Call Del_QR_RS
                            Call QR_offline_Click 'Call AddQR
                            End If
                    End If
    End If
End Sub

Private Sub cmdEXPORTSTP_Click()
End Sub

Private Sub cmdQRADD_Click()
End Sub

Private Sub MultiPage1_Change()
End Sub

Private Sub Partlist_current_Click()
    On Error Resume Next
    opbtAllSheetsAutoPartlist.Value = False
    opbtRangeAutoPartlist.Value = False
    opbtCurrentSheetAutoPartlist.Value = True
    Call cmdAUTOPARTLIST2D_Click
End Sub

Private Sub Del_QR_Click()
Call Del_QR_RS
Call DelQR_offline
End
End Sub

Private Sub Del_QR_RS()
Dim swSketch As SldWorks.Sketch
Dim vPics As Variant
Dim swSkPic As SldWorks.SketchPicture
Dim swFeat As SldWorks.Feature
Dim i As Integer
 Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    Set swSketch = swModel.SketchManager.ActiveSketch
    If Not swSketch Is Nothing Then
        vPics = swSketch.GetSketchPictures
        If Not IsEmpty(vPics) Then
            For i = 0 To UBound(vPics)
                Set swSkPic = vPics(i)
                Set swFeat = swSkPic.GetFeature
                swModel.ClearSelection2 True
                swModel.Extension.SelectByID2 swFeat.name, "SKETCHBITMAP", 0, 0, 0, False, 0, Nothing, 0
                swModel.Extension.DeleteSelection2 swDeleteSelectionOptions_e.swDelete_Absorbed
            Next
        End If
    End If
End Sub

'þ========================================================

Private Sub EXPGC_Click()
Call cmdEXPORTPDF_Click
Call cmdEXPORTDXF_DWG_Click
    EndProgram.EndByESCKey
    swDraw.ActivateSheet (vSheetNames(intSheetNumber))
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    If Not swView Is Nothing Then
        strConfigurationName = swView.ReferencedConfiguration
        strPathName = swView.GetReferencedModelName
Call OpenModel(strPathName)
End If
Call STEP_exp_Click
End Sub

Private Sub opbtCheckInternet_Click()
End Sub

Private Sub opbtCurrentSheetAutoPartlist_Click()
End Sub

Private Sub opbtDWG_Click()
    If opbtDWG = True Then
        opbtDXF = False
    End If
End Sub

Private Sub opbtDXF_Click()
    If opbtDXF = True Then
        opbtDWG = False
    End If
End Sub

Private Sub opbtQRAdd_Click()
End Sub

Private Sub strRangeQR_Change()
    opbtRangeQR.Value = True
End Sub

Private Sub opbtRangeAutoPartlist_Click()
End Sub

Private Sub opbtRangeCAD2D_Click()
End Sub

Private Sub strRangeAutoPartlist_Change()
    opbtRangeAutoPartlist.Value = True
End Sub

Private Sub FilterPartNameTopLevelTemp(strPartNameTopLevelTemp As String)
    blnCheckFirstSlashes = False
    If InStr(1, strPartNameTopLevelTemp, "/", vbTextCompare) <> 0 Then
        intFirstSlashesPosition = InStr(1, strPartNameTopLevelTemp, "/", vbTextCompare)
        strPartNameTopLevel_Temp = Left(strPartNameTopLevelTemp, intFirstSlashesPosition - 1)
        blnCheckFirstSlashes = True
    Else
        strPartNameTopLevel_Temp = strPartNameTopLevelTemp
    End If
End Sub

Private Sub FilterPartNameTopLevelLast(strPartNameTopLevelLast As String)
    blnCheckLastHorizontalMark = False
    If InStrRev(strPartNameTopLevelLast, "-", -1, vbTextCompare) <> 0 Then
        intHorizontalMarkPosition = InStrRev(strPartNameTopLevelLast, "-", -1, vbTextCompare)
        strPartNameTopLevel_Last = Left(strPartNameTopLevelLast, intHorizontalMarkPosition - 1)
        blnCheckLastHorizontalMark = True
    Else
        strPartNameTopLevel_Last = strPartNameTopLevelLast
    End If
End Sub

Sub SetClickMouse()
    If blnStatusExportExcel = False And blnStatusBackup = False And blnStatusExportPDF = False And statusExportCAD2D = False And blnAutoRemoveExistDrawingNo = False And blnStatusSetDrawingNo = False And blnSetAutoQty = False Then
        Set swModelView = swOpenModel.GetFirstModelView
        Set TheMouse = swModelView.GetMouse
        Set MousePosition = New Class_ClickMouse
        MousePosition.GetCursorPosition Xposition
        MousePosition.InfoDisplayMonitor DisplayWidth, DisplayHeight
        swApp.ActivateTaskPane (swCustomProps)
        If Xposition <= DisplayWidth Then
            TheMouse.Move DisplayWidth - 260, 160, swMouse_MouseMove & swMouse_Click
        Else
            TheMouse.Move DisplayWidth * 2 - 260, 160, swMouse_MouseMove & swMouse_Click
        End If
    End If
End Sub

Private Sub TurnOffAutoSaved()
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swAutoSaveEnable, False
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swAutoSaveInterval, 0
End Sub

Private Sub TurnOnAutoSaved()
    'swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swAutoSaveEnable, True
End Sub

Private Sub GetRange(strGetStringInput As String, iThreshold As Integer)
    If blnStatusMainLoop = True Or blnStatusExportExcel = True Or blnStatusExportPDF = True Or statusExportCAD2D = True Or statusExportDWG = True Then
        flag = False
        strGetStringInput = Trim(strGetStringInput)
        If InStr(1, strGetStringInput, ".", vbTextCompare) <> 0 Then
            swApp.SendMsgToUser2 "Please input symbol , not symbol . !", swMbWarning, swMbOk
            flag = True
            Exit Sub
        ElseIf (opbtRangePDF = True And strGetStringInput = "") Or (opbtRangeCAD2D = True And strGetStringInput = "") Or (opbtRangeAutoPartlist = True And strGetStringInput = "") Or (opbtRangeExportExcel = True And strGetStringInput = "") Or (opbtRangeSTP = True And strGetStringInput = "") Then
            swApp.SendMsgToUser2 "Please input number in textbox!", swMbWarning, swMbOk
            flag = True
            Exit Sub
        End If
        Dim strSplitSpring() As String
        ReDim strSplitSpring(Len(strGetStringInput)) As String
        ReDim intGetRange(1000) As Integer
        strSplitSpring = Split(strGetStringInput, ",", -1, vbTextCompare)
        For a = 0 To UBound(strSplitSpring)
            If strSplitSpring(a) = "" Then
                swApp.SendMsgToUser2 "Please input number between symbol , !", swMbWarning, swMbOk
                flag = True
                Exit Sub
            ElseIf Asc(strSplitSpring(a)) = 48 Then
                swApp.SendMsgToUser2 "Please input positive number!", swMbWarning, swMbOk
                flag = True
                Exit Sub
            End If
        Next a
        b = 0
        Dim intFirstStringValue As Integer
        Dim intLastStringValue As Integer
        For a = 0 To UBound(strSplitSpring)
            intHorizontalMarkPosition = InStr(1, strSplitSpring(a), "-", vbTextCompare)
            If intHorizontalMarkPosition <> 0 Then
                If IsNumeric(Left(strSplitSpring(a), intHorizontalMarkPosition - 1)) = True Then
                    intFirstStringValue = CInt(Left(strSplitSpring(a), intHorizontalMarkPosition - 1))
                    If intFirstStringValue > iThreshold + 1 Or intFirstStringValue < CInt(Trim(txtStartRenameSheet)) Then
                        swApp.SendMsgToUser2 "Please input data in allowed threshold!", swMbWarning, swMbOk
                        flag = True
                        Exit Sub
                    End If
                Else
                    swApp.SendMsgToUser2 "Please input data again!", swMbWarning, swMbOk
                    flag = True
                    Exit Sub
                End If
                If IsNumeric(Right(strSplitSpring(a), Len(strSplitSpring(a)) - intHorizontalMarkPosition)) = True Then
                    intLastStringValue = CInt(Right(strSplitSpring(a), Len(strSplitSpring(a)) - intHorizontalMarkPosition))
                    If intLastStringValue > iThreshold + 1 Or intLastStringValue < CInt(Trim(txtStartRenameSheet)) Then
                        swApp.SendMsgToUser2 "Please input data in allowed threshold!", swMbWarning, swMbOk
                        flag = True
                        Exit Sub
                    End If
                Else
                    swApp.SendMsgToUser2 "Please input data again!", swMbWarning, swMbOk
                    flag = True
                    Exit Sub
                End If
                If intLastStringValue - intFirstStringValue > 0 Then
                    For c = intFirstStringValue To intLastStringValue
                        intGetRange(b) = c
                        b = b + 1
                    Next c
                Else
                    swApp.SendMsgToUser2 "Please input range data ascending!", swMbWarning, swMbOk
                    flag = True
                    Exit Sub
                End If
            Else
                If IsNumeric(strSplitSpring(a)) = True Then
                    intGetRange(b) = CInt(strSplitSpring(a))
                    If intGetRange(b) > iThreshold + 1 Or intGetRange(b) < CInt(Trim(txtStartRenameSheet)) Then
                        swApp.SendMsgToUser2 "Please input data in allowed threshold!", swMbWarning, swMbOk
                        flag = True
                        Exit Sub
                    End If
                    b = b + 1
                Else
                    swApp.SendMsgToUser2 "Please input data again!", swMbWarning, swMbOk
                    flag = True
                    Exit Sub
                End If
            End If
        Next a
        iMemory2 = b
        For a = 0 To iMemory2 - 2
CheckDuplicateRange:
            For b = a + 1 To iMemory2 - 1
                If intGetRange(b) = intGetRange(a) Then
                    For c = b To iMemory2 - 1
                        intGetRange(c) = intGetRange(c + 1)
                    Next c
                    iMemory2 = iMemory2 - 1
                End If
                If intGetRange(b) = intGetRange(a) Then GoTo CheckDuplicateRange
            Next b
        Next a
        Dim iMin As Integer
        For a = 0 To iMemory2 - 2
            For b = a + 1 To iMemory2 - 1
                iMin = intGetRange(a)
                If intGetRange(b) < iMin Then
                    intGetRange(a) = intGetRange(b)
                    intGetRange(b) = iMin
                End If
            Next b
        Next a
    End If
    If strGetStringInput <> "" And blnStatusExportExcel = False And blnStatusExportPDF = False And statusExportCAD2D = False Then
        If intGetRange(iMemory2 - 1) > iThreshold + 1 And blnStatusMainLoop = False And intCountDuplicateSheet > 0 Then
            iMemory2 = iMemory2 - intCountDuplicateSheet
        ElseIf intGetRange(iMemory2 - 1) > iThreshold + 1 And blnStatusMainLoop = False And intCountEmptySheet > 0 Then
            iMemory2 = iMemory2 - intCountEmptySheet
        End If
    End If
End Sub

Private Sub ErrorHandling2()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "Please open an assembly!", vbCritical
        End
    ElseIf swModel.GetType <> swDocASSEMBLY Then
        MsgBox "Please open an assembly!", vbCritical
        End
    End If
End Sub

Private Sub cmdSetSheetScale_Click()
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    AUTOPARTLIST.Hide
    Setsheetscale.Show
End Sub

Private Sub cmdRemoveEnvelope_Click()
    AUTOPARTLIST.Hide
    Call ErrorHandling2
    Set swAsm = swModel
    vComps = swAsm.GetComponents(False)
    For i = LBound(vComps) To UBound(vComps)
        Set swComp = vComps(i)
        If swComp.IsEnvelope() Then
            If swComp.Solving() = 0 Or swComp.Solving() = -1 Then
                swComp.Select3 False, Nothing
                swAsm.CompConfigProperties5 swComponentFullyResolved, swComponentRigidSolving, True, True, Empty, False, False
            ElseIf swComp.Solving() = 1 Then
                swComp.Select3 False, Nothing
                swAsm.CompConfigProperties5 swComponentFullyResolved, swComponentFlexibleSolving, True, True, Empty, False, False
            End If
        End If
    Next
    swApp.SendMsgToUser2 "REMOVE ENVELOPE COMPLETED!", swMbInformation, swMbOk
    Unload AUTOPARTLIST
End Sub

Private Sub cmdEXIT_Click()
    End
End Sub

Private Sub cmdEXPORTPDF_Click()
Call CheckPDFFolderExists
End Sub

Private Sub cmdEXPORTPDF_Click_1()
Call CheckRun
    blnStatusExportPDF = True
    AUTOPARTLIST.Hide
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Call GetRange(strRangePDF, CInt(UBound(vSheetNames) + txtStartRenameSheet - 1))
    If flag = False Then
'        If swApp.SendMsgToUser2("Please choose a created folder to save file PDF!", swMbInformation, swMbYesNo) = swMbHitYes Then
            Call EXPORTPDF
            End
'        Else
'            GoTo EndExportPDF
'        End If
    End If
    blnStatusExportPDF = False
EndExportPDF:
    AUTOPARTLIST.Show
End Sub

Private Sub EXPORTPDF()
Dim Pathdwg As String
    strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    PathPDF = strGetCurrentPath + "\PDF\"
    Set EndProgram = New Class_ClickESC
    Set swExportPDFData = swApp.GetExportFileData(1)
    If blnStatusExportExcel = False Then strFilePath = PathPDF 'BrowseFolder
    If strFilePath <> "" Then
        If opbtAllToSeparateFilePDF Then
            For intSheetNumber = 0 To UBound(vSheetNames)
                EndProgram.EndByESCKey
                ProcessExportPDF intSheetNumber
            Next intSheetNumber
        ElseIf opbtAllInOneFilePDF Then
            blnExportExcel = swExportPDFData.SetSheets(swExportData_ExportAllSheet, vSheetNames)
            If opbtDrawingNoPDF = True Then
                blnExportExcel = swModel.Extension.SaveAs(strFilePath & "\" & UCase(strPoCode) & "-" & strSoMaySoCum & ".pdf", 0, 0, swExportPDFData, Empty, Empty)
            ElseIf opbtPartnamePDF = True Then
                Call FilterPartNameFromPathName(swModel.GetPathName)
                blnExportExcel = swModel.Extension.SaveAs(strFilePath & "\" & strPartName & ".pdf", 0, 0, swExportPDFData, Empty, Empty)
            End If
        ElseIf opbtRangePDF Then
            For i = 0 To iMemory2 - 1
                EndProgram.EndByESCKey
                intSheetNumber = CInt(intGetRange(i) - txtStartRenameSheet)
                ProcessExportPDF intSheetNumber
            Next i
        ElseIf opbtCurrentSheetPDF Then
            For intSheetNumber = 0 To UBound(vSheetNames)
                If vSheetNames(intSheetNumber) = swDraw.GetCurrentSheet.GetName Then Call ProcessExportPDF(intSheetNumber)
            Next intSheetNumber
        End If
        swApp.SendMsgToUser2 "EXPORT FILE PDF COMPLETED!", swMbInformation, swMbOk
    Else
        Exit Sub
    End If
End Sub

Private Sub ProcessExportPDF(intSheetNumber As Integer)
    swDraw.ActivateSheet vSheetNames(intSheetNumber)
    blnExportExcel = swExportPDFData.SetSheets(swExportData_ExportCurrentSheet, vSheetNames)
    If opbtDrawingNoPDF = True Then
        Call InfoDrawingNo
        blnExportExcel = swModel.Extension.SaveAs(strFilePath & "\" & strDrawingNo & ".pdf", 0, 0, swExportPDFData, Empty, Empty)
    ElseIf opbtPartnamePDF = True Then
        Set swView = swDraw.GetFirstView
        Set swView = swView.GetNextView
        strPathName = swView.GetReferencedModelName
        Call OpenModel(strPathName)
        Call InfoPartname(swOpenModel)
        Call SaveAndCloseModel3D(strPathName)
        blnExportExcel = swModel.Extension.SaveAs(strFilePath & "\" & strPartName & ".pdf", 0, 0, swExportPDFData, Empty, Empty)
    End If
End Sub

Private Sub strRangePDF_Change()
    opbtRangePDF.Value = True
End Sub

'----------

Private Sub cmdEXPORTDXF_DWG_Click()
Call CheckDXFFolderExists
End Sub

Private Sub cmdEXPORTDXF_DWG_1()
Call CheckRun
    statusExportCAD2D = True
    AUTOPARTLIST.Hide
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Call GetRange(strRangeCAD2D, CInt(UBound(vSheetNames) + txtStartRenameSheet - 1))
    If flag = False Then
'        If swApp.SendMsgToUser2("Please choose a created folder to save file DXF/DWG!", swMbInformation, swMbYesNo) = swMbHitYes Then
            If opbtDXF = True Then
                Call EXPORTCAD2D("dxf")
            ElseIf opbtDWG = True Then
                Call EXPORTCAD2D("dwg")
            End If
            End
'        Else
'            GoTo EndExportDXF_DWG
'        End If
    End If
    statusExportCAD2D = False
EndExportDXF_DWG:
    AUTOPARTLIST.Show
End Sub

Private Sub EXPORTCAD2D(CAD2D As String)
'Call CheckDXFFolderExists
Dim Pathdwg As String
    strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    Pathdwg = strGetCurrentPath + "\DXF-DWG\"
    Set EndProgram = New Class_ClickESC
    Select Case cbExportCAD2D
    Case "R2007"
        swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfVersion, swDxfFormat_R2007
    Case "R2010"
        swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfVersion, swDxfFormat_R2010
    Case "R2013"
        swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfVersion, swDxfFormat_R2013
    End Select
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfOutputFonts, 0
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfOutputLineStyles, 0
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swDXFDontShowMap, True
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfOutputNoScale, 1
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swDxfEndPointMerge, False
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swDxfEndPointMerge, False
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swDxfExportSplinesAsSplines, True
    If blnStatusExportExcel = False Then strFilePath = Pathdwg 'BrowseFolder
    If strFilePath <> "" Then
        If opbtAllToSeparateFileCAD2D Then
            swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, swDxfMultisheet_e.swDxfActiveSheetOnly
            For intSheetNumber = 0 To UBound(vSheetNames)
                EndProgram.EndByESCKey
                ProcessExportCAD2D intSheetNumber, CAD2D
            Next intSheetNumber
        ElseIf opbtAllInOneFileCAD2D Then
            swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, swDxfMultisheet_e.swDxfMultiSheet
            For intSheetNumber = 0 To UBound(vSheetNames)
                EndProgram.EndByESCKey
                swDraw.ActivateSheet (vSheetNames(intSheetNumber))
                Call SetSheetScaleCAD2D
            Next intSheetNumber
            If opbtDrawingNoCAD2D = True Then
                blnExportCAD2D = swModel.Extension.SaveAs(strFilePath & "\" & UCase(strPoCode) & "-" & strSoMaySoCum & "." & CAD2D, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Empty, Empty)
            ElseIf opbtPartnameCAD2D = True Then
                Call FilterPartNameFromPathName(swModel.GetPathName)
                blnExportCAD2D = swModel.Extension.SaveAs(strFilePath & "\" & strPartName & "." & CAD2D, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Empty, Empty)
            End If
        ElseIf opbtRangeCAD2D Then
            swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, swDxfMultisheet_e.swDxfActiveSheetOnly
            For i = 0 To iMemory2 - 1
                EndProgram.EndByESCKey
                intSheetNumber = CInt(intGetRange(i) - txtStartRenameSheet)
                ProcessExportCAD2D intSheetNumber, CAD2D
            Next i
        ElseIf opbtCurrentSheetCAD2D Then
            For intSheetNumber = 0 To UBound(vSheetNames)
                If vSheetNames(intSheetNumber) = swDraw.GetCurrentSheet.GetName Then Call ProcessExportCAD2D(intSheetNumber, CAD2D)
            Next intSheetNumber
        End If
        If opbtDXF = True Then
            swApp.SendMsgToUser2 "EXPORT FILE DXF COMPLETED!", swMbInformation, swMbOk
        ElseIf opbtDWG = True Then
            swApp.SendMsgToUser2 "EXPORT FILE DWG COMPLETED!", swMbInformation, swMbOk
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub ProcessExportCAD2D(intSheetNumber As Integer, CAD2D As String)
    swDraw.ActivateSheet vSheetNames(intSheetNumber)
    Call SetSheetScaleCAD2D
    If opbtDrawingNoCAD2D = True Then
        Call InfoDrawingNo
        blnExportCAD2D = swModel.Extension.SaveAs(strFilePath & "\" & strDrawingNo & "." & CAD2D, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Empty, Empty)
    ElseIf opbtPartnameCAD2D = True Then
        Set swView = swDraw.GetFirstView
        Set swView = swView.GetNextView
        strPathName = swView.GetReferencedModelName
        Call OpenModel(strPathName)
        Call InfoPartname(swOpenModel)
        Call SaveAndCloseModel3D(strPathName)
        blnExportCAD2D = swModel.Extension.SaveAs(strFilePath & "\" & strPartName & "." & CAD2D, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Empty, Empty)
    End If
End Sub

Private Sub SetSheetScaleCAD2D()
    Set swView = swDraw.GetFirstView
    vGetSheetProperties = swDraw.GetCurrentSheet.GetProperties2()
    j = 0
    Do While Not swView Is Nothing
        Dim swBaseView As View
        Set swBaseView = swView.GetBaseView
        If Not swBaseView Is Nothing Then
            vGetScaleRatio = swBaseView.ScaleRatio
            If vGetSheetProperties(0) = 8 Then
                swDraw.GetCurrentSheet.SetProperties2 swDwgPaperA3size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
            ElseIf vGetSheetProperties(0) = 9 Then
                swDraw.GetCurrentSheet.SetProperties2 swDwgPaperA2size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
            ElseIf vGetSheetProperties(0) = 10 Then
                swDraw.GetCurrentSheet.SetProperties2 swDwgPaperA1size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
            ElseIf vGetSheetProperties(0) = 11 Then
                swDraw.GetCurrentSheet.SetProperties2 swDwgPaperA0size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
            End If
            swBaseView.UseSheetScale() = True
            Exit Do
        End If
        j = j + 1
        Set swView = swView.GetNextView
    Loop
    'Set set scale for drawing have one view
    If j = 2 Then
        Set swView = swDraw.GetFirstView
        Set swView = swView.GetNextView
        vGetScaleRatio = swView.ScaleRatio
        If vGetSheetProperties(0) = 8 Then
            swDraw.GetCurrentSheet.SetProperties2 swDwgPaperA3size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
        ElseIf vGetSheetProperties(0) = 9 Then
            swDraw.GetCurrentSheet.SetProperties2 swDwgPaperA2size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
        ElseIf vGetSheetProperties(0) = 10 Then
            swDraw.GetCurrentSheet.SetProperties2 swDwgPaperA1size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
        ElseIf vGetSheetProperties(0) = 11 Then
            swDraw.GetCurrentSheet.SetProperties2 swDwgPaperA0size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
        End If
        swView.UseSheetScale() = True
    End If
End Sub

Private Sub strRangeCAD2D_Change()
    opbtRangeCAD2D.Value = True
End Sub

Function BrowseFolder(Optional Caption As String, Optional InitialFolder As String) As String
    Dim SH As Shell32.Shell
    Dim F As Shell32.Folder
    Const BIF_RETURNONLYFSDIRS As Long = &H1
    Set SH = New Shell32.Shell
    Set F = SH.BrowseForFolder(0&, Caption, BIF_RETURNONLYFSDIRS, InitialFolder)
    If Not F Is Nothing Then
        On Error Resume Next
        BrowseFolder = F.items.item.path
    End If
End Function

Private Sub cmdBACKUP_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    blnStatusBackup = True
    Set EndProgram = New Class_ClickESC
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.WindowState = xlMaximized
    Set xlWorkbook = xlApp.Workbooks.Add
    With xlWorkbook.Worksheets(1)
        'Write title for first row in Excel
        On Error Resume Next
        .Cells(1, 1).Value = "PART NAME"
        On Error Resume Next
        .Cells(1, 2).Value = "DRAWING NO."
        On Error Resume Next
        .Cells(1, 3).Value = "MATERIAL"
        On Error Resume Next
        .Cells(1, 4).Value = "TREATMENT"
        On Error Resume Next
        .Cells(1, 5).Value = "REMARK"
        On Error Resume Next
        .Cells(1, 6).Value = "FINISH"
        On Error Resume Next
        .Cells(1, 7).Value = "QTY"
        On Error Resume Next
        .Cells(1, 8).Value = "CLASS"
        On Error Resume Next
        .Cells(1, 9).Value = "ROUGHNESS Ra"
        On Error Resume Next
        .Cells(1, 10).Value = "CONFIGURATION NAME"
        On Error Resume Next
        .Cells(1, 11).Value = "FILE NAME"
        On Error Resume Next
        .Columns("A").NumberFormat = "@"
        On Error Resume Next
        .Columns("C").NumberFormat = "@"
        On Error Resume Next
        .Columns("D").NumberFormat = "@"
        On Error Resume Next
        .Columns("F").NumberFormat = "@"
        On Error Resume Next
        .Columns("I").NumberFormat = "@"
        On Error Resume Next
        .Columns("J").NumberFormat = "@"
        On Error Resume Next
        .Columns("K").NumberFormat = "@"
        For intSheetNumber = 0 To UBound(vSheetNames)
            EndProgram.EndByESCKey
            swDraw.ActivateSheet (vSheetNames(intSheetNumber))
            Set swView = swDraw.GetFirstView
            Set swView = swView.GetNextView
            strPathName = swView.GetReferencedModelName
            Call OpenModel(strPathName)
            'Write value into Excel from Model Configuration Manager
            Dim strValueBackup() As String
            ReDim strValueBackup(swModelConfigMgr.Count) As String
            For i = 0 To swModelConfigMgr.Count
                On Error Resume Next
                swModelConfigMgr.Get3 vModelConfigMgr(i), False, Empty, strValueBackup(i)
                Select Case UCase(vModelConfigMgr(i))
                Case "PART NAME"
                    InfoPartname swOpenModel
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 1).Value = UCase(strPartName)
                Case "DRAWING NO."
                    Call InfoDrawingNo
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 2).Value = strDrawingNo
                Case "MATERIAL"
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 3).Value = strValueBackup(i)
                Case "TREATMENT"
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 4).Value = UCase(strValueBackup(i))
                Case "REMARK"
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 5).Value = strValueBackup(i)
                Case "FINISH"
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 6).Value = strValueBackup(i)
                Case "QTY"
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 7).Value = strValueBackup(i)
                Case "CLASS"
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 8).Value = strValueBackup(i)
                Case "ROUGHNESS"
                    On Error Resume Next
                    .Cells(intSheetNumber + 2, 9).Value = strValueBackup(i)
                End Select
            Next i
            On Error Resume Next
            .Cells(intSheetNumber + 2, 10).Value = swView.ReferencedConfiguration
            FilterPartNameFromPathName strPathName
            On Error Resume Next
            .Hyperlinks.Add .Cells(intSheetNumber + 2, 11), strPathName, TextToDisplay:=strPartName
            Call SaveAndCloseModel3D(strPathName)
            On Error Resume Next
            .Columns("A:K").AutoFit
            On Error Resume Next
            .Columns("A:K").HorizontalAlignment = xlCenter
        Next intSheetNumber
    End With
    blnStatusBackup = False
    xlApp.WindowState = xlMinimized
    swApp.SendMsgToUser2 "CREATE FILE BACKUP COMPLETED!" & vbCrLf & "YOU HAVE TO SAVE THIS FILE BACKUP!", swMbInformation, swMbOk
    xlApp.WindowState = xlMaximized
    Unload AUTOPARTLIST
End Sub

Private Sub cmdRECOVERY_Click()
Call CheckRun
AUTOPARTLIST.Hide
    Dim blnStatusRecovery As Boolean
    blnStatusRecovery = True
    Call ErrorHandling
    If swApp.SendMsgToUser2("PLEASE FIND A FILE BACKUP!", swMbInformation, swMbYesNo) = swMbHitYes Then
        Call RECOVERY_PARTLIST
    Else
        GoTo EndRecovery
    End If
    blnStatusRecovery = False
EndRecovery:
    AUTOPARTLIST.Show
End Sub

Private Sub RECOVERY_PARTLIST()
    Set EndProgram = New Class_ClickESC
    'Find a backup file
    strFilter = "Excel File(*.xls; *.xlsx; *.xlsm)|*.xls; *.xlsx; *.xlsm|"
    strFileName = swApp.GetOpenFileName("Open file", Empty, strFilter, Empty, Empty, Empty)
    If strFileName = "" Then Exit Sub
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlWorkbook = xlApp.Workbooks.Open(strFileName)
    'Find number of rows with data
    Dim intCheckRow As Integer
    intCheckRow = 2
    With xlWorkbook.Worksheets(1)
        While .Cells(intCheckRow, 10).Value <> ""
            intCheckRow = intCheckRow + 1
        Wend
    End With
    'Get data from Excel
    ReDim strPartNameRecovery(intCheckRow - 2) As String
    ReDim strDrawingNoRecovery(intCheckRow - 2) As String
    ReDim strMaterialRecovery(intCheckRow - 2) As String
    ReDim strTreatmentRecovery(intCheckRow - 2) As String
    ReDim strRemarkRecovery(intCheckRow - 2) As String
    ReDim strFinishRecovery(intCheckRow - 2) As String
    ReDim strQtyRecovery(intCheckRow - 2) As String
    ReDim strClassRecovery(intCheckRow - 2) As String
    ReDim strRoughnessRecovery(intCheckRow - 2) As String
    ReDim strConfigurationNameRecovery(intCheckRow - 2) As String
    ReDim strFileNameRecovery(intCheckRow - 2) As String
    intCheckRow = 2
    With xlWorkbook.Worksheets(1)
        While intCheckRow - 2 < UBound(strPartNameRecovery())
            strPartNameRecovery(intCheckRow - 2) = .Cells(intCheckRow, 1).Value
            strDrawingNoRecovery(intCheckRow - 2) = .Cells(intCheckRow, 2).Value
            strMaterialRecovery(intCheckRow - 2) = .Cells(intCheckRow, 3).Value
            strTreatmentRecovery(intCheckRow - 2) = .Cells(intCheckRow, 4).Value
            strRemarkRecovery(intCheckRow - 2) = .Cells(intCheckRow, 5).Value
            strFinishRecovery(intCheckRow - 2) = .Cells(intCheckRow, 6).Value
            strQtyRecovery(intCheckRow - 2) = .Cells(intCheckRow, 7).Value
            strClassRecovery(intCheckRow - 2) = .Cells(intCheckRow, 8).Value
            strRoughnessRecovery(intCheckRow - 2) = .Cells(intCheckRow, 9).Value
            strConfigurationNameRecovery(intCheckRow - 2) = .Cells(intCheckRow, 10).Value
            strFileNameRecovery(intCheckRow - 2) = .Cells(intCheckRow, 11).Value
            intCheckRow = intCheckRow + 1
        Wend
    End With
    xlWorkbook.Close True
    xlApp.Quit
    'Write recovery data in model custom property
    Call TurnOffAutoSaved
    swApp.FrameState = swWindowMaximized
    intFrequenceUnequal = 0
    intCheckExistPartName = 0
    Set swDraw = swModel
    vSheetNames = swDraw.GetSheetNames()
    For intSheetNumber = 0 To UBound(vSheetNames)
        EndProgram.EndByESCKey
        swDraw.ActivateSheet (vSheetNames(intSheetNumber))
        Set swView = swDraw.GetFirstView
        Set swView = swView.GetNextView
        intCheckExistPartName = 0
        If Not swView Is Nothing Then
            strPathName = swView.GetReferencedModelName
            Call FilterPartNameFromPathName(strPathName)
            For i = 0 To UBound(strFileNameRecovery())
                If UCase(strPartName) = UCase(strFileNameRecovery(i)) And UCase(swView.ReferencedConfiguration) = UCase(strConfigurationNameRecovery(i)) Then
                    Call ProcessRecovery
                    'Export info for not flat-pattern view
                    Set swView = swDraw.GetFirstView
                    Set swView = swView.GetNextView
                    If Not swView Is Nothing Then
                        If swView.IsFlatPatternView() Then
                            Set swView = swView.GetNextView
                            Do While Not swView Is Nothing
                                If swView.IsFlatPatternView() = False Then
                                    Call ProcessRecovery
                                    Exit Do
                                End If
                                Set swView = swView.GetNextView
                            Loop
                        End If
                    End If
                Else
                    intCheckExistPartName = intCheckExistPartName + 1
                End If
            Next i
        End If
        If intCheckExistPartName - 1 = UBound(strPartNameRecovery()) Then intFrequenceUnequal = intFrequenceUnequal + 1
    Next intSheetNumber
    Call TurnOnAutoSaved
    'Check whether selected backup file is correct
    If intFrequenceUnequal <> 0 Then
        swApp.SendMsgToUser2 "This file backup seems inexactly!" & vbCrLf & "You need use a exact backup file!", swMbWarning, swMbOk
    Else
        swApp.SendMsgToUser2 "RECOVERY COMPLETED!", swMbInformation, swMbOk
        swApp.SendMsgToUser2 "Sort of sheets could have been changed!" & vbCrLf & "You should run AUTO PARLIST 2D again to edit Drawing No.!", swMbWarning, swMbOk
    End If
End Sub

Private Sub ProcessRecovery()
    Set swOpenModel = swApp.ActivateDoc3(strPathName, False, 2, 0)
    swOpenModel.ShowConfiguration2 swView.ReferencedConfiguration
    Set swModelView = swOpenModel.GetFirstModelView
    swModelView.FrameState = swWindowMaximized
LoopClickMouseRecovery:
    Call SetClickMouse
    Call SetConfigModel
    blnStatusSetDrawingNo = False
    For j = 0 To swModelConfigMgr.Count
        On Error Resume Next
        Select Case UCase(vModelConfigMgr(j))
        Case "PART NAME"
           swModelConfigMgr.Set "PART NAME", strPartNameRecovery(i)
        Case "DRAWING NO."
           swModelConfigMgr.Set "DRAWING NO.", strDrawingNoRecovery(i)
            blnStatusSetDrawingNo = True
        Case "MATERIAL"
           swModelConfigMgr.Set "MATERIAL", strMaterialRecovery(i)
        Case "TREATMENT"
           swModelConfigMgr.Set "TREATMENT", strTreatmentRecovery(i)
        Case "REMARK"
           swModelConfigMgr.Set "REMARK", strRemarkRecovery(i)
        Case "FINISH"
           swModelConfigMgr.Set "FINISH", strFinishRecovery(i)
        Case "QTY"
           swModelConfigMgr.Set "QTY", strQtyRecovery(i)
        Case "CLASS"
           swModelConfigMgr.Set "CLASS", strClassRecovery(i)
        Case "ROUGHNESS"
           swModelConfigMgr.Set "ROUGHNESS", strRoughnessRecovery(i)
        End Select
    Next j
    If blnStatusSetDrawingNo = False Then
        If Xposition <= DisplayWidth Then
            TheMouse.Move DisplayWidth * 2 - 260, 160, swMouse_MouseMove
            GoTo LoopClickMouseRecovery
        Else
            TheMouse.Move DisplayWidth - 260, 160, swMouse_MouseMove
            GoTo LoopClickMouseRecovery
        End If
    End If
    swOpenModel.Save3 swSaveAsOptions_Silent, Empty, Empty
    swApp.CloseDoc strPathName
End Sub

Private Sub cmdAutoRemoveExistDrawingNo_Click()
    blnAutoRemoveExistDrawingNo = True
    AUTOPARTLIST.Hide
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Set EndProgram = New Class_ClickESC
    swDraw.ActivateSheet vSheetNames(0)
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    If Not swView Is Nothing Then
        strPathName = swView.GetReferencedModelName
        If UCase(Right(strPathName, 6)) = "SLDASM" Then
            TurnOffAutoSaved
            Set swOpenModel = swApp.ActivateDoc3(strPathName, True, 2, 0)
            Set swAsm = swOpenModel
            swAsm.ResolveAllLightWeightComponents False
            vComps = swAsm.GetComponents(False)
            For i = LBound(vComps) To UBound(vComps)
                EndProgram.EndByESCKey
                Set swComp = vComps(i)
                If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.IsSuppressed() = False Then
                    Call OpenModel(swComp.GetPathName)
                    If Not swOpenModel Is Nothing Then
                        If swOpenModel.GetConfigurationCount > 1 Then
                            Call TurnOffAutoSaved
                            Set swOpenModel = swApp.ActivateDoc3(swComp.GetPathName, True, 2, 0)
                            vconfigNames = swOpenModel.GetConfigurationNames()
                            For Each vconfigName In vconfigNames
                                swOpenModel.ShowConfiguration2 vconfigName
                                Call SetConfigModel
                                If Not swModelConfigMgr Is Nothing Then
                                    For j = 0 To swModelConfigMgr.Count
                                        On Error Resume Next
                                        If UCase(vModelConfigMgr(j)) = "DRAWING NO." Then swModelConfigMgr.Set "DRAWING NO.", ""
                                    Next j
                                End If
                            Next
                            blnCheckOpenModel3D = True
                        Else
                            If Not swModelConfigMgr Is Nothing Then
                                For j = 0 To swModelConfigMgr.Count
                                    On Error Resume Next
                                    If UCase(vModelConfigMgr(j)) = "DRAWING NO." Then swModelConfigMgr.Set "DRAWING NO.", ""
                                Next j
                            End If
                        End If
                    End If
                    Call SaveAndCloseModel3D(swComp.GetPathName)
                End If
            Next i
            swApp.CloseDoc strPathName
        End If
    End If
    swModel.EditRebuild3
    swApp.SendMsgToUser2 "AUTO REMOVE EXIST DRAWING NO FINISH!", swMbInformation, swMbOk
    End
End Sub

Private Sub cmdCheckExcessDrawing_Click()
Call CheckRun
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    AUTOPARTLIST.Hide
    ListExcessDrawings.Show
End Sub

Private Sub cmdCheckMissingDrawings_Click()
    Call CheckRun
    AUTOPARTLIST.Hide
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Set EndProgram = New Class_ClickESC
    swDraw.ActivateSheet vSheetNames(0)
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    If Not swView Is Nothing Then
        strPathName = swView.GetReferencedModelName
        If UCase(Right(strPathName, 6)) = "SLDASM" Then
            TurnOffAutoSaved
            Set swOpenModel = swApp.ActivateDoc3(strPathName, True, 2, 0)
            Set swAsm = swOpenModel
            vComps = swAsm.GetComponents(False)
            Dim strCheckConfigurationName3D() As String
            ReDim strCheckConfigurationName3D(UBound(vComps)) As String
            ReDim strGetParentCompSubFeat(UBound(vComps)) As String
            ReDim strCheckPartname(UBound(vComps)) As String
            Dim strGetExtension3D() As String
            ReDim strGetExtension3D(UBound(vComps)) As String
            Dim strPathname3D() As String
            ReDim strPathname3D(UBound(vComps)) As String
            'Collect all components 3D
            intCountAllComponent = 0
            For i = 0 To UBound(vComps)
                EndProgram.EndByESCKey
                Set swComp = vComps(i)
                If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.ExcludeFromBOM() = False Then
                    strPathname3D(intCountAllComponent) = swComp.GetPathName
                    Call FilterPartNameFromPathName(swComp.GetPathName)
                    strCheckPartname(intCountAllComponent) = UCase(strPartName)
                    strGetExtension3D(intCountAllComponent) = UCase(Right(swComp.GetPathName, 6))
                    strCheckConfigurationName3D(intCountAllComponent) = UCase(swComp.ReferencedConfiguration)
                    Dim swParentComp As Component2
                    Set swParentComp = swComp.GetParent
                    If Not swParentComp Is Nothing Then
                        Call FilterPartNameFromPathName(swParentComp.GetPathName)
                        strGetParentCompSubFeat(intCountAllComponent) = UCase(strPartName)
                    Else
                        Call FilterPartNameFromPathName(swOpenModel.GetPathName)
                        strGetParentCompSubFeat(intCountAllComponent) = UCase(strPartName)
                    End If
                    intCountAllComponent = intCountAllComponent + 1
                End If
            Next i
            'Remove duplicate data for all components
            For i = 0 To intCountAllComponent - 2
CheckDuplicatePartnameMissingComponents:
                For j = i + 1 To intCountAllComponent - 1
                    If strCheckPartname(j) = strCheckPartname(i) And strCheckConfigurationName3D(j) = strCheckConfigurationName3D(i) And strGetExtension3D(j) = strGetExtension3D(i) Then
                        For k = j To intCountAllComponent - 2
                            strPathname3D(k) = strPathname3D(k + 1)
                            strCheckPartname(k) = strCheckPartname(k + 1)
                            strCheckConfigurationName3D(k) = strCheckConfigurationName3D(k + 1)
                            strGetExtension3D(k) = strGetExtension3D(k + 1)
                            strGetParentCompSubFeat(k) = strGetParentCompSubFeat(k + 1)
                        Next k
                        intCountAllComponent = intCountAllComponent - 1
                    End If
                If strCheckPartname(j) = strCheckPartname(i) And strCheckConfigurationName3D(j) = strCheckConfigurationName3D(i) And strGetExtension3D(j) = strGetExtension3D(i) Then GoTo CheckDuplicatePartnameMissingComponents
                Next j
            Next i
            'Collect standard device from New Folder
            Set swOpenModel = swApp.ActiveDoc
            Call CollectStandardDevice
            'Save and close 3D
            Dim strPartNameTotal As String
            FilterPartNameFromPathName swOpenModel.GetPathName
            strPartNameTotal = strPartName
            swOpenModel.Save3 swSaveAsOptions_Silent, Empty, Empty
            swApp.CloseDoc swOpenModel.GetPathName
            'Collect all components 2D
            Dim strGetPartName2D() As String
            ReDim strGetPartName2D(UBound(vSheetNames)) As String
            Dim strCheckConfigurationName2D() As String
            ReDim strCheckConfigurationName2D(UBound(vSheetNames)) As String
            Dim strGetExtension2D() As String
            ReDim strGetExtension2D(UBound(vSheetNames)) As String
            j = 0
            If UBound(vSheetNames) > 0 Then
                For i = 1 To UBound(vSheetNames)
                    EndProgram.EndByESCKey
                    swDraw.ActivateSheet vSheetNames(i)
                    Set swView = swDraw.GetFirstView
                    Set swView = swView.GetNextView
                    If Not swView Is Nothing Then
                        strPathName = swView.GetReferencedModelName
                        FilterPartNameFromPathName strPathName
                        strGetPartName2D(j) = UCase(strPartName)
                        strCheckConfigurationName2D(j) = UCase(swView.ReferencedConfiguration)
                        strGetExtension2D(j) = UCase(Right(swView.GetReferencedModelName, 6))
                        j = j + 1
                    End If
                Next i
            End If
            'Export missing components
            Set xlApp = CreateObject("Excel.Application")
            xlApp.Visible = True
            xlApp.WindowState = xlMaximized
            Set xlWorkbook = xlApp.Workbooks.Add
            With xlWorkbook.Worksheets(1)
                On Error Resume Next
                .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH CHI TI" & ChrW(202) & ChrW(769) & "T CH" & ChrW(431) & "A T" & ChrW(193) & "CH BA" & ChrW(777) & "N VE" & ChrW(771)
                On Error Resume Next
                .Range(.Cells(1, 1), .Cells(1, 4)).Merge
                On Error Resume Next
                .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartNameTotal
                On Error Resume Next
                .Range(.Cells(2, 1), .Cells(2, 4)).Merge
                On Error Resume Next
                .Cells(3, 1).Value = "STT"
                On Error Resume Next
                .Cells(3, 2).Value = "T" & ChrW(202) & "N CHI TI" & ChrW(7870) & "T"
                On Error Resume Next
                .Cells(3, 3).Value = "CHI TI" & ChrW(202) & ChrW(769) & "T TRONG CU" & ChrW(803) & "M"
                On Error Resume Next
                .Cells(3, 4).Value = "T" & ChrW(202) & "N C" & ChrW(7844) & "U H" & ChrW(204) & "NH"
                On Error Resume Next
                .Range("A1:D3").Font.Bold = True
                On Error Resume Next
                .Columns("B").NumberFormat = "@"
                On Error Resume Next
                .Columns("C").NumberFormat = "@"
                On Error Resume Next
                .Columns("D").NumberFormat = "@"
                Dim blnCheckMissingComponent As Boolean
                Dim strGetMissingComponentPath() As String
                ReDim strGetMissingComponentPath(UBound(vComps)) As String
                m = 0
                For i = 0 To intCountAllComponent - 1
                    j = 0
                    For k = 0 To intCountStandardDevice - 1
                        If UCase(strStandardDevice(k)) = strCheckPartname(i) Or UCase(strStandardDevice(k)) = UCase(strGetParentCompSubFeat(i)) Then j = j + 1
                    Next k
                    For k = 0 To intCountChildrenStandardDevice - 1
                        If UCase(strGetSubFeatFolder(k)) = strCheckPartname(i) Then j = j + 1
                    Next k
                    n = 0
                    For l = 0 To intCountSemiStandardDevice - 1
                        If strCheckPartname(i) = UCase(strSemiStandardDevice(l)) Then n = n + 1
                    Next l
                    If j = 0 Or n > 0 Then
                        blnCheckMissingComponent = False
                        For j = 0 To UBound(vSheetNames)
                            If strCheckPartname(i) = strGetPartName2D(j) And strCheckConfigurationName3D(i) = strCheckConfigurationName2D(j) And strGetExtension3D(i) = strGetExtension2D(j) Then blnCheckMissingComponent = True
                        Next j
                        If blnCheckMissingComponent = False Then
                            On Error Resume Next
                            .Cells(m + 4, 1).Value = m + 1
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(m + 4, 2), strPathname3D(i), TextToDisplay:=strCheckPartname(i)
                            On Error Resume Next
                            .Cells(m + 4, 3) = strGetParentCompSubFeat(i)
                            On Error Resume Next
                            .Cells(m + 4, 4).Value = strCheckConfigurationName3D(i)
                            strGetMissingComponentPath(m) = strPathname3D(i)
                            m = m + 1
                        End If
                    End If
                Next i
                If m = 0 Then
                    On Error Resume Next
                    .Range(.Cells(m + 4, 1), .Cells(m + 4, 4)).Merge
                    On Error Resume Next
                    .Cells(m + 4, 1).Value = "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y b" & ChrW(7843) & "n v" & ChrW(7869) & " thi" & ChrW(7871) & "u!"
                    On Error Resume Next
                    .Cells(m + 4, 1).Font.Bold = True
                    On Error Resume Next
                    .Cells(m + 4, 1).Font.Italic = True
                    On Error Resume Next
                    .Cells(m + 4, 1).Font.Color = RGB(255, 0, 0)
                End If
                On Error Resume Next
                .Columns("A:D").AutoFit
                On Error Resume Next
                .Columns("A:D").HorizontalAlignment = xlCenter
                On Error Resume Next
                .Range(.Cells(1, 1), .Cells(m + 3, 4)).Borders.LineStyle = xlContinuous
            End With
            If m > 0 Then
                If swApp.SendMsgToUser2("This program has found " & m & " missing components." & vbCrLf & "Do you only view list Excel without opening all missing components? Please push Yes button to accept, or push No button to open all missing components now!", swMbQuestion, swMbYesNo) = swMbHitNo Then
                    For i = 0 To m - 1
                        swApp.ActivateDoc3 strGetMissingComponentPath(i), False, 2, 0
                    Next i
                End If
                swApp.FrameState = swWindowMinimized
            End If
        End If
    End If
    End
End Sub

Private Sub AutoPoCodeForExternalComponent()
    Set EndProgram = New Class_ClickESC
    swModel.EditRebuild3
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotation_BillOfMaterials And strExtension = "SLDASM" Then
            Set swBOMTable = swTable
            Dim ItemNumber As String
            Dim PartNumber As String
            Dim GetPathComponentBOM As Variant
            For i = 0 To swTable.RowCount - 3
                EndProgram.EndByESCKey
                If Len(swTable.Text(i, 2)) > 1 Then
                    Dim intFirstHorizonalMark As Integer
                    Dim intSecondHorizonalMark As Integer
                    Dim intLastHorizonalMark As Integer
                    Dim strDrawingNoExternalComponent As String
                    On Error Resume Next
                    If InStr(1, swTable.Text(i, 2), "-", vbTextCompare) <> 0 Then
                        intFirstHorizonalMark = InStr(1, swTable.Text(i, 2), "-", vbTextCompare)
                        On Error Resume Next
                        If InStr(intFirstHorizonalMark + 1, swTable.Text(i, 2), "-", vbTextCompare) <> 0 Then
                            intSecondHorizonalMark = InStr(intFirstHorizonalMark + 1, swTable.Text(i, 2), "-", vbTextCompare)
                        End If
                    End If
                    On Error Resume Next
                    If InStrRev(swTable.Text(i, 2), "-", -1, vbTextCompare) <> 0 Then intLastHorizonalMark = InStrRev(swTable.Text(i, 2), "-", -1, vbTextCompare)
                    Dim strSomaySocumExternalComponent As String
                    On Error Resume Next
                    strSomaySocumExternalComponent = Mid(swTable.Text(i, 2), intSecondHorizonalMark + 1, intLastHorizonalMark - intSecondHorizonalMark - 1)
                    If StrComp(strSoMaySoCum, strSomaySocumExternalComponent, vbTextCompare) <> 0 Then
                        On Error Resume Next
                        strDrawingNoExternalComponent = Right(swTable.Text(i, 2), Len(swTable.Text(i, 2)) - intSecondHorizonalMark + 1)
                        strDrawingNoExternalComponent = UCase(strPoCode) & strDrawingNoExternalComponent
                        GetPathComponentBOM = swBOMTable.GetModelPathNames(i, ItemNumber, PartNumber)
                        strPathName = GetPathComponentBOM(0)
                        FilterPartNameFromPathName strPathName
                        TurnOffAutoSaved
                        Set swOpenModel = swApp.ActivateDoc3(strPathName, True, 2, 0)
                        If strPartName = swTable.Text(i, 1) Then
                            swOpenModel.ShowConfiguration2 "Default"
                        Else
                            On Error Resume Next
                            swOpenModel.ShowConfiguration2 swTable.Text(i, 1)
                        End If
                        blnCheckOpenModel3D = True
                        SetConfigModel
                        If Not swModelConfigMgr Is Nothing Then
                            For m = 0 To swModelConfigMgr.Count
                                On Error Resume Next
                                Select Case UCase(vModelConfigMgr(m))
                                Case "PART NAME"
                                    swModelConfigMgr.Set "PART NAME", strPartName
                                Case "DRAWING NO."
                                    swModelConfigMgr.Set "DRAWING NO.", strDrawingNoExternalComponent
                                End Select
                            Next m
                        End If
                        SaveAndCloseModel3D strPathName
                    End If
                End If
            Next i
        End If
        Set swTable = swTable.GetNext
    Wend
    End
End Sub

Private Sub cmdUndoAutoNote_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Dim vAttachedEnts As Variant
    Dim swEnt As Entity
    Set swView = swDraw.GetFirstView
    While Not swView Is Nothing
        vNote = swView.GetNotes
        For i = 0 To UBound(vNote)
            Set swNote = vNote(i)
            If swNote.HasBalloon() = False And swNote.GetLeaderCount() > 0 Then
                Set swAnn = swNote.GetAnnotation
                vAttachedEnts = swAnn.GetAttachedEntities3
                If IsEmpty(vAttachedEnts) = False Then
                    swAnn.Select3 False, Nothing
                    Set swNote = swModel.Extension.EditBalloonProperties2(swBS_Circular, swBF_2Chars, swBalloonTextItemNumber, Empty, Empty, Empty, Empty, False, Empty, Empty, Empty)
                End If
            End If
        Next i
        Set swView = swView.GetNextView
    Wend
    swApp.SendMsgToUser2 "UNDO AUTO NOTE COMPLETED!", swMbInformation, swMbOk
    Unload Me
End Sub

Private Sub cmdRenameComponents_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    RenameComponents.Show
End Sub

Private Sub cmdLableMaker_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    LABLE_MAKER.Show
End Sub

Private Sub cmdCheckUncheckBoxExcludedFromBom_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    Call ErrorHandling2
    Dim swSelMgr As SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    If swSelMgr.GetSelectedObjectCount2(0) = 0 Then
        swApp.SendMsgToUser2 "No selections detected!", swMbWarning, swMbOk
        Exit Sub
    Else
        Set swAsm = swModel
        For i = 1 To swSelMgr.GetSelectedObjectCount2(0)
            Set swComp = swSelMgr.GetSelectedObjectsComponent3(i, -1)
            If Not swComp Is Nothing Then
                If swComp.ExcludeFromBOM = False Then
                    swComp.ExcludeFromBOM = True
                    flag = True
                Else
                    swComp.ExcludeFromBOM = False
                    flag = False
                End If
            End If
        Next i
    End If
    swModel.EditRebuild3
    If flag = False Then
        swApp.SendMsgToUser2 "UNCHECKBOX EXCLUDE FROM BOM COMPLETED!" & vbLf & "This component will be calculated for your quantities total!", swMbInformation, swMbOk
    Else
        swApp.SendMsgToUser2 "CHECKBOX EXCLUDE FROM BOM COMPLETED!" & vbLf & "This component will NOT be calculated for your quantities total!", swMbWarning, swMbOk
    End If
    End
End Sub

Private Sub cmdCheckComponentsExcludedFromBOM_Click()
Call CheckRun
    Call ErrorHandling2
    AUTOPARTLIST.Hide
    Set EndProgram = New Class_ClickESC
    Set swAsm = swModel
    'swAsm.ResolveAllLightWeightComponents False
    vComps = swAsm.GetComponents(False)
    Dim PressPageUp_Down As New Class_ClickPageUP_DOWN
    Set PressPageUp_Down = New Class_ClickPageUP_DOWN
    Dim intTypeKey As Integer
    Dim vvComps() As Variant
    ReDim vvComps(UBound(vComps)) As Variant
    Dim rootNode As SldWorks.TreeControlItem
    Set rootNode = swModel.FeatureManager.GetFeatureTreeRootItem2(swFeatMgrPaneBottom)
    rootNode.Expanded = True
    j = 0
    For i = LBound(vComps) To UBound(vComps)
        EndProgram.EndByESCKey
        Set swComp = vComps(i)
        On Error Resume Next
        If swComp.ExcludeFromBOM() Then
            Set vvComps(j) = swComp
            j = j + 1
        End If
    Next i
    If j = 0 Then
        swApp.SendMsgToUser2 "HAVE NO COMPONENTS EXCLUDED FROM BOM!", swMbInformation, swMbOk
        End
    Else
        swApp.SendMsgToUser2 "You have to press button PageUp or PageDown on your keyboard to traverse all components Excluded From BOM!", swMbInformation, swMbOk
    End If
    For i = 0 To j - 1
LoopWaitKeyPageUp:
        EndProgram.EndByESCKey
        If i < 0 Then i = 0
        On Error Resume Next
        vvComps(i).Select3 False, Nothing
        intTypeKey = 0
LoopWaitKeyPageDown:
        PressPageUp_Down.WaitUntilPageUp_Down intTypeKey
        If intTypeKey = 0 Then
            GoTo LoopWaitKeyPageDown
        ElseIf intTypeKey = 1 Then
            i = i - 1
            GoTo LoopWaitKeyPageUp
        End If
    Next i
LoopWaitKeyPageDownForEnd:
    intTypeKey = 0
    PressPageUp_Down.WaitUntilPageUp_Down intTypeKey
    If intTypeKey = 2 Then
        swApp.SendMsgToUser2 "TRAVERSE ALL COMPONENTS EXCLUDED FROM BOM COMPLETED!", swMbInformation, swMbOk
    Else
        GoTo LoopWaitKeyPageDownForEnd
    End If
    If swApp.SendMsgToUser2("Export Excel all components Excluded From BOM?", swMbWarning, swMbYesNo) = swMbHitYes Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add
        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"
            On Error Resume Next
            .Columns("C").NumberFormat = "@"
            On Error Resume Next
            .Columns("D").NumberFormat = "@"
            On Error Resume Next
            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH LO" & ChrW(7840) & "I TR" & ChrW(7914) & " " & ChrW(272) & ChrW(7882) & "NH M" & ChrW(7912) & "C V" & ChrW(7852) & "T LI" & ChrW(7878) & "U"
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(1, 4)).Merge
            On Error Resume Next
            .Cells(1, 1).Font.Size = 10
            On Error Resume Next
            .Cells(2, 1).Value = "STT"
            On Error Resume Next
            .Cells(2, 2).Value = "T" & ChrW(202) & "N CHI TI" & ChrW(7870) & "T"
            On Error Resume Next
            .Cells(2, 3).Value = "CHI TI" & ChrW(7870) & "T TRONG CU" & ChrW(803) & "M"
            On Error Resume Next
            .Cells(2, 4).Value = "T" & ChrW(202) & "N C" & ChrW(7844) & "U H" & ChrW(204) & "NH"
            On Error Resume Next
            .Range("A1:D2").Font.Bold = True
            For i = 0 To j - 1
                EndProgram.EndByESCKey
                On Error Resume Next
                .Cells(i + 3, 1).Value = i + 1
                On Error Resume Next
                FilterPartNameFromPathName vvComps(i).GetPathName
                .Cells(i + 3, 2).Value = strPartName
                Dim swParentComp As Component2
                Set swParentComp = vvComps(i).GetParent
                If Not swParentComp Is Nothing Then
                    Set swConfigMgr = swParentComp.GetModelDoc2.ConfigurationManager
                    Set swConfig = swConfigMgr.ActiveConfiguration
                    Set swParentComp = swConfig.GetRootComponent3(True)
                    FilterPartNameFromPathName swParentComp.GetPathName
                    On Error Resume Next
                    .Cells(i + 3, 3).Value = strPartName
                End If
                On Error Resume Next
                .Cells(i + 3, 4).Value = vvComps(i).ReferencedConfiguration
            Next i
            On Error Resume Next
            .Columns("A:D").AutoFit
            On Error Resume Next
            .Columns("A:D").HorizontalAlignment = xlCenter
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(j + 2, 4)).Borders.LineStyle = xlContinuous
        End With
    End If
    End
End Sub

Private Sub cmdCOMPARE_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    Set EndProgram = New Class_ClickESC
    Call ErrorHandling2
    Call TurnOffAutoSaved
    swApp.FrameState = swWindowMaximized
    Set swAsm = swModel
    swAsm.ResolveAllLightWeightComponents False
    vComps = swAsm.GetComponents(False)
    ReDim strCheckPartname(UBound(vComps)) As String
    Dim intTotalQuantity() As Integer
    ReDim intTotalQuantity(UBound(vComps)) As Integer
    Dim strCheckConfigurationName() As String
    ReDim strCheckConfigurationName(UBound(vComps)) As String
    ReDim strCheckExtension(UBound(vComps)) As String
    Dim strGetPathname() As String
    ReDim strGetPathname(UBound(vComps)) As String
    ReDim strGetParentCompSubFeat(UBound(vComps)) As String
    Dim swParentComp As Component2
    Dim strGetQTY() As String
    ReDim strGetQTY(UBound(vComps)) As String
    'Collect all components
    intCountAllComponent = 0
    Dim strQty As String
    For i = LBound(vComps) To UBound(vComps)
        EndProgram.EndByESCKey
        Set swComp = vComps(i)
        If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.IsSuppressed() = False And swComp.ExcludeFromBOM = False Then
            strExtension = Right(UCase(swComp.GetPathName), 6)
            strCheckExtension(intCountAllComponent) = strExtension
            Call FilterPartNameFromPathName(swComp.GetPathName)
            strCheckPartname(intCountAllComponent) = strPartName
            strCheckConfigurationName(intCountAllComponent) = swComp.ReferencedConfiguration
            strGetPathname(intCountAllComponent) = swComp.GetPathName
            Set swParentComp = swComp.GetParent
            If Not swParentComp Is Nothing Then
                Call FilterPartNameFromPathName(swParentComp.GetPathName)
                strGetParentCompSubFeat(intCountAllComponent) = strPartName
            Else
                Call FilterPartNameFromPathName(swModel.GetPathName)
                strGetParentCompSubFeat(intCountAllComponent) = strPartName
            End If
            intTotalQuantity(intCountAllComponent) = 1
            Set swOpenModel = swComp.GetModelDoc2
            If swOpenModel.GetConfigurationCount > 1 Then
                intCountConfigNameDefault = 0
                vconfigNames = swOpenModel.GetConfigurationNames()
                For Each vconfigName In vconfigNames
                    strConfigName = vconfigName
                    If InStr(1, strConfigName, "Default", vbTextCompare) <> 0 Then intCountConfigNameDefault = intCountConfigNameDefault + 1
                Next
                If intCountConfigNameDefault <> UBound(vconfigNames) + 1 Then
                    Set swOpenModel = swApp.ActivateDoc3(swComp.GetPathName, False, 2, 0)
                    swOpenModel.ShowConfiguration2 swComp.ReferencedConfiguration
                End If
            End If
            Call SetConfigModel
            If Not swModelConfigMgr Is Nothing Then
                For j = 0 To swModelConfigMgr.Count
                    On Error Resume Next
                    If UCase(vModelConfigMgr(j)) = "QTY" Then swModelConfigMgr.Get3 "QTY", True, Empty, strQty
                Next j
            End If
            If swOpenModel.GetConfigurationCount > 1 Then
                swApp.CloseDoc swComp.GetPathName
            End If
            strGetQTY(intCountAllComponent) = strQty
            strQty = ""
            intCountAllComponent = intCountAllComponent + 1
        End If
    Next
    'Remove duplicate data and calculate QTY for all components
    For i = 0 To intCountAllComponent - 2
CheckDuplicatePartnameQTY:
        EndProgram.EndByESCKey
        For j = i + 1 To intCountAllComponent - 1
            If strCheckPartname(j) = strCheckPartname(i) And strCheckConfigurationName(j) = strCheckConfigurationName(i) And strCheckExtension(j) = strCheckExtension(i) Then
                intTotalQuantity(i) = intTotalQuantity(i) + 1
                For k = j To intCountAllComponent - 2
                    strCheckPartname(k) = strCheckPartname(k + 1)
                    strCheckConfigurationName(k) = strCheckConfigurationName(k + 1)
                    strCheckExtension(k) = strCheckExtension(k + 1)
                    intTotalQuantity(k) = intTotalQuantity(k + 1)
                    strGetPathname(k) = strGetPathname(k + 1)
                    strGetParentCompSubFeat(k) = strGetParentCompSubFeat(k + 1)
                    strGetQTY(k) = strGetQTY(k + 1)
                Next k
                intCountAllComponent = intCountAllComponent - 1
            End If
        If strCheckPartname(j) = strCheckPartname(i) And strCheckConfigurationName(j) = strCheckConfigurationName(i) And strCheckExtension(j) = strCheckExtension(i) Then GoTo CheckDuplicatePartnameQTY
        Next j
    Next i
    'Collect component exclude from BOM
    Dim strGetPartnameExcludeFromBOM() As String
    ReDim strGetPartnameExcludeFromBOM(UBound(vComps)) As String
    Dim strGetExtensionExcludeFromBOM() As String
    ReDim strGetExtensionExcludeFromBOM(UBound(vComps)) As String
    Dim intTotalQuantityExcludeFromBOM() As Integer
    ReDim intTotalQuantityExcludeFromBOM(UBound(vComps)) As Integer
    Dim strCheckConfigurationNameExcludeFromBOM() As String
    ReDim strCheckConfigurationNameExcludeFromBOM(UBound(vComps)) As String
    Dim intCountAllComponentExcludeFromBOM As Integer
    intCountAllComponentExcludeFromBOM = 0
    Dim swCompModel As ModelDoc2
    For i = LBound(vComps) To UBound(vComps)
        EndProgram.EndByESCKey
        Set swComp = vComps(i)
        Set swCompModel = swComp.GetModelDoc2
        strExtension = Right(UCase(swComp.GetPathName), 6)
        On Error Resume Next
        If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And strExtension = "SLDASM" And swComp.ExcludeFromBOM Then
            Dim vCompExcludeFromBOM As Variant
            Set swAsm = swCompModel
            vCompExcludeFromBOM = swAsm.GetComponents(False)
            For j = LBound(vCompExcludeFromBOM) To UBound(vCompExcludeFromBOM)
                Set swComp = vCompExcludeFromBOM(j)
                Set swParentComp = swComp.GetParent
                Call FilterPartNameTopLevelTemp(swComp.Name2)
                If blnCheckFirstSlashes = True Then
                    On Error Resume Next
                    If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.ExcludeFromBOM = False And swParentComp.ExcludeFromBOM = False Then
                        strCheckConfigurationNameExcludeFromBOM(intCountAllComponentExcludeFromBOM) = swComp.ReferencedConfiguration
                        Call FilterPartNameFromPathName(swComp.GetPathName)
                        strGetPartnameExcludeFromBOM(intCountAllComponentExcludeFromBOM) = strPartName
                        strExtension = Right(UCase(swComp.GetPathName), 6)
                        strGetExtensionExcludeFromBOM(intCountAllComponentExcludeFromBOM) = strExtension
                        intTotalQuantityExcludeFromBOM(intCountAllComponentExcludeFromBOM) = 1
                        intCountAllComponentExcludeFromBOM = intCountAllComponentExcludeFromBOM + 1
                    End If
                Else
                    On Error Resume Next
                    If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.ExcludeFromBOM = False Then
                        strCheckConfigurationNameExcludeFromBOM(intCountAllComponentExcludeFromBOM) = swComp.ReferencedConfiguration
                        Call FilterPartNameFromPathName(swComp.GetPathName)
                        strGetPartnameExcludeFromBOM(intCountAllComponentExcludeFromBOM) = strPartName
                        strExtension = Right(UCase(swComp.GetPathName), 6)
                        strGetExtensionExcludeFromBOM(intCountAllComponentExcludeFromBOM) = strExtension
                        intTotalQuantityExcludeFromBOM(intCountAllComponentExcludeFromBOM) = 1
                        intCountAllComponentExcludeFromBOM = intCountAllComponentExcludeFromBOM + 1
                    End If
                End If
            Next
        End If
    Next
    'Remove duplicate data and calculate QTY for components exclude from BOM
    For i = 0 To intCountAllComponentExcludeFromBOM - 2
CheckDuplicatePartnameQTYExcludeFromBOM:
        EndProgram.EndByESCKey
        For j = i + 1 To intCountAllComponentExcludeFromBOM - 1
            If strGetPartnameExcludeFromBOM(j) = strGetPartnameExcludeFromBOM(i) And strCheckConfigurationNameExcludeFromBOM(j) = strCheckConfigurationNameExcludeFromBOM(i) And strGetExtensionExcludeFromBOM(j) = strGetExtensionExcludeFromBOM(i) Then
                intTotalQuantityExcludeFromBOM(i) = intTotalQuantityExcludeFromBOM(i) + 1
                For k = j To intCountAllComponentExcludeFromBOM - 2
                    strCheckConfigurationNameExcludeFromBOM(k) = strCheckConfigurationNameExcludeFromBOM(k + 1)
                    intTotalQuantityExcludeFromBOM(k) = intTotalQuantityExcludeFromBOM(k + 1)
                    strGetPartnameExcludeFromBOM(k) = strGetPartnameExcludeFromBOM(k + 1)
                    strGetExtensionExcludeFromBOM(k) = strGetExtensionExcludeFromBOM(k + 1)
                Next k
                intCountAllComponentExcludeFromBOM = intCountAllComponentExcludeFromBOM - 1
            End If
        If strGetPartnameExcludeFromBOM(j) = strGetPartnameExcludeFromBOM(i) And strCheckConfigurationNameExcludeFromBOM(j) = strCheckConfigurationNameExcludeFromBOM(i) And strGetExtensionExcludeFromBOM(j) = strGetExtensionExcludeFromBOM(i) Then GoTo CheckDuplicatePartnameQTYExcludeFromBOM
        Next j
    Next i
    'Subtract components exclude from BOM into all components
    For i = 0 To intCountAllComponent - 1
        For j = 0 To intCountAllComponentExcludeFromBOM - 1
            If strCheckPartname(i) = strGetPartnameExcludeFromBOM(j) And strCheckConfigurationName(i) = strCheckConfigurationNameExcludeFromBOM(j) And strCheckExtension(i) = strGetExtensionExcludeFromBOM(j) Then
                intTotalQuantity(i) = intTotalQuantity(i) - intTotalQuantityExcludeFromBOM(j)
            End If
        Next j
    Next i
    'Remove all components have quantity equal zero
    For i = 0 To intCountAllComponent - 1
        If intTotalQuantity(i) <= 0 Then
CheckQTYzero:
            EndProgram.EndByESCKey
            For j = i To intCountAllComponent - 2
                strCheckPartname(j) = strCheckPartname(j + 1)
                strCheckConfigurationName(j) = strCheckConfigurationName(j + 1)
                strCheckExtension(j) = strCheckExtension(j + 1)
                intTotalQuantity(j) = intTotalQuantity(j + 1)
                strGetPathname(j) = strGetPathname(j + 1)
                strGetParentCompSubFeat(j) = strGetParentCompSubFeat(j + 1)
                strGetQTY(j) = strGetQTY(j + 1)
            Next j
            intCountAllComponent = intCountAllComponent - 1
        End If
        If intCountAllComponent = 1 Then Exit For
        If intTotalQuantity(i) <= 0 And i <= intCountAllComponent - 1 Then GoTo CheckQTYzero
    Next i
    'Collect standard devices and remove duplicate data
    Set swOpenModel = swModel
    Call CollectStandardDevice
    'Compare exist quantity values with new one.
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    xlApp.WindowState = xlMaximized
    Set xlWorkbook = xlApp.Workbooks.Add
    With xlWorkbook.Worksheets(1)
        On Error Resume Next
        .Columns("B").NumberFormat = "@"
        On Error Resume Next
        .Columns("C").NumberFormat = "@"
        On Error Resume Next
        .Columns("D").NumberFormat = "@"
        On Error Resume Next
        .Columns("E").NumberFormat = "@"
        On Error Resume Next
        .Columns("F").NumberFormat = "@"
        On Error Resume Next
        .Cells(1, 1).Value = "SO S" & ChrW(193) & "NH S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG " & ChrW(272) & ChrW(195) & " L" & ChrW(431) & "U VA" & ChrW(768) & " S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
        On Error Resume Next
        .Range(.Cells(1, 1), .Cells(1, 6)).Merge
        FilterPartNameFromPathName swModel.GetPathName
        On Error Resume Next
        .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
        On Error Resume Next
        .Range(.Cells(2, 1), .Cells(2, 6)).Merge
        On Error Resume Next
        .Cells(3, 1).Value = "STT"
        On Error Resume Next
        .Cells(3, 2).Value = "T" & ChrW(202) & "N CHI TI" & ChrW(7870) & "T"
        On Error Resume Next
        .Cells(3, 3).Value = "CHI TI" & ChrW(7870) & "T TRONG CU" & ChrW(803) & "M"
        On Error Resume Next
        .Cells(3, 4).Value = "T" & ChrW(202) & "N C" & ChrW(7844) & "U H" & ChrW(204) & "NH"
        On Error Resume Next
        .Cells(3, 5).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG " & ChrW(272) & ChrW(195) & " L" & ChrW(431) & "U"
        On Error Resume Next
        .Cells(3, 6).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
        On Error Resume Next
        .Range("A1:F3").Font.Bold = True
        On Error Resume Next
        .Range("A1:F3").Font.Size = 10
        m = 0
        For i = 0 To intCountAllComponent - 1
            j = 0
            If blnExcludeStandardDevice = True Then
                For k = 0 To intCountStandardDevice - 1
                    If strStandardDevice(k) = strCheckPartname(i) Or strStandardDevice(k) = strGetParentCompSubFeat(i) Then j = j + 1
                Next k
                For k = 0 To intCountChildrenStandardDevice - 1
                    If strGetSubFeatFolder(k) = strCheckPartname(i) Then j = j + 1
                Next k
            End If
            n = 0
            For l = 0 To intCountSemiStandardDevice - 1
                If strCheckPartname(i) = strSemiStandardDevice(l) Or strSemiStandardDevice(l) = strGetParentCompSubFeat(i) Then n = 1
            Next l
            If j = 0 Or n = 1 Then
                On Error Resume Next
                .Cells(m + 4, 1).Value = m + 1
                On Error Resume Next
                .Hyperlinks.Add .Cells(m + 4, 2), strGetPathname(i), TextToDisplay:=strCheckPartname(i)
                On Error Resume Next
                .Cells(m + 4, 3).Value = strGetParentCompSubFeat(i)
                On Error Resume Next
                .Cells(m + 4, 4).Value = strCheckConfigurationName(i)
                On Error Resume Next
                .Cells(m + 4, 5).Value = strGetQTY(i)
                On Error Resume Next
                .Cells(m + 4, 6).Value = CStr(intTotalQuantity(i) * strTotalSetAutoQTY)
                If .Cells(m + 4, 5).Value <> .Cells(m + 4, 6).Value Then .Range(.Cells(m + 4, 1), .Cells(m + 4, 6)).Interior.Color = RGB(255, 0, 0)
                m = m + 1
            End If
        Next i
        On Error Resume Next
        .Columns("A:F").AutoFit
        On Error Resume Next
        .Columns("A:F").HorizontalAlignment = xlCenter
        On Error Resume Next
        .Range(.Cells(1, 1), .Cells(m + 3, 6)).Borders.LineStyle = xlContinuous
    End With
    End
End Sub

Private Sub CollectStandardDevice()
    'Collect all standard device
    ReDim strStandardDevice(UBound(vComps)) As String
    Set swFeat = swOpenModel.FirstFeature
    j = 0
    Call GetNameStandardDevice
    'Collect all semi-standard device
    ReDim strSemiStandardDevice(UBound(vComps)) As String
    Set swFeat = swOpenModel.FirstFeature
    m = 0
    Call GetNameSemiStandardDevice
    'Collect all mechanical standard device
    ReDim strMechanicalStandardDevice(UBound(vComps)) As String
    Set swFeat = swOpenModel.FirstFeature
    n = 0
    Call GetNameMechanicalStandardDevice
    'Collect all electrical standard device
    ReDim strElectricalStandardDevice(UBound(vComps)) As String
    Set swFeat = swOpenModel.FirstFeature
    a = 0
    Call GetNameElectricalStandardDevice
    'Collect all hydraulics and pneumatics standard device
    ReDim strHydraulics_PneumaticsStandardDevice(UBound(vComps)) As String
    Set swFeat = swOpenModel.FirstFeature
    b = 0
    Call GetNameHydraulics_PneumaticsStandardDevice
    'Collect all bolt standard device
    ReDim strBoltStandardDevice(UBound(vComps)) As String
    Set swFeat = swOpenModel.FirstFeature
    c = 0
    Call GetNameBoltStandardDevice
    For k = LBound(vComps) To UBound(vComps)
        Set swComp = vComps(k)
        Set swFeat = swComp.FirstFeature
        Call GetNameStandardDevice
        Set swFeat = swComp.FirstFeature
        Call GetNameSemiStandardDevice
        Set swFeat = swComp.FirstFeature
        Call GetNameMechanicalStandardDevice
        Set swFeat = swComp.FirstFeature
        Call GetNameElectricalStandardDevice
        Set swFeat = swComp.FirstFeature
        Call GetNameHydraulics_PneumaticsStandardDevice
        Set swFeat = swComp.FirstFeature
        Call GetNameBoltStandardDevice
    Next
    'Remove duplicate data of standard device
    For i = 0 To j - 2
CheckDuplicateStandardDevice:
        For k = i + 1 To j - 1
            If strStandardDevice(k) = strStandardDevice(i) Or strStandardDevice(k) = "" Then
                For l = k To j - 2
                    strStandardDevice(l) = strStandardDevice(l + 1)
                Next l
                j = j - 1
            End If
            If strStandardDevice(k) = strStandardDevice(i) Or strStandardDevice(k) = "" Then GoTo CheckDuplicateStandardDevice
        Next k
    Next i
    intCountStandardDevice = j
    'Remove duplicate data of semi-standard device
    For i = 0 To m - 2
CheckDuplicateSemiStandardDevice:
        For k = i + 1 To m - 1
            If strSemiStandardDevice(k) = strSemiStandardDevice(i) Or strSemiStandardDevice(k) = "" Then
                For l = k To m - 2
                    strSemiStandardDevice(l) = strSemiStandardDevice(l + 1)
                Next l
                m = m - 1
            End If
            If strSemiStandardDevice(k) = strSemiStandardDevice(i) Or strSemiStandardDevice(k) = "" Then GoTo CheckDuplicateSemiStandardDevice
        Next k
    Next i
    intCountSemiStandardDevice = m
    'Remove duplicate data of mechanical standard device
    For i = 0 To n - 2
CheckDuplicateMechanicalStandardDevice:
        For k = i + 1 To n - 1
            If strMechanicalStandardDevice(k) = strMechanicalStandardDevice(i) Or strMechanicalStandardDevice(k) = "" Then
                For l = k To n - 2
                    strMechanicalStandardDevice(l) = strMechanicalStandardDevice(l + 1)
                Next l
                n = n - 1
            End If
            If strMechanicalStandardDevice(k) = strMechanicalStandardDevice(i) Or strMechanicalStandardDevice(k) = "" Then GoTo CheckDuplicateMechanicalStandardDevice
        Next k
    Next i
    intCountMechanicalStandardDevice = n
    'Remove duplicate data of electrical standard device
    For i = 0 To a - 2
CheckDuplicateElectricalStandardDevice:
        For k = i + 1 To a - 1
            If strElectricalStandardDevice(k) = strElectricalStandardDevice(i) Or strElectricalStandardDevice(k) = "" Then
                For l = k To a - 2
                    strElectricalStandardDevice(l) = strElectricalStandardDevice(l + 1)
                Next l
                a = a - 1
            End If
            If strElectricalStandardDevice(k) = strElectricalStandardDevice(i) Or strElectricalStandardDevice(k) = "" Then GoTo CheckDuplicateElectricalStandardDevice
        Next k
    Next i
    intCountElectricalStandardDevice = a
    'Remove duplicate data of Hydraulics & Pneumatics standard device
    For i = 0 To b - 2
CheckDuplicateHydraulics_PneumaticsStandardDevice:
        For k = i + 1 To b - 1
            If strHydraulics_PneumaticsStandardDevice(k) = strHydraulics_PneumaticsStandardDevice(i) Or strHydraulics_PneumaticsStandardDevice(k) = "" Then
                For l = k To b - 2
                    strHydraulics_PneumaticsStandardDevice(l) = strHydraulics_PneumaticsStandardDevice(l + 1)
                Next l
                b = b - 1
            End If
            If strHydraulics_PneumaticsStandardDevice(k) = strHydraulics_PneumaticsStandardDevice(i) Or strHydraulics_PneumaticsStandardDevice(k) = "" Then GoTo CheckDuplicateHydraulics_PneumaticsStandardDevice
        Next k
    Next i
    intCountHydraulics_PneumaticsStandardDevice = b
    'Remove duplicate data of bolt standard device
    For i = 0 To c - 2
CheckDuplicateBoltStandardDevice:
        For k = i + 1 To c - 1
            If strBoltStandardDevice(k) = strBoltStandardDevice(i) Or strBoltStandardDevice(k) = "" Then
                For l = k To c - 2
                    strBoltStandardDevice(l) = strBoltStandardDevice(l + 1)
                Next l
                c = c - 1
            End If
            If strBoltStandardDevice(k) = strBoltStandardDevice(i) Or strBoltStandardDevice(k) = "" Then GoTo CheckDuplicateBoltStandardDevice
        Next k
    Next i
    intCountBoltStandardDevice = c
    'Get Children Standard Device inside sub feature in Folder
    ReDim strGetSubFeatFolder(UBound(vComps)) As String
    l = 0
    For i = 0 To intCountAllComponent - 1
        For j = 0 To intCountStandardDevice - 1
            If strStandardDevice(j) <> "" And strStandardDevice(j) = strGetParentCompSubFeat(i) Then
                For k = 0 To intCountAllComponent - 1
                    If strGetParentCompSubFeat(k) = strCheckPartname(i) Then
                        strGetSubFeatFolder(l) = strCheckPartname(k) 'Sub Feat level 2
                        l = l + 1
                        If l = intCountAllComponent Then GoTo ExitSubFeature
                        For m = 0 To intCountAllComponent - 1
                            If strGetParentCompSubFeat(m) = strCheckPartname(k) Then
                                strGetSubFeatFolder(l) = strCheckPartname(m) 'Sub Feat level 3
                                l = l + 1
                                If l = intCountAllComponent Then GoTo ExitSubFeature
                                For n = 0 To intCountAllComponent - 1
                                    If strGetParentCompSubFeat(n) = strCheckPartname(m) Then
                                        strGetSubFeatFolder(l) = strCheckPartname(n) 'Sub Feat level 4
                                        l = l + 1
                                        If l = intCountAllComponent Then GoTo ExitSubFeature
                                        For a = 0 To intCountAllComponent - 1
                                            If strGetParentCompSubFeat(a) = strCheckPartname(n) Then
                                                If l < intCountAllComponent Then strGetSubFeatFolder(l) = strCheckPartname(a) 'Sub Feat level 5
                                                l = l + 1
                                                If l = intCountAllComponent Then GoTo ExitSubFeature
                                            End If
                                        Next a
                                    End If
                                Next n
                            End If
                        Next m
                    End If
                Next k
            End If
        Next j
    Next i
ExitSubFeature:
    intCountChildrenStandardDevice = l
End Sub

Private Sub GetNameStandardDevice()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "FtrFolder" And InStr(1, swFeat.name, "EndTag", vbTextCompare) = 0 Then
            Set swFeatFolder = swFeat.GetSpecificFeature2
            vFeatures = swFeatFolder.GetFeatures
            For i = 0 To (swFeatFolder.GetFeatureCount - 1)
                Set swFtrFolder = vFeatures(i)
                FilterPartNameTopLevelLast (swFtrFolder.name)
                If blnCheckLastHorizontalMark = True Then strStandardDevice(j) = strPartNameTopLevel_Last
                j = j + 1
            Next i
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

Private Sub GetNameSemiStandardDevice()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "FtrFolder" And (StrComp(UCase(Trim(swFeat.name)), "TBGC+TBTC", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "TBTC+TBGC", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "BTC", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "TBBTC", vbTextCompare) = 0) Then
            Set swFeatFolder = swFeat.GetSpecificFeature2
            vFeatures = swFeatFolder.GetFeatures
            For i = 0 To (swFeatFolder.GetFeatureCount - 1)
                Set swFtrFolder = vFeatures(i)
                FilterPartNameTopLevelLast (swFtrFolder.name)
                If blnCheckLastHorizontalMark = True Then strSemiStandardDevice(m) = strPartNameTopLevel_Last
                m = m + 1
            Next i
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

Private Sub GetNameMechanicalStandardDevice()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "FtrFolder" And (StrComp(UCase(Trim(swFeat.name)), "TBC", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "TBTCC", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "TBTC TKC", vbTextCompare) = 0) Then
            Set swFeatFolder = swFeat.GetSpecificFeature2
            vFeatures = swFeatFolder.GetFeatures
            For i = 0 To (swFeatFolder.GetFeatureCount - 1)
                Set swFtrFolder = vFeatures(i)
                FilterPartNameTopLevelLast (swFtrFolder.name)
                If blnCheckLastHorizontalMark = True Then strMechanicalStandardDevice(n) = strPartNameTopLevel_Last
                n = n + 1
            Next i
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

Private Sub GetNameElectricalStandardDevice()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "FtrFolder" And (StrComp(UCase(Trim(swFeat.name)), "TBD", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "TBTC TKD", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "TBTCD", vbTextCompare) = 0) Then
            Set swFeatFolder = swFeat.GetSpecificFeature2
            vFeatures = swFeatFolder.GetFeatures
            For i = 0 To (swFeatFolder.GetFeatureCount - 1)
                Set swFtrFolder = vFeatures(i)
                FilterPartNameTopLevelLast (swFtrFolder.name)
                If blnCheckLastHorizontalMark = True Then strElectricalStandardDevice(a) = strPartNameTopLevel_Last
                a = a + 1
            Next i
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

Private Sub GetNameBoltStandardDevice()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "FtrFolder" And (StrComp(UCase(Trim(swFeat.name)), "BULONG", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "BOLT", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "BOLT NUT", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "BOLT - NUT", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "BOLT-NUT", vbTextCompare) = 0) Then
            Set swFeatFolder = swFeat.GetSpecificFeature2
            vFeatures = swFeatFolder.GetFeatures
            For i = 0 To (swFeatFolder.GetFeatureCount - 1)
                Set swFtrFolder = vFeatures(i)
                FilterPartNameTopLevelLast (swFtrFolder.name)
                If blnCheckLastHorizontalMark = True Then strBoltStandardDevice(c) = strPartNameTopLevel_Last
                c = c + 1
            Next i
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

Private Sub GetNameHydraulics_PneumaticsStandardDevice()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "FtrFolder" And (StrComp(UCase(Trim(swFeat.name)), "TLKN", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "TBTC TLKN", vbTextCompare) = 0 Or StrComp(UCase(Trim(swFeat.name)), "TBTLKN", vbTextCompare) = 0) Then
            Set swFeatFolder = swFeat.GetSpecificFeature2
            vFeatures = swFeatFolder.GetFeatures
            For i = 0 To (swFeatFolder.GetFeatureCount - 1)
                Set swFtrFolder = vFeatures(i)
                FilterPartNameTopLevelLast (swFtrFolder.name)
                If blnCheckLastHorizontalMark = True Then strHydraulics_PneumaticsStandardDevice(b) = strPartNameTopLevel_Last
                b = b + 1
            Next i
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

Private Sub blnNewAbsoluteQTY_Click()
    If blnNewAbsoluteQTY.Value = False Then
        opbtOnlyQTY3D.Value = True
        opbtOnlyExcelQTY.Enabled = False
        opbtOnlyExcelQTY.Value = False
        opbtElectricalStandardDevice.Enabled = False
        opbtElectricalStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = False
        opbtMechanicalStandardDevice.Value = False
        opbtAllStandardDevice.Enabled = False
        opbtAllStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = False
        opbtBoltStandardDevice.Value = False
    Else
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = True
        opbtOnlyExcelQTY.Value = True
        opbtElectricalStandardDevice.Enabled = True
        opbtElectricalStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = True
        opbtMechanicalStandardDevice.Value = False
        opbtAllStandardDevice.Enabled = True
        opbtAllStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = True
        opbtBoltStandardDevice.Value = False
    End If
End Sub

Private Sub opbtAllStandardDevice_Click()
    If opbtAllStandardDevice = True Then
        opbtOnlyQTY3D.Enabled = False
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = False
        opbtOnlyExcelQTY.Value = False
        FrameExportQTYManufacturingComponents.Enabled = False
        blnExcludeStandardDevice.Enabled = False
        blnExcludeStandardDevice.Value = False
        blnNewAbsoluteQTY.Enabled = False
        blnNewAbsoluteQTY.Value = True
        opbtElectricalStandardDevice.Enabled = False
        opbtElectricalStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = False
        opbtMechanicalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = False
        opbtBoltStandardDevice.Value = False
    Else
        opbtOnlyQTY3D.Enabled = True
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = True
        opbtOnlyExcelQTY.Value = True
        FrameExportQTYManufacturingComponents.Enabled = True
        blnExcludeStandardDevice.Enabled = True
        blnExcludeStandardDevice.Value = True
        blnNewAbsoluteQTY.Enabled = True
        blnNewAbsoluteQTY.Value = True
        opbtElectricalStandardDevice.Enabled = True
        opbtElectricalStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = True
        opbtMechanicalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = True
        opbtBoltStandardDevice.Value = False
    End If
End Sub

Private Sub opbtMechanicalStandardDevice_Click()
    If opbtMechanicalStandardDevice = True Then
        opbtOnlyQTY3D.Enabled = False
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = False
        opbtOnlyExcelQTY.Value = False
        FrameExportQTYManufacturingComponents.Enabled = False
        blnExcludeStandardDevice.Enabled = False
        blnExcludeStandardDevice.Value = False
        blnNewAbsoluteQTY.Enabled = False
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = False
        opbtAllStandardDevice.Value = False
        opbtElectricalStandardDevice.Enabled = False
        opbtElectricalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = False
        opbtBoltStandardDevice.Value = False
    Else
        opbtOnlyQTY3D.Enabled = True
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = True
        opbtOnlyExcelQTY.Value = True
        FrameExportQTYManufacturingComponents.Enabled = True
        blnExcludeStandardDevice.Enabled = True
        blnExcludeStandardDevice.Value = True
        blnNewAbsoluteQTY.Enabled = True
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = True
        opbtAllStandardDevice.Value = False
        opbtElectricalStandardDevice.Enabled = True
        opbtElectricalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = True
        opbtBoltStandardDevice.Value = False
    End If
End Sub

Private Sub opbtElectricalStandardDevice_Click()
    If opbtElectricalStandardDevice = True Then
        opbtOnlyQTY3D.Enabled = False
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = False
        opbtOnlyExcelQTY.Value = False
        FrameExportQTYManufacturingComponents.Enabled = False
        blnExcludeStandardDevice.Enabled = False
        blnExcludeStandardDevice.Value = False
        blnNewAbsoluteQTY.Enabled = False
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = False
        opbtAllStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = False
        opbtMechanicalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = False
        opbtBoltStandardDevice.Value = False
    Else
        opbtOnlyQTY3D.Enabled = True
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = True
        opbtOnlyExcelQTY.Value = True
        FrameExportQTYManufacturingComponents.Enabled = True
        blnExcludeStandardDevice.Enabled = True
        blnExcludeStandardDevice.Value = True
        blnNewAbsoluteQTY.Enabled = True
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = True
        opbtAllStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = True
        opbtMechanicalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = True
        opbtBoltStandardDevice.Value = False
    End If
End Sub

Private Sub opbtHydraulics_PneumaticsStandardDevice_Click()
    If opbtHydraulics_PneumaticsStandardDevice = True Then
        opbtOnlyQTY3D.Enabled = False
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = False
        opbtOnlyExcelQTY.Value = False
        FrameExportQTYManufacturingComponents.Enabled = False
        blnExcludeStandardDevice.Enabled = False
        blnExcludeStandardDevice.Value = False
        blnNewAbsoluteQTY.Enabled = False
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = False
        opbtAllStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = False
        opbtMechanicalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = False
        opbtElectricalStandardDevice.Enabled = False
        opbtElectricalStandardDevice.Value = False
    Else
        opbtOnlyQTY3D.Enabled = True
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = True
        opbtOnlyExcelQTY.Value = True
        FrameExportQTYManufacturingComponents.Enabled = True
        blnExcludeStandardDevice.Enabled = True
        blnExcludeStandardDevice.Value = True
        blnNewAbsoluteQTY.Enabled = True
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = True
        opbtAllStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = True
        opbtMechanicalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = True
        opbtElectricalStandardDevice.Enabled = True
        opbtElectricalStandardDevice.Value = False
    End If
End Sub

Private Sub opbtBoltStandardDevice_Click()
    On Error Resume Next
    If opbtBoltStandardDevice = True Then
        chkBulongCum.Enabled = True
    Else
        chkBulongCum.Value = False
        chkBulongCum.Enabled = False
    End If
    On Error GoTo 0
    If opbtBoltStandardDevice = True Then
        opbtOnlyQTY3D.Enabled = False
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = False
        opbtOnlyExcelQTY.Value = False
        FrameExportQTYManufacturingComponents.Enabled = False
        blnExcludeStandardDevice.Enabled = False
        blnExcludeStandardDevice.Value = False
        blnNewAbsoluteQTY.Enabled = False
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = False
        opbtAllStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = False
        opbtMechanicalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = False
        opbtElectricalStandardDevice.Enabled = False
        opbtElectricalStandardDevice.Value = False
        opbtHydraulics_PneumaticsStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Value = False
    Else
        opbtOnlyQTY3D.Enabled = True
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = True
        opbtOnlyExcelQTY.Value = True
        FrameExportQTYManufacturingComponents.Enabled = True
        blnExcludeStandardDevice.Enabled = True
        blnExcludeStandardDevice.Value = True
        blnNewAbsoluteQTY.Enabled = True
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = True
        opbtAllStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = True
        opbtMechanicalStandardDevice.Value = False
        opbtSemiStandardDevice.Enabled = True
        opbtElectricalStandardDevice.Enabled = True
        opbtElectricalStandardDevice.Value = False
        opbtHydraulics_PneumaticsStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Value = False
    End If
End Sub

Private Sub opbtSemiStandardDevice_Click()
    If opbtSemiStandardDevice = True Then
        opbtOnlyQTY3D.Enabled = False
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = False
        opbtOnlyExcelQTY.Value = False
        FrameExportQTYStandardDevice.Enabled = False
        blnExcludeStandardDevice.Enabled = False
        blnExcludeStandardDevice.Value = False
        blnNewAbsoluteQTY.Enabled = False
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = False
        opbtAllStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = False
        opbtMechanicalStandardDevice.Value = False
        opbtElectricalStandardDevice.Enabled = False
        opbtElectricalStandardDevice.Value = False
        opbtHydraulics_PneumaticsStandardDevice.Enabled = False
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = False
        opbtBoltStandardDevice.Value = False
    Else
        opbtOnlyQTY3D.Enabled = True
        opbtOnlyQTY3D.Value = False
        opbtOnlyExcelQTY.Enabled = True
        opbtOnlyExcelQTY.Value = True
        FrameExportQTYStandardDevice.Enabled = True
        blnExcludeStandardDevice.Enabled = True
        blnExcludeStandardDevice.Value = True
        blnNewAbsoluteQTY.Enabled = True
        blnNewAbsoluteQTY.Value = True
        opbtAllStandardDevice.Enabled = True
        opbtAllStandardDevice.Value = False
        opbtMechanicalStandardDevice.Enabled = True
        opbtMechanicalStandardDevice.Value = False
        opbtElectricalStandardDevice.Enabled = True
        opbtElectricalStandardDevice.Value = False
        opbtHydraulics_PneumaticsStandardDevice.Enabled = True
        opbtHydraulics_PneumaticsStandardDevice.Value = False
        opbtBoltStandardDevice.Enabled = True
        opbtBoltStandardDevice.Value = False
    End If
End Sub

Private Sub cmdAUTOQUANTITY_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    Set EndProgram = New Class_ClickESC
    Call ErrorHandling2
    While InStr(1, strTotalSetAutoQTY, ".", vbTextCompare) <> 0 Or InStr(1, strTotalSetAutoQTY, ",", vbTextCompare) <> 0
        swApp.SendMsgToUser2 "Please input number again!", swMbWarning, swMbOk
        Exit Sub
    Wend
    While strTotalSetAutoQTY <= 0 Or strTotalSetAutoQTY = "" Or IsNumeric(strTotalSetAutoQTY) = False
        swApp.SendMsgToUser2 "Please input number again!", swMbWarning, swMbOk
        Exit Sub
    Wend

    '========================================================================
    ' NEU CHON XUAT BULONG: CHAY QUY TRINH DEM DA CAP DOC LAP VA THOAT
    ' (Bao toan 100% tinh nang goc va khong gay anh huong den cac muc khac)

    '========================================================================
    Dim blnRunBulong As Boolean
    blnRunBulong = False
    On Error Resume Next
    blnRunBulong = (opbtBoltStandardDevice.Value = True)
    On Error GoTo 0
    If blnRunBulong Then
        Call ProcessExportBulong
        Unload AUTOPARTLIST
        Exit Sub
    End If
    If (opbtOnlyQTY3D = True And blnNewAbsoluteQTY = True) Or (opbtBothAutoQTY = True And blnNewAbsoluteQTY = True) Then
        If swApp.SendMsgToUser2("This program will create new quantities replace for exist values. Do you want to process? ", swMbWarning, swMbYesNo) = swMbHitNo Then Exit Sub
    End If
    Call TurnOffAutoSaved
    swApp.FrameState = swWindowMaximized
    Set swAsm = swModel
    vComps = swAsm.GetComponents(False)
    If opbtOnlyQTY3D = True Then swAsm.ResolveAllLightWeightComponents False
    ReDim strCheckPartname(UBound(vComps)) As String
    Dim intTotalQuantity() As Integer
    ReDim intTotalQuantity(UBound(vComps)) As Integer
    Dim strCheckConfigurationName() As String
    ReDim strCheckConfigurationName(UBound(vComps)) As String
    ReDim strCheckExtension(UBound(vComps)) As String
    Dim strGetPathname() As String
    ReDim strGetPathname(UBound(vComps)) As String
    ReDim strGetParentCompSubFeat(UBound(vComps)) As String
    Dim swParentComp As Component2
    'Collect all components
    intCountAllComponent = 0
    For i = LBound(vComps) To UBound(vComps)
        EndProgram.EndByESCKey
        Set swComp = vComps(i)
        On Error Resume Next
        If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.ExcludeFromBOM = False And swComp.GetSuppression <> 0 Then
            strExtension = Right(UCase(swComp.GetPathName), 6)
            strCheckExtension(intCountAllComponent) = strExtension
            If strExtension = "SLDASM" Then swComp.SetSuppression2 swComponentResolved
            Call FilterPartNameFromPathName(swComp.GetPathName)
            strCheckPartname(intCountAllComponent) = strPartName
            If Trim(swComp.ReferencedConfiguration) = strPartName Then
                strCheckConfigurationName(intCountAllComponent) = "Default"
            Else
                strCheckConfigurationName(intCountAllComponent) = Trim(swComp.ReferencedConfiguration)
            End If
            strGetPathname(intCountAllComponent) = swComp.GetPathName
            Set swParentComp = swComp.GetParent
            If Not swParentComp Is Nothing Then
                Call FilterPartNameFromPathName(swParentComp.GetPathName)
                strGetParentCompSubFeat(intCountAllComponent) = strPartName
            Else
                Call FilterPartNameFromPathName(swModel.GetPathName)
                strGetParentCompSubFeat(intCountAllComponent) = strPartName
            End If
            intTotalQuantity(intCountAllComponent) = 1
            intCountAllComponent = intCountAllComponent + 1
        End If
    Next
    'Remove duplicate data and calculate QTY for all components
    For i = 0 To intCountAllComponent - 2
CheckDuplicatePartnameQTY:
        EndProgram.EndByESCKey
        For j = i + 1 To intCountAllComponent - 1
            If strCheckPartname(j) = strCheckPartname(i) And strCheckConfigurationName(j) = strCheckConfigurationName(i) And strCheckExtension(j) = strCheckExtension(i) Then
                intTotalQuantity(i) = intTotalQuantity(i) + 1
                For k = j To intCountAllComponent - 2
                    strCheckPartname(k) = strCheckPartname(k + 1)
                    strCheckConfigurationName(k) = strCheckConfigurationName(k + 1)
                    strCheckExtension(k) = strCheckExtension(k + 1)
                    intTotalQuantity(k) = intTotalQuantity(k + 1)
                    strGetPathname(k) = strGetPathname(k + 1)
                    strGetParentCompSubFeat(k) = strGetParentCompSubFeat(k + 1)
                Next k
                intCountAllComponent = intCountAllComponent - 1
            End If
        If strCheckPartname(j) = strCheckPartname(i) And strCheckConfigurationName(j) = strCheckConfigurationName(i) And strCheckExtension(j) = strCheckExtension(i) Then GoTo CheckDuplicatePartnameQTY
        Next j
    Next i
    'Collect component exclude from BOM
    Dim strGetPartnameExcludeFromBOM() As String
    ReDim strGetPartnameExcludeFromBOM(UBound(vComps)) As String
    Dim strGetExtensionExcludeFromBOM() As String
    ReDim strGetExtensionExcludeFromBOM(UBound(vComps)) As String
    Dim intTotalQuantityExcludeFromBOM() As Integer
    ReDim intTotalQuantityExcludeFromBOM(UBound(vComps)) As Integer
    Dim strCheckConfigurationNameExcludeFromBOM() As String
    ReDim strCheckConfigurationNameExcludeFromBOM(UBound(vComps)) As String
    Dim intCountAllComponentExcludeFromBOM As Integer
    intCountAllComponentExcludeFromBOM = 0
    Dim swCompModel As ModelDoc2
    For i = LBound(vComps) To UBound(vComps)
        Set swComp = vComps(i)
        Set swCompModel = swComp.GetModelDoc2
        strExtension = Right(UCase(swComp.GetPathName), 6)
        On Error Resume Next
        If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And strExtension = "SLDASM" And swComp.ExcludeFromBOM Then
            Dim vCompExcludeFromBOM As Variant
            Set swAsm = swCompModel
            vCompExcludeFromBOM = swAsm.GetComponents(False)
            For j = LBound(vCompExcludeFromBOM) To UBound(vCompExcludeFromBOM)
                Set swComp = vCompExcludeFromBOM(j)
                Set swParentComp = swComp.GetParent
                Call FilterPartNameTopLevelTemp(swComp.Name2)
                If blnCheckFirstSlashes = True Then
                    On Error Resume Next
                    If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.ExcludeFromBOM = False And swParentComp.ExcludeFromBOM = False Then
                        strCheckConfigurationNameExcludeFromBOM(intCountAllComponentExcludeFromBOM) = swComp.ReferencedConfiguration
                        Call FilterPartNameFromPathName(swComp.GetPathName)
                        strGetPartnameExcludeFromBOM(intCountAllComponentExcludeFromBOM) = strPartName
                        strExtension = Right(UCase(swComp.GetPathName), 6)
                        strGetExtensionExcludeFromBOM(intCountAllComponentExcludeFromBOM) = strExtension
                        intTotalQuantityExcludeFromBOM(intCountAllComponentExcludeFromBOM) = 1
                        intCountAllComponentExcludeFromBOM = intCountAllComponentExcludeFromBOM + 1
                    End If
                Else
                    On Error Resume Next
                    If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.ExcludeFromBOM = False Then
                        strCheckConfigurationNameExcludeFromBOM(intCountAllComponentExcludeFromBOM) = swComp.ReferencedConfiguration
                        Call FilterPartNameFromPathName(swComp.GetPathName)
                        strGetPartnameExcludeFromBOM(intCountAllComponentExcludeFromBOM) = strPartName
                        strExtension = Right(UCase(swComp.GetPathName), 6)
                        strGetExtensionExcludeFromBOM(intCountAllComponentExcludeFromBOM) = strExtension
                        intTotalQuantityExcludeFromBOM(intCountAllComponentExcludeFromBOM) = 1
                        intCountAllComponentExcludeFromBOM = intCountAllComponentExcludeFromBOM + 1
                    End If
                End If
            Next
        End If
    Next
    'Remove duplicate data and calculate QTY for components exclude from BOM
    For i = 0 To intCountAllComponentExcludeFromBOM - 2
CheckDuplicatePartnameQTYExcludeFromBOM:
        EndProgram.EndByESCKey
        For j = i + 1 To intCountAllComponentExcludeFromBOM - 1
            If strGetPartnameExcludeFromBOM(j) = strGetPartnameExcludeFromBOM(i) And strCheckConfigurationNameExcludeFromBOM(j) = strCheckConfigurationNameExcludeFromBOM(i) And strGetExtensionExcludeFromBOM(j) = strGetExtensionExcludeFromBOM(i) Then
                intTotalQuantityExcludeFromBOM(i) = intTotalQuantityExcludeFromBOM(i) + 1
                For k = j To intCountAllComponentExcludeFromBOM - 2
                    strCheckConfigurationNameExcludeFromBOM(k) = strCheckConfigurationNameExcludeFromBOM(k + 1)
                    intTotalQuantityExcludeFromBOM(k) = intTotalQuantityExcludeFromBOM(k + 1)
                    strGetPartnameExcludeFromBOM(k) = strGetPartnameExcludeFromBOM(k + 1)
                    strGetExtensionExcludeFromBOM(k) = strGetExtensionExcludeFromBOM(k + 1)
                Next k
                intCountAllComponentExcludeFromBOM = intCountAllComponentExcludeFromBOM - 1
            End If
        If strGetPartnameExcludeFromBOM(j) = strGetPartnameExcludeFromBOM(i) And strCheckConfigurationNameExcludeFromBOM(j) = strCheckConfigurationNameExcludeFromBOM(i) And strGetExtensionExcludeFromBOM(j) = strGetExtensionExcludeFromBOM(i) Then GoTo CheckDuplicatePartnameQTYExcludeFromBOM
        Next j
    Next i
    'Subtract components exclude from BOM into all components
    For i = 0 To intCountAllComponent - 1
        For j = 0 To intCountAllComponentExcludeFromBOM - 1
            If strCheckPartname(i) = strGetPartnameExcludeFromBOM(j) And strCheckConfigurationName(i) = strCheckConfigurationNameExcludeFromBOM(j) And strCheckExtension(i) = strGetExtensionExcludeFromBOM(j) Then
                intTotalQuantity(i) = intTotalQuantity(i) - intTotalQuantityExcludeFromBOM(j)
            End If
        Next j
    Next i
    'Remove all components have quantity equal zero
    For i = 0 To intCountAllComponent - 1
        If intTotalQuantity(i) <= 0 Then
CheckQTYzero:
            EndProgram.EndByESCKey
            For j = i To intCountAllComponent - 2
                strCheckPartname(j) = strCheckPartname(j + 1)
                strCheckConfigurationName(j) = strCheckConfigurationName(j + 1)
                strCheckExtension(j) = strCheckExtension(j + 1)
                intTotalQuantity(j) = intTotalQuantity(j + 1)
                strGetPathname(j) = strGetPathname(j + 1)
                strGetParentCompSubFeat(j) = strGetParentCompSubFeat(j + 1)
            Next j
            intCountAllComponent = intCountAllComponent - 1
        End If
        If intCountAllComponent = 1 Then Exit For
        If intTotalQuantity(i) <= 0 And i <= intCountAllComponent - 1 Then GoTo CheckQTYzero
    Next i
    'Collect standard devices and remove duplicate data
    Set swOpenModel = swModel
    Call CollectStandardDevice
    'Export QTY data into Configuration Specific of 3D model
    If opbtOnlyQTY3D = True Then
LoopClickMouseAutoQty_1:
        EndProgram.EndByESCKey
        Dim blnSetAutoQty As Boolean
        blnSetAutoQty = False
        swOpenModel.ClearSelection2 (True)
        Call SetClickMouse
        Call SetConfigModel
        Dim iMemory As Integer
        Dim strQty As String
        flag = False
        If blnNewAbsoluteQTY = False Then
            swModelConfigMgr.Get3 "QTY", True, Empty, strQty
            If strQty <> "0" And strQty <> "" Then
                iMemory = CInt(strTotalSetAutoQTY) - CInt(strQty)
                If iMemory < 0 Then
                    swApp.SendMsgToUser2 "Please check again input quantity = " & strTotalSetAutoQTY & vbNewLine & "This parameter must be larger than old exist quantity!", swMbWarning, swMbOk
                    End
                Else
                    If swApp.SendMsgToUser2("This program will create new quantities replace for exist values. Do you want to process? ", swMbWarning, swMbYesNo) = swMbHitNo Then Exit Sub
                End If
            Else
                Call FilterPartNameFromPathName(swModel.GetPathName)
                swApp.SendMsgToUser2 "Please check again quantity of assembly: " & strPartName, swMbWarning, swMbOk
                swApp.ActivateTaskPane (5)
                End
            End If
        End If
        If Not swModelConfigMgr Is Nothing Then
            For i = 0 To swModelConfigMgr.Count
                On Error Resume Next
                If UCase(vModelConfigMgr(i)) = "QTY" Then
                    swModelConfigMgr.Set2 "QTY", CInt(strTotalSetAutoQTY)
                    swModelConfigMgr.Get3 "QTY", False, Empty, strQty
                End If
            Next i
        End If
        If strQty <> "" And strQty <> "0" Then blnSetAutoQty = True
        If blnSetAutoQty = False Then
            If Xposition <= DisplayWidth Then
                TheMouse.Move DisplayWidth * 2 - 260, 160, swMouse_MouseMove
                GoTo LoopClickMouseAutoQty_1
            Else
                TheMouse.Move DisplayWidth - 260, 160, swMouse_MouseMove
                GoTo LoopClickMouseAutoQty_1
            End If
        End If
        strQty = CStr(strQty)
        strTotalSetAutoQTY = CStr(strTotalSetAutoQTY)
        swModel.Save3 swSaveAsOptions_Silent, Empty, Empty
        For i = 0 To intCountAllComponent - 1
            EndProgram.EndByESCKey
            j = 0
            If blnExcludeStandardDevice = True Then
                For k = 0 To intCountStandardDevice - 1
                    If strStandardDevice(k) = strCheckPartname(i) Or strStandardDevice(k) = strGetParentCompSubFeat(i) Then j = j + 1
                Next k
                For k = 0 To intCountChildrenStandardDevice - 1
                    If strGetSubFeatFolder(k) = strCheckPartname(i) Then j = j + 1
                Next k
            End If
            n = 0
            For l = 0 To intCountSemiStandardDevice - 1
                If strCheckPartname(i) = strSemiStandardDevice(l) Then n = n + 1
            Next l
            If j = 0 Or n > 0 Then
                Set swOpenModel = swApp.ActivateDoc3(strGetPathname(i), False, 2, 0)
                swOpenModel.ShowConfiguration2 strCheckConfigurationName(i)
LoopClickMouseAutoQty_2:
                Call SetClickMouse
                Call SetConfigModel
                blnSetAutoQty = False
                If blnNewAbsoluteQTY = True Then
                    If Not swModelConfigMgr Is Nothing Then
                        For k = 0 To swModelConfigMgr.Count
                            On Error Resume Next
                            If UCase(vModelConfigMgr(k)) = "QTY" Then
                                swModelConfigMgr.Set2 "QTY", CStr(intTotalQuantity(i) * CInt(strTotalSetAutoQTY))
                                swModelConfigMgr.Get3 "QTY", True, Empty, strQty
                            End If
                        Next k
                    End If
                    If strQty <> "" And strQty <> "0" Then blnSetAutoQty = True
                Else
                    If Not swModelConfigMgr Is Nothing Then
                        For k = 0 To swModelConfigMgr.Count
                            On Error Resume Next
                            If UCase(vModelConfigMgr(k)) = "QTY" Then
                                swModelConfigMgr.Get3 "QTY", True, Empty, strQty
                                If iMemory > 0 Then swModelConfigMgr.Set2 "QTY", CStr(CInt(strQty) + iMemory * intTotalQuantity(i))
                                swModelConfigMgr.Get3 "QTY", True, Empty, strQty
                            End If
                        Next k
                    End If
                    If strQty <> "" And strQty <> "0" Then blnSetAutoQty = True
                End If
                If blnSetAutoQty = False Then
                    EndProgram.EndByESCKey
                    If Xposition <= DisplayWidth Then
                        TheMouse.Move DisplayWidth * 2 - 260, 160, swMouse_MouseMove
                        GoTo LoopClickMouseAutoQty_2
                    Else
                        TheMouse.Move DisplayWidth - 260, 160, swMouse_MouseMove
                        GoTo LoopClickMouseAutoQty_2
                    End If
                End If
                strQty = CStr(strQty)
                strTotalSetAutoQTY = CStr(strTotalSetAutoQTY)
                swOpenModel.Save3 swSaveAsOptions_Silent, Empty, Empty
                swApp.CloseDoc strGetPathname(i)
            End If
        Next i
    End If
    'Export QTY for all manufacturing device into EXCEL
    If opbtOnlyExcelQTY = True Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add
        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"
            On Error Resume Next
            .Columns("D").NumberFormat = "@"
            On Error Resume Next
            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH THI" & ChrW(7870) & "T B" & ChrW(7882) & " GIA C" & ChrW(212) & "NG"
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(1, 5)).Merge
            Call FilterPartNameFromPathName(swModel.GetPathName)
            On Error Resume Next
            .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            On Error Resume Next
            .Range(.Cells(2, 1), .Cells(2, 5)).Merge
            On Error Resume Next
            .Cells(3, 1).Value = "STT"
            On Error Resume Next
            .Cells(3, 2).Value = "T" & ChrW(202) & "N CHI TI" & ChrW(7870) & "T"
            On Error Resume Next
            .Cells(3, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            On Error Resume Next
            .Cells(3, 4).Value = "T" & ChrW(202) & "N C" & ChrW(7844) & "U H" & ChrW(204) & "NH"
            On Error Resume Next
            .Cells(3, 5).Value = "GHI CH" & ChrW(218)
            On Error Resume Next
            .Range("A1:E3").Font.Bold = True
            m = 0
            On Error Resume Next
            .Cells(m + 4, 1).Value = 1
            On Error Resume Next
            .Cells(m + 4, 2).Value = strPartName
            On Error Resume Next
            .Cells(m + 4, 3).Value = strTotalSetAutoQTY
            On Error Resume Next
            .Cells(m + 4, 4).Value = "Default"
            On Error Resume Next
            .Cells(m + 4, 5).Value = "C" & ChrW(7908) & "M"
            For i = 0 To intCountAllComponent - 1
                j = 0
                If blnExcludeStandardDevice = True Then
                    For k = 0 To intCountStandardDevice - 1
                        If strStandardDevice(k) = strCheckPartname(i) Or strStandardDevice(k) = strGetParentCompSubFeat(i) Then j = j + 1
                    Next k
                    For k = 0 To intCountChildrenStandardDevice - 1
                        If strGetSubFeatFolder(k) = strCheckPartname(i) Then j = j + 1
                    Next k
                End If
                n = 0
                For l = 0 To intCountSemiStandardDevice - 1
                    If strCheckPartname(i) = strSemiStandardDevice(l) Or strSemiStandardDevice(l) = strGetParentCompSubFeat(i) Then n = 1
                Next l
                If j = 0 Or n = 1 Then
                    On Error Resume Next
                    .Cells(m + 5, 1).Value = m + 2
                    On Error Resume Next
                    .Hyperlinks.Add .Cells(m + 5, 2), strGetPathname(i), TextToDisplay:=strCheckPartname(i)
                    On Error Resume Next
                    .Cells(m + 5, 3).Value = CStr(intTotalQuantity(i) * CInt(strTotalSetAutoQTY))
                    On Error Resume Next
                    .Cells(m + 5, 4).Value = strCheckConfigurationName(i)
                    On Error Resume Next
                    If strCheckExtension(i) = "SLDASM" Then .Cells(m + 4, 5).Value = "C" & ChrW(7908) & "M"
                    m = m + 1
                End If
            Next i
            On Error Resume Next
            .Columns("A:E").AutoFit
            On Error Resume Next
            .Columns("A:E").HorizontalAlignment = xlCenter
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(m + 4, 5)).Borders.LineStyle = xlContinuous
        End With
        Unload AUTOPARTLIST
    End If
    If opbtOnlyQTY3D = True And flag = False Then
        swApp.SendMsgToUser2 "AUTO QUANTITY COMPLETED!", swMbInformation, swMbOk
        End
    End If
    'Export QTY for all standard device
    If opbtAllStandardDevice = True Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add
        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"
            On Error Resume Next
            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH T" & ChrW(7892) & "NG H" & ChrW(7906) & "P THI" & ChrW(7870) & "T B" & ChrW(7882) & " TI" & ChrW(202) & "U CHU" & ChrW(7848) & "N"
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(1, 3)).Merge
            Call FilterPartNameFromPathName(swModel.GetPathName)
            On Error Resume Next
            .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            On Error Resume Next
            .Range(.Cells(2, 1), .Cells(2, 3)).Merge
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(3, 3)).Font.Size = 10
            On Error Resume Next
            .Cells(3, 1).Value = "STT"
            On Error Resume Next
            .Cells(3, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
            On Error Resume Next
            .Cells(3, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            On Error Resume Next
            .Range("A1:C3").Font.Bold = True
            j = 0
            For i = 0 To intCountAllComponent - 1
                For k = 0 To intCountStandardDevice - 1
                    If strStandardDevice(k) = strCheckPartname(i) And strStandardDevice(k) <> strGetParentCompSubFeat(i) Then
                        On Error Resume Next
                        .Cells(j + 4, 1).Value = j + 1
                        If InStr(1, strCheckConfigurationName(i), "Default", vbTextCompare) = 0 Then
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strCheckConfigurationName(i)
                        Else
                            strPartNameTopLevel_Last = strStandardDevice(k)
                            Call EditPartNameStandardDevice
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strPartNameTopLevel_Last
                        End If
                        On Error Resume Next
                        .Cells(j + 4, 3).Value = CStr(intTotalQuantity(i) * CInt(strTotalSetAutoQTY))
                        j = j + 1
                    End If
                Next k
            Next i
            On Error Resume Next
            .Columns("A:C").AutoFit
            On Error Resume Next
            .Columns("A:C").HorizontalAlignment = xlCenter
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(j + 3, 3)).Borders.LineStyle = xlContinuous
        End With
        Unload AUTOPARTLIST
    End If
    'Export QTY for all semi-standard device
    If opbtSemiStandardDevice = True Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add
        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"
            On Error Resume Next
            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH THI" & ChrW(7870) & "T B" & ChrW(7882) & " B" & ChrW(193) & "N TI" & ChrW(202) & "U CHU" & ChrW(7848) & "N"
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(1, 3)).Merge
            Call FilterPartNameFromPathName(swModel.GetPathName)
            On Error Resume Next
            .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            On Error Resume Next
            .Range(.Cells(2, 1), .Cells(2, 3)).Merge
            On Error Resume Next
            .Cells(3, 1).Value = "STT"
            On Error Resume Next
            .Cells(3, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
            On Error Resume Next
            .Cells(3, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            On Error Resume Next
            .Range("A1:C3").Font.Bold = True
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(3, 3)).Font.Size = 10
            j = 0
            For i = 0 To intCountAllComponent - 1
                For k = 0 To intCountSemiStandardDevice - 1
                    If strSemiStandardDevice(k) = strCheckPartname(i) And strSemiStandardDevice(k) <> strGetParentCompSubFeat(i) Then
                        On Error Resume Next
                        .Cells(j + 4, 1).Value = j + 1
                        If InStr(1, strCheckConfigurationName(i), "Default", vbTextCompare) = 0 Then
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strCheckConfigurationName(i)
                        Else
                            strPartNameTopLevel_Last = strSemiStandardDevice(k)
                            Call EditPartNameStandardDevice
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strPartNameTopLevel_Last
                        End If
                        On Error Resume Next
                        .Cells(j + 4, 3).Value = CStr(intTotalQuantity(i) * CInt(strTotalSetAutoQTY))
                        j = j + 1
                    End If
                Next k
            Next i
            On Error Resume Next
            .Columns("A:C").AutoFit
            On Error Resume Next
            .Columns("A:C").HorizontalAlignment = xlCenter
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(j + 3, 3)).Borders.LineStyle = xlContinuous
        End With
        Unload AUTOPARTLIST
    End If
    'Export QTY for all mechanical standard device
    If opbtMechanicalStandardDevice = True Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add
        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"
            On Error Resume Next
            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH THI" & ChrW(7870) & "T B" & ChrW(7882) & " TI" & ChrW(202) & "U CHU" & ChrW(7848) & "N C" & ChrW(416)
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(1, 3)).Merge
            Call FilterPartNameFromPathName(swModel.GetPathName)
            On Error Resume Next
            .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            On Error Resume Next
            .Range(.Cells(2, 1), .Cells(2, 3)).Merge
            On Error Resume Next
           .Range(.Cells(1, 1), .Cells(3, 3)).Font.Size = 10
            On Error Resume Next
            .Cells(3, 1).Value = "STT"
            On Error Resume Next
            .Cells(3, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
            On Error Resume Next
            .Cells(3, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            On Error Resume Next
            .Range("A1:C3").Font.Bold = True
            j = 0
            For i = 0 To intCountAllComponent - 1
                n = 0
                For k = 0 To intCountSemiStandardDevice - 1
                    If strCheckPartname(i) = strSemiStandardDevice(k) And strSemiStandardDevice(k) <> strGetParentCompSubFeat(i) Then n = 1
                Next k
                For k = 0 To intCountMechanicalStandardDevice - 1
                    If strMechanicalStandardDevice(k) = strCheckPartname(i) And strMechanicalStandardDevice(k) <> strGetParentCompSubFeat(i) Then n = 1
                Next k
                If n = 1 Then
                    On Error Resume Next
                    .Cells(j + 4, 1).Value = j + 1
                    If InStr(1, strCheckConfigurationName(i), "Default", vbTextCompare) = 0 Then
                        On Error Resume Next
                        .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strCheckConfigurationName(i)
                    Else
                        strPartNameTopLevel_Last = strCheckPartname(i)
                        Call EditPartNameStandardDevice
                        On Error Resume Next
                        .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strPartNameTopLevel_Last
                    End If
                    On Error Resume Next
                    .Cells(j + 4, 3).Value = CStr(intTotalQuantity(i) * CInt(strTotalSetAutoQTY))
                    j = j + 1
                End If
            Next i
            On Error Resume Next
            .Columns("A:C").AutoFit
            On Error Resume Next
            .Columns("A:C").HorizontalAlignment = xlCenter
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(j + 3, 3)).Borders.LineStyle = xlContinuous
        End With
        Unload AUTOPARTLIST
    End If
    'Export QTY for all electrical standard device
    If opbtElectricalStandardDevice = True Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add
        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"
            On Error Resume Next
            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH THI" & ChrW(7870) & "T B" & ChrW(7882) & " TI" & ChrW(202) & "U CHU" & ChrW(7848) & "N " & ChrW(208) & "I" & ChrW(7878) & "N"
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(1, 3)).Merge
            Call FilterPartNameFromPathName(swModel.GetPathName)
            On Error Resume Next
            .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            On Error Resume Next
            .Range(.Cells(2, 1), .Cells(2, 3)).Merge
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(3, 3)).Font.Size = 10
            On Error Resume Next
            .Cells(3, 1).Value = "STT"
            On Error Resume Next
            .Cells(3, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
            On Error Resume Next
            .Cells(3, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            On Error Resume Next
            .Range("A1:C3").Font.Bold = True
            j = 0
            For i = 0 To intCountAllComponent - 1
                For k = 0 To intCountElectricalStandardDevice - 1
                    If strElectricalStandardDevice(k) = strCheckPartname(i) And strElectricalStandardDevice(k) <> strGetParentCompSubFeat(i) Then
                        On Error Resume Next
                        .Cells(j + 4, 1).Value = j + 1
                        If InStr(1, strCheckConfigurationName(i), "Default", vbTextCompare) = 0 Then
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strCheckConfigurationName(i)
                        Else
                            strPartNameTopLevel_Last = strElectricalStandardDevice(k)
                            Call EditPartNameStandardDevice
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strPartNameTopLevel_Last
                        End If
                        On Error Resume Next
                        .Cells(j + 4, 3).Value = CStr(intTotalQuantity(i) * CInt(strTotalSetAutoQTY))
                        j = j + 1
                    End If
                Next k
            Next i
            On Error Resume Next
            .Columns("A:C").AutoFit
            On Error Resume Next
            .Columns("A:C").HorizontalAlignment = xlCenter
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(j + 3, 3)).Borders.LineStyle = xlContinuous
        End With
        Unload AUTOPARTLIST
    End If
    'Export QTY for all Hydraulics & Pneumatics standard device
    If opbtHydraulics_PneumaticsStandardDevice = True Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add
        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"
            On Error Resume Next
            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH THI" & ChrW(7870) & "T B" & ChrW(7882) & " TI" & ChrW(202) & "U CHU" & ChrW(7848) & "N " & "TH" & ChrW(7910) & "Y L" & ChrW(7920) & "C KHI" & ChrW(769) & " NE" & ChrW(769) & "N"
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(1, 3)).Merge
            Call FilterPartNameFromPathName(swModel.GetPathName)
            On Error Resume Next
            .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            On Error Resume Next
            .Range(.Cells(2, 1), .Cells(2, 3)).Merge
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(3, 3)).Font.Size = 10
            On Error Resume Next
            .Cells(3, 1).Value = "STT"
            On Error Resume Next
            .Cells(3, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
            On Error Resume Next
            .Cells(3, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            On Error Resume Next
            .Range("A1:C3").Font.Bold = True
            j = 0
            For i = 0 To intCountAllComponent - 1
                For k = 0 To intCountHydraulics_PneumaticsStandardDevice - 1
                    If strHydraulics_PneumaticsStandardDevice(k) = strCheckPartname(i) And strHydraulics_PneumaticsStandardDevice(k) <> strGetParentCompSubFeat(i) Then
                        On Error Resume Next
                        .Cells(j + 4, 1).Value = j + 1
                        If InStr(1, strCheckConfigurationName(i), "Default", vbTextCompare) = 0 Then
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strCheckConfigurationName(i)
                        Else
                            strPartNameTopLevel_Last = strHydraulics_PneumaticsStandardDevice(k)
                            Call EditPartNameStandardDevice
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strPartNameTopLevel_Last
                        End If
                        On Error Resume Next
                        .Cells(j + 4, 3).Value = CStr(intTotalQuantity(i) * CInt(strTotalSetAutoQTY))
                        j = j + 1
                    End If
                Next k
            Next i
            On Error Resume Next
            .Columns("A:C").AutoFit
            On Error Resume Next
            .Columns("A:C").HorizontalAlignment = xlCenter
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(j + 3, 3)).Borders.LineStyle = xlContinuous
        End With
        Unload AUTOPARTLIST
    End If
    'Export QTY for all bolt standard device
    If opbtBoltStandardDevice = True Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add
        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"
            On Error Resume Next
            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH BULONG"
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(1, 3)).Merge
            Call FilterPartNameFromPathName(swModel.GetPathName)
            On Error Resume Next
            .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            On Error Resume Next
            .Range(.Cells(2, 1), .Cells(2, 3)).Merge
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(3, 3)).Font.Size = 10
            On Error Resume Next
            .Cells(3, 1).Value = "STT"
            On Error Resume Next
            .Cells(3, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
            On Error Resume Next
            .Cells(3, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            On Error Resume Next
            .Range("A1:C3").Font.Bold = True
            j = 0
            For i = 0 To intCountAllComponent - 1
                For k = 0 To intCountBoltStandardDevice - 1
                    If strBoltStandardDevice(k) = strCheckPartname(i) And strBoltStandardDevice(k) <> strGetParentCompSubFeat(i) Then
                        On Error Resume Next
                        .Cells(j + 4, 1).Value = j + 1
                        If InStr(1, strCheckConfigurationName(i), "Default", vbTextCompare) = 0 Then
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strCheckConfigurationName(i)
                        Else
                            strPartNameTopLevel_Last = strBoltStandardDevice(k)
                            Call EditPartNameStandardDevice
                            On Error Resume Next
                            .Hyperlinks.Add .Cells(j + 4, 2), strGetPathname(i), TextToDisplay:=strPartNameTopLevel_Last
                        End If
                        On Error Resume Next
                        .Cells(j + 4, 3).Value = CStr(intTotalQuantity(i) * CInt(strTotalSetAutoQTY))
                        j = j + 1
                    End If
                Next k
            Next i
            On Error Resume Next
            .Columns("A:C").AutoFit
            On Error Resume Next
            .Columns("A:C").HorizontalAlignment = xlCenter
            On Error Resume Next
            .Range(.Cells(1, 1), .Cells(j + 3, 3)).Borders.LineStyle = xlContinuous
        End With
        Unload AUTOPARTLIST
    End If
End Sub

Private Sub cmdAutoRemoveInsertThread_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    'Remove insert thread of dimension
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    Do While Not swView Is Nothing
        Set swDispDim = swView.GetFirstDisplayDimension5
        Do While Not swDispDim Is Nothing
            strGetCalloutAboveDefinition = swDispDim.GetText(swDimensionTextCalloutAboveDefinition)
            strGetCalloutBelow = swDispDim.GetText(swDimensionTextCalloutBelow)
            strGetPrefixTextDefinition = swDispDim.GetText(swDimensionTextPrefixDefinition)
            strGetSuffixTextDefinition = swDispDim.GetText(swDimensionTextSuffixDefinition)
            If swDispDim.IsHoleCallout() Then
                Call GetAddNoteInsertThread(strGetCalloutAboveDefinition)
                Call GetAddNoteInsertThread(strGetCalloutBelow)
                Call GetAddNoteInsertThread(strGetPrefixTextDefinition)
                Call GetAddNoteInsertThread(strGetSuffixTextDefinition)
                Call RemoveInsertThread(strGetCalloutAboveDefinition)
                Call RemoveInsertThread(strGetCalloutBelow)
                Call RemoveInsertThread(strGetPrefixTextDefinition)
                Call RemoveInsertThread(strGetSuffixTextDefinition)
                If strGetAddNoteInsertThread <> "" Then swDispDim.SetText swDimensionTextCalloutBelow, strGetAddNoteInsertThread
            Else
                Set swDisplayData = swDispDim.GetDisplayData()
                For i = 1 To 5
                    If swDisplayData.GetTextAtIndex(i) <> "" And InStr(1, swDisplayData.GetTextAtIndex(i), "<HOLE-DEPTH>", vbTextCompare) = 0 And InStr(1, UCase(swDisplayData.GetTextAtIndex(i)), "INSERT REN", vbTextCompare) = 0 And IsNumeric(swDisplayData.GetTextAtIndex(i)) = False Then
                        strGetAddNoteInsertThread = swDisplayData.GetTextAtIndex(i)
                    End If
                Next i
                Call RemoveInsertThread(strGetCalloutAboveDefinition)
                Call RemoveInsertThread(strGetCalloutBelow)
                Call RemoveInsertThread(strGetPrefixTextDefinition)
                Call RemoveInsertThread(strGetSuffixTextDefinition)
                If strGetAddNoteInsertThread <> "" Then swDispDim.SetText swDimensionTextCalloutBelow, strGetAddNoteInsertThread
            End If
            Set swDispDim = swDispDim.GetNext3
        Loop
        Set swView = swView.GetNextView
    Loop
    'Remove insert thread of Hole Table
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotationType_e.swTableAnnotation_HoleChart Then
            Set swAnn = swTable.GetAnnotation
            swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationHidden
            For m = 0 To swTable.ColumnCount - 1
                If InStr(1, UCase(swTable.Text(0, m)), "NOTE", vbTextCompare) <> 0 Or InStr(1, UCase(swTable.Text(0, m)), "GHI CH" & ChrW(769), vbTextCompare) <> 0 Then
                    For i = 1 To swTable.RowCount - 1
                        If swTable.Text(i, m) <> "" And InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare) Then
                            If InStr(1, swTable.Text(i, m), "M", vbTextCompare) Then
                                If InStr(InStr(1, swTable.Text(i, m), "M", vbTextCompare), swTable.Text(i, m), "x", vbTextCompare) Then
                                    If InStr(InStr(1, swTable.Text(i, m), "x", vbTextCompare), swTable.Text(i, m), "D", vbTextCompare) Then
                                        strGetAddNoteInsertThread = Right(swTable.Text(i, m), Len(swTable.Text(i, m)) - InStr(InStr(1, swTable.Text(i, m), "x", vbTextCompare), swTable.Text(i, m), "D", vbTextCompare))
                                    Else
                                        If InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) Then
                                            strGetAddNoteInsertThread = Right(swTable.Text(i, m), Len(swTable.Text(i, m)) - InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) - 1)
                                        End If
                                    End If
                                Else
                                    If InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) Then
                                        strGetAddNoteInsertThread = Right(swTable.Text(i, m), Len(swTable.Text(i, m)) - InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) - 1)
                                    End If
                                End If
                            Else
                                If InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) Then
                                    strGetAddNoteInsertThread = Right(swTable.Text(i, m), Len(swTable.Text(i, m)) - InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) - 1)
                                End If
                            End If
                        ElseIf swTable.Text(i, m) <> "" And InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare) = 0 Then
                            strGetAddNoteInsertThread = swTable.Text(i, m)
                        End If
                        swTable.Text(i, m) = strGetAddNoteInsertThread
                        strGetAddNoteInsertThread = ""
                    Next i
                End If
            Next m
            swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationVisible
        End If
        Set swTable = swTable.GetNext
    Wend
    swModel.EditRebuild3
    swApp.SendMsgToUser2 "REMOVE INSERT THREAD COMPLETED!", swMbInformation, swMbOk
    Unload AUTOPARTLIST
End Sub

Private Sub cmdUpdateTemplate_Click()
Call CheckRun
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Set EndProgram = New Class_ClickESC
    AUTOPARTLIST.Hide
    flag = False
    Dim intUpdateTemplate As Integer
    intUpdateTemplate = swApp.SendMsgToUser2("Do you want to update template automatically? Please push Yes button to accept, or push No button to update template manually!", swMbQuestion, swMbYesNoCancel)
    If intUpdateTemplate = swMbHitNo Then
        swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swUseFolderSearchRules, 1
        strFilter = "Template File(*.slddrt)|*.slddrt|"
        If swApp.SendMsgToUser2("Please choose a file sheet format of Solidworks!", swMbInformation, swMbYesNo) = swMbHitYes Then
            strFileName = swApp.GetOpenFileName("Open file", Empty, strFilter, Empty, Empty, Empty)
            If strFileName = "" Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        flag = True
    ElseIf intUpdateTemplate = swMbHitCancel Then
        End
    End If
    vSheetNames = swDraw.GetSheetNames()
    For intSheetNumber = LBound(vSheetNames) To UBound(vSheetNames)
        EndProgram.EndByESCKey
        swDraw.ActivateSheet vSheetNames(intSheetNumber)
        Set swView = swDraw.GetFirstView
        Set swTable = swView.GetFirstTableAnnotation
        blnHoleTable = True
        While Not swTable Is Nothing
            If swTable.Type = swTableAnnotation_BillOfMaterials Then blnHoleTable = False
            Set swTable = swTable.GetNext
        Wend
        i = 0
        Do While Not swView Is Nothing
            Dim swBaseView As View
            Set swBaseView = swView.GetBaseView
            If Not swBaseView Is Nothing Then
                vGetScaleRatio = swBaseView.ScaleRatio
                Call ProcessUpdateTemplate(swBaseView)  'Update template for drawing have many views
                Exit Do
            End If
            i = i + 1
            Set swView = swView.GetNextView
        Loop
        If i = 2 Then
            Set swView = swDraw.GetFirstView
            Set swView = swView.GetNextView
            vGetScaleRatio = swView.ScaleRatio
            Call ProcessUpdateTemplate(swView) 'Update template for drawing have one view
        End If
        If intSheetNumber = UBound(vSheetNames) Then Call ResetSketchVisibility
    Next
    'Set name of Drawing Custom Properties
    Set swModel = swApp.NewDocument(swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates) & "\A3-TCI.drwdot", swDwgPaperA3size, Empty, Empty)
    Dim strDraw As String
    Dim strDesign As String
    Dim strCheck As String
    Dim strApprove As String
    Set swDrawCustPropMgr = swModel.Extension.CustomPropertyManager(Empty)
    ReDim strValue(swDrawCustPropMgr.Count) As String
    vDrawCustPropNames = swDrawCustPropMgr.GetNames()
    For i = 0 To UBound(vDrawCustPropNames)
        swDrawCustPropMgr.Get3 vDrawCustPropNames(i), False, Empty, strValue(i)
        Select Case UCase(vDrawCustPropNames(i))
        Case "DRAWN BY"
            strDraw = strValue(i)
        Case "DESIGN BY"
            strDesign = strValue(i)
        Case "CHECKED BY"
            strCheck = strValue(i)
        Case "APPROVED BY"
            strApprove = strValue(i)
        End Select
    Next i
    swApp.CloseDoc swModel.GetTitle
    Set swModel = swApp.ActiveDoc
    Set swDrawCustPropMgr = swModel.Extension.CustomPropertyManager(Empty)
    ReDim strValue(swDrawCustPropMgr.Count) As String
    vDrawCustPropNames = swDrawCustPropMgr.GetNames()
    For i = 0 To UBound(vDrawCustPropNames)
        swDrawCustPropMgr.Get3 vDrawCustPropNames(i), False, Empty, strValue(i)
        Select Case UCase(vDrawCustPropNames(i))
        Case "DRAWN BY"
            If Len(strValue(i)) <= 1 Then swDrawCustPropMgr.Set "DRAWN BY", strDraw
        Case "DESIGN BY"
            If Len(strValue(i)) <= 1 Then swDrawCustPropMgr.Set "DESIGN BY", strDesign
        Case "CHECKED BY"
            If Len(strValue(i)) <= 1 Then swDrawCustPropMgr.Set "CHECKED BY", strCheck
        Case "APPROVED BY"
            If Len(strValue(i)) <= 1 Then swDrawCustPropMgr.Set "APPROVED BY", strApprove
        End Select
    Next i
    swModel.Extension.LoadDraftingStandard swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsDraftingStandard) & "\ISO-TCI.sldstd"
    swModel.Extension.SetUserPreferenceToggle swUserPreferenceIntegerValue_e.swDetailingDimensionStandard, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swDetailingStandard_e.swDetailingStandardUserDefined
    swApp.SendMsgToUser2 "UPDATE SHEET FORMATS COMPLETED!", swMbInformation, swMbOk
    Unload AUTOPARTLIST
End Sub

Private Sub ProcessUpdateTemplate(swView As View)
    Set swSheet = swDraw.GetCurrentSheet()
    vGetSheetProperties = swSheet.GetProperties2()
    swModel.EditTemplate
    Call SelectRoughnessBox
    Set swSelMgr = swModel.SelectionManager
    Dim blnRoughnessBox As Boolean
    blnRoughnessBox = False
    If swSelMgr.GetSelectedObjectCount2(-1) <> 0 Then blnRoughnessBox = True
    swModel.ClearSelection2 True
    swModel.Extension.SketchBoxSelect "-0.01", "0.85", "0.00", "1.2", "-0.01", "0.00" 'Select all entities of template
    swModel.Extension.DeleteSelection2 (0) 'Delete all entities of template
    If vGetSheetProperties(0) = 8 Then 'Form A3 =swDwgPaperA3size
        Dim strSheetFormatA3 As String
        On Error Resume Next
        strSheetFormatA3 = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsSheetFormat) & "\A3 TCI.slddrt"
        swSheet.SetProperties2 swDwgPaperA3size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
        On Error Resume Next
        If flag = False Then
            swSheet.SetTemplateName strSheetFormatA3
        Else
            swSheet.SetTemplateName strFileName
        End If
    ElseIf vGetSheetProperties(0) = 9 Then 'Form A2=swDwgPaperA2size
        Dim strSheetFormatA2 As String
        On Error Resume Next
        strSheetFormatA2 = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsSheetFormat) & "\A2-TCI QR_up.slddrt"
        swSheet.SetProperties2 swDwgPaperA2size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
        On Error Resume Next
        If flag = False Then
            swSheet.SetTemplateName strSheetFormatA2
        Else
            swSheet.SetTemplateName strFileName
        End If
    ElseIf vGetSheetProperties(0) = 10 Then 'Form A1=swDwgPaperA1size
        Dim strSheetFormatA1 As String
        On Error Resume Next
        strSheetFormatA1 = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsSheetFormat) & "\A1-TCI QR_up.slddrt"
        swSheet.SetProperties2 swDwgPaperA1size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
        On Error Resume Next
        If flag = False Then
            swSheet.SetTemplateName strSheetFormatA1
        Else
            swSheet.SetTemplateName strFileName
        End If
    ElseIf vGetSheetProperties(0) = 11 Then 'Form A0=swDwgPaperA0size
        Dim strSheetFormatA0 As String
        On Error Resume Next
        strSheetFormatA0 = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swFileLocationsSheetFormat) & "\A0-TCI.slddrt"
        swSheet.SetProperties2 swDwgPaperA0size, swDwgTemplateCustom, vGetScaleRatio(0), vGetScaleRatio(1), True, Empty, Empty, False
        On Error Resume Next
        If flag = False Then
            swSheet.SetTemplateName strSheetFormatA0
        Else
            swSheet.SetTemplateName strFileName
        End If
    End If
    swSheet.ReloadTemplate True
    If blnRoughnessBox = False Then
        Call SelectRoughnessBox
        swModel.Extension.DeleteSelection2 (0)
    End If
    If blnHoleTable = False Then
        'Call SelectHoleTableBox
        'swModel.Extension.DeleteSelection2 (0)
    End If
    swModel.EditSheet
    swView.UseSheetScale() = True
End Sub

Private Sub SelectRoughnessBox()
    If vGetSheetProperties(0) = 11 Then 'Form A0=swDwgPaperA0size
        swModel.Extension.SketchBoxSelect "1.1", "0.84", "0.00", "1.2", "0.038", "0.00" 'Select Roughness Box
    ElseIf vGetSheetProperties(0) = 10 Then 'Form A1 = swDwgPaperA1size
        swModel.Extension.SketchBoxSelect "0.78", "0.6", "0.00", "0.84", "0.038", "0.00" 'Select Roughness Box
    ElseIf vGetSheetProperties(0) = 9 Then 'Form A2=swDwgPaperA2size
        swModel.Extension.SketchBoxSelect "0.52", "0.42", "0.00", "0.6", "0.038", "0.00" 'Select Roughness Box
    ElseIf vGetSheetProperties(0) = 8 Then 'Form A3 = swDwgPaperA3size
        swModel.Extension.SketchBoxSelect "0.35", "0.3", "0.00", "0.425", "0.038", "0.00" 'Select Roughness Box
    End If
End Sub

Private Sub SelectHoleTableBox()
    If vGetSheetProperties(0) = 11 Then 'Form A0=swDwgPaperA0size
        swModel.Extension.SketchBoxSelect "0.0", "0.04", "0.00", "0.87", "0.0", "0.0" 'Select Hole Table Box
    ElseIf vGetSheetProperties(0) = 10 Then 'Form A1 = swDwgPaperA1size
        swModel.Extension.SketchBoxSelect "0.0", "0.03", "0.00", "0.53", "0.0", "0.0" 'Select Hole Table Box
    ElseIf vGetSheetProperties(0) = 9 Then  'Form A2=swDwgPaperA2size
        swModel.Extension.SketchBoxSelect "0.0", "0.04", "0.00", "0.435", "0.0", "0.0" 'Select Hole Table Box
    ElseIf vGetSheetProperties(0) = 8 Then 'Form A3=swDwgPaperA3size
        swModel.Extension.SketchBoxSelect "0.0", "0.03", "0.00", "0.265", "0.0", "0.0" 'Select Hole Table Box
    End If
End Sub

Private Sub cmdDeleteRoughnessbox_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    Call ErrorHandling
    swModel.EditTemplate
    Set swSheet = swDraw.GetCurrentSheet()
    vGetSheetProperties = swSheet.GetProperties2()
    Call SelectRoughnessBox
    Set swSelMgr = swModel.SelectionManager
    If swSelMgr.GetSelectedObjectCount2(-1) <> 0 Then
        Call SelectRoughnessBox
        swModel.Extension.DeleteSelection2 (0)
    Else
        swModel.ClearSelection2 True
        Call SelectHoleTableBox
        Dim blnHoleTableBox As Boolean
        blnHoleTableBox = False
        If swSelMgr.GetSelectedObjectCount2(-1) <> 0 Then blnHoleTableBox = True
        swModel.Extension.SketchBoxSelect "-0.01", "0.85", "0.00", "1.2", "-0.01", "0.00" 'Select all entities of template
        swModel.Extension.DeleteSelection2 (0) 'Delete all entities of template
        swSheet.ReloadTemplate True
        If blnHoleTableBox = False Then
            'Call SelectHoleTableBox
            'swModel.Extension.DeleteSelection2 (0) 'Delete Hole Table
        End If
        Call ResetSketchVisibility
    End If
    swModel.EditSheet
    swModel.EditRebuild3
    AUTOPARTLIST.Show
End Sub

Private Sub ResetSketchVisibility()
    Set swView = swDraw.GetFirstView
    Do While Not swView Is Nothing
        swView.ResetSketchVisibility
        Set swView = swView.GetNextView
    Loop
End Sub

Private Sub RemoveInsertThread(strCheckInsertThread As String)
    intInsertThreadPosition = InStr(1, UCase(strCheckInsertThread), "INSERT REN", vbTextCompare)
    If intInsertThreadPosition <> 0 Then
        Select Case strCheckInsertThread
        Case strGetCalloutAboveDefinition
            swDispDim.SetText swDimensionTextCalloutAbove, ""
        Case strGetCalloutBelow
            swDispDim.SetText swDimensionTextCalloutBelow, ""
        Case strGetPrefixTextDefinition
            swDispDim.SetText swDimensionTextPrefix, ""
        Case strGetSuffixTextDefinition
            swDispDim.SetText swDimensionTextSuffix, ""
        End Select
    End If
    If strGetAddNoteInsertThread <> "" And InStr(1, strCheckInsertThread, strGetAddNoteInsertThread, vbTextCompare) Then
        Select Case strCheckInsertThread
        Case strGetCalloutAboveDefinition
            swDispDim.SetText swDimensionTextCalloutAbove, ""
        Case strGetCalloutBelow
            swDispDim.SetText swDimensionTextCalloutBelow, ""
        Case strGetPrefixTextDefinition
            swDispDim.SetText swDimensionTextPrefix, ""
        Case strGetSuffixTextDefinition
            swDispDim.SetText swDimensionTextSuffix, ""
        End Select
    End If
End Sub

Private Sub GetAddNoteInsertThread(strCheckInsertThread As String)
    If strCheckInsertThread <> "" And InStr(1, UCase(strCheckInsertThread), "INSERT REN", vbTextCompare) = 0 And InStr(1, strCheckInsertThread, "<hw-threaddesc>", vbTextCompare) = 0 Then strGetAddNoteInsertThread = strCheckInsertThread
End Sub

Private Sub CheckInsertThreadMaterial()
    blnInsertThreadMaterial = False
    Dim vInsertThreadMaterial As Variant
    vInsertThreadMaterial = Array("AL", "A2024", "A5052", "A6061", "A7075", "ACRYLIC", "PET", "EPDM", "MICA", "MIKA", "MC NYLON", "NYLON", "POM", "TEFLON", "HDPE", "PP", "PPE", "PE", "PPS", "PEEK", "PTFE", "PEH", "POLYETHYLENE", "POLYPROPYLENE", "POLYURETHANE", "POLYCARBONATE", "URETHANE", "UHMWPE", "UHMW", "BAKELIT", "BAKELITE", "PU", "ABS", "EPOXY", "SILICON", "RUBBER", "CAO SU")
    For i = LBound(vInsertThreadMaterial) To UBound(vInsertThreadMaterial)
        If InStr(1, UCase(strMaterial), vInsertThreadMaterial(i), vbTextCompare) <> 0 Then blnInsertThreadMaterial = True
    Next
End Sub

Private Sub TrimSpacePhrase(ByRef strPhrase As String)
    strPhrase = Trim(strPhrase)
    Dim intSpacePosition As Integer
    Dim strSpaceReplace As String
    For i = 1 To Len(strPhrase)
        intSpacePosition = InStr(i, strPhrase, " ", vbTextCompare)
        If intSpacePosition <> 0 Then
            While InStr(intSpacePosition + 1, strPhrase, " ", vbTextCompare) = intSpacePosition + 1
                strSpaceReplace = Left(strPhrase, intSpacePosition) & Right(strPhrase, Len(strPhrase) - intSpacePosition - 1)
                strPhrase = strSpaceReplace
            Wend
        End If
    Next i
End Sub

Private Sub FilterThreadStringHoleTable()
    For k = 1 To Len(strThread)
        If InStr(k, strThread, " ", vbTextCompare) Then
            strThread = Left(strThread, InStr(k, strThread, " ", vbTextCompare) - 1)
        End If
        If InStr(k, strThread, "-", vbTextCompare) Then
            strThread = Left(strThread, InStr(k, strThread, "-", vbTextCompare) - 1)
        End If
    Next k
End Sub

Private Sub INSERT_THREAD()
    Dim holeVariables As Variant
    Dim intThreadClassPosition As Integer
    Dim blnCheckInfoThread As Boolean
    ReDim strValue(swModelConfigMgr.Count) As String
    For i = 0 To UBound(vModelConfigMgr)
        swModelConfigMgr.Get3 vModelConfigMgr(i), True, Empty, strValue(i)
        If UCase(vModelConfigMgr(i)) = "MATERIAL" Then strMaterial = strValue(i)
    Next i
    Call TrimSpacePhrase(strMaterial)
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    Do While Not swView Is Nothing
        Set swDispDim = swView.GetFirstDisplayDimension5
        Do While Not swDispDim Is Nothing
LoopGetHoleCallout:
            strGetCalloutAboveDefinition = swDispDim.GetText(swDimensionTextCalloutAboveDefinition)
            strGetCalloutBelow = swDispDim.GetText(swDimensionTextCalloutBelow)
            strGetPrefixTextDefinition = swDispDim.GetText(swDimensionTextPrefixDefinition)
            strGetSuffixTextDefinition = swDispDim.GetText(swDimensionTextSuffixDefinition)
            If swDispDim.IsHoleCallout() Then
                holeVariables = swDispDim.GetHoleCalloutVariables
                If IsEmpty(holeVariables) = False Then
                    For i = 0 To UBound(holeVariables)
                        Set swCalloutVariable = holeVariables(i)
                        If swCalloutVariable.Type = swCalloutVariableType_e.swCalloutVariableType_String Then
                            Set swCalloutStringVariable = swCalloutVariable
                            If swCalloutVariable.UserReadableVariableName = "Thread Description" Then
                                Dim intNumThreadPosition As Integer
                                If i = 2 Then
                                    intThreadClassPosition = InStr(1, strGetPrefixTextDefinition, " - 6H", vbTextCompare)
                                    If intThreadClassPosition <> 0 Then
                                        swDispDim.SetText swDimensionTextPrefixDefinition, Left(strGetPrefixTextDefinition, intThreadClassPosition) & Right(strGetPrefixTextDefinition, Len(strGetPrefixTextDefinition) - intThreadClassPosition - 5)
                                    End If
                                    intNumThreadPosition = InStr(1, strGetCalloutAboveDefinition, "<NUM_INST> x", vbTextCompare)
                                    If intNumThreadPosition <> 0 Then
                                        swDispDim.SetText swDimensionTextPrefixDefinition, "<NUM_INST> - " & swDispDim.GetText(swDimensionTextPrefixDefinition)
                                    End If
                                    swDispDim.SetText swDimensionTextCalloutAbove, swDispDim.GetText(swDimensionTextPrefixDefinition)
                                    swDispDim.SetText swDimensionTextPrefix, ""
                                    GoTo LoopGetHoleCallout
                                ElseIf i = 0 Then
                                    intThreadClassPosition = InStr(1, strGetPrefixTextDefinition, " - 6H", vbTextCompare)
                                    If intThreadClassPosition <> 0 Then
                                        swDispDim.SetText swDimensionTextCalloutAboveDefinition, Left(strGetPrefixTextDefinition, intThreadClassPosition) & Right(strGetPrefixTextDefinition, Len(strGetPrefixTextDefinition) - intThreadClassPosition - 5)
                                    End If
                                End If
                                'Delete "-" Prefix Symbol
                                intNumThreadPosition = InStr(1, strGetPrefixTextDefinition, " - ", vbTextCompare)
                                If intNumThreadPosition = 1 Then
                                    swDispDim.SetText swDimensionTextCalloutAboveDefinition, Right(strGetPrefixTextDefinition, Len(strGetPrefixTextDefinition) - 3)
                                End If
                                'Delete NUM Prefix Symbol
                                intNumThreadPosition = InStr(1, strGetPrefixTextDefinition, "<NUM_INST> x ", vbTextCompare)
                                If intNumThreadPosition <> 0 Then
                                    swDispDim.SetText swDimensionTextCalloutAboveDefinition, Left(swDispDim.GetText(swDimensionTextPrefixDefinition), intNumThreadPosition - 1) & "<NUM_INST> - " & Right(swDispDim.GetText(swDimensionTextPrefixDefinition), Len(strGetPrefixTextDefinition) - intNumThreadPosition - 12)
                                End If
                                If strGetCalloutAboveDefinition = "" Then
                                    swDispDim.SetText swDimensionTextCalloutAbove, swDispDim.GetText(swDimensionTextPrefixDefinition)
                                    swDispDim.SetText swDimensionTextPrefix, ""
                                End If
                                On Error Resume Next
                                Set swCalloutVariable = holeVariables(1)
                                If swCalloutVariable.Type = swCalloutVariableType_e.swCalloutVariableType_Length Then
                                    Set swCalloutLengthVariable = swCalloutVariable
                                    strThreadDepth = 1000 * swCalloutLengthVariable.Length  'Get Callout Thread Depth
                                End If
                                Call CalculateInsertThread(swCalloutStringVariable.String, strThreadDepth)
                                Call GetAddNoteInsertThread(strGetCalloutAboveDefinition)
                                Call GetAddNoteInsertThread(strGetCalloutBelow)
                                Call GetAddNoteInsertThread(strGetPrefixTextDefinition)
                                Call GetAddNoteInsertThread(strGetSuffixTextDefinition)
                                Call RemoveInsertThread(strGetCalloutAboveDefinition)
                                Call RemoveInsertThread(strGetCalloutBelow)
                                Call RemoveInsertThread(strGetPrefixTextDefinition)
                                Call RemoveInsertThread(strGetSuffixTextDefinition)
                                Call CheckInsertThreadMaterial
                                If blnInsertThreadMaterial = True Then
                                    If intInsertThreadPosition = 0 Then
                                        If strThreadSize < 14 Then
                                            flag = False
                                            Select Case swCalloutStringVariable.String
                                            Case "M12"
                                                If InStr(1, UCase(strGetCalloutAboveDefinition), "X1", vbTextCompare) = 0 And InStr(1, UCase(strGetCalloutAboveDefinition), "X1.0", vbTextCompare) = 0 And InStr(1, UCase(strGetCalloutAboveDefinition), "X1.5", vbTextCompare) = 0 And InStr(1, UCase(strGetCalloutAboveDefinition), "X1.25", vbTextCompare) = 0 Then flag = True
                                            Case "M10"
                                                If InStr(1, UCase(strGetCalloutAboveDefinition), "X1", vbTextCompare) = 0 And InStr(1, UCase(strGetCalloutAboveDefinition), "X1.0", vbTextCompare) = 0 And InStr(1, UCase(strGetCalloutAboveDefinition), "X1.25", vbTextCompare) = 0 Then flag = True
                                            Case "M8"
                                                If InStr(1, UCase(strGetCalloutAboveDefinition), "X0.75", vbTextCompare) = 0 And InStr(1, UCase(strGetCalloutAboveDefinition), "X1.0", vbTextCompare) = 0 And InStr(1, UCase(strGetCalloutAboveDefinition), "X1", vbTextCompare) = 0 Then flag = True
                                            Case "M6"
                                                If InStr(1, UCase(strGetCalloutAboveDefinition), "X0.75", vbTextCompare) = 0 Then flag = True
                                            Case "M5"
                                                If InStr(1, UCase(strGetCalloutAboveDefinition), "X0.5", vbTextCompare) = 0 Then flag = True
                                            Case "M4"
                                                If InStr(1, UCase(strGetCalloutAboveDefinition), "X0.5", vbTextCompare) = 0 Then flag = True
                                            Case "M3"
                                                flag = True
                                            Case "M2"
                                                flag = True
                                            End Select
                                            If flag = True Then
                                                If blnCheckDepthThru = False Then
                                                    swDispDim.SetText swDimensionTextPrefix, "(Insert ren " & swCalloutStringVariable.String & "x" & singleInsertThread & "D)"
                                                    If strGetAddNoteInsertThread <> "" Then swDispDim.SetText swDimensionTextCalloutBelow, strGetAddNoteInsertThread
                                                Else
                                                    swDispDim.SetText swDimensionTextPrefix, "(Insert ren)"
                                                    If strGetAddNoteInsertThread <> "" Then swDispDim.SetText swDimensionTextCalloutBelow, strGetAddNoteInsertThread
                                                End If
                                            Else
                                                If strGetAddNoteInsertThread <> "" Then swDispDim.SetText swDimensionTextCalloutBelow, strGetAddNoteInsertThread
                                            End If
                                        End If
                                    End If
                                Else
                                    If strGetAddNoteInsertThread <> "" Then swDispDim.SetText swDimensionTextCalloutBelow, strGetAddNoteInsertThread
                                End If
                                strGetAddNoteInsertThread = ""
                            End If
                        End If
                    Next i
                End If
            Else
                'Get Cosmetic Thread
                Set swDisplayData = swDispDim.GetDisplayData()
                intThreadSymbolPosition = InStr(1, swDisplayData.GetTextAtIndex(0), "M", vbTextCompare)
                If intThreadSymbolPosition <> 0 Then
                    If Len(swDisplayData.GetTextAtIndex(0)) > 1 Then
                        If IsNumeric(Mid(swDisplayData.GetTextAtIndex(0), intThreadSymbolPosition + 1, 4)) Then
                            strThread = Mid(swDisplayData.GetTextAtIndex(0), intThreadSymbolPosition + 1, 4)
                        ElseIf IsNumeric(Mid(swDisplayData.GetTextAtIndex(0), intThreadSymbolPosition + 1, 3)) Then
                            strThread = Mid(swDisplayData.GetTextAtIndex(0), intThreadSymbolPosition + 1, 3)
                        ElseIf IsNumeric(Mid(swDisplayData.GetTextAtIndex(0), intThreadSymbolPosition + 1, 2)) Then
                            strThread = Mid(swDisplayData.GetTextAtIndex(0), intThreadSymbolPosition + 1, 2)
                        ElseIf IsNumeric(Mid(swDisplayData.GetTextAtIndex(0), intThreadSymbolPosition + 1, 1)) Then
                            strThread = Mid(swDisplayData.GetTextAtIndex(0), intThreadSymbolPosition + 1, 1)
                        End If
                        'Get more additional info of insert thread
                        For i = 1 To 5
                            If swDisplayData.GetTextAtIndex(i) <> "" And InStr(1, swDisplayData.GetTextAtIndex(i), "<HOLE-DEPTH>", vbTextCompare) = 0 And InStr(1, UCase(swDisplayData.GetTextAtIndex(i)), "INSERT REN", vbTextCompare) = 0 And IsNumeric(swDisplayData.GetTextAtIndex(i)) = False Then
                                strGetAddNoteInsertThread = swDisplayData.GetTextAtIndex(i)
                            End If
                        Next i
                        Call RemoveInsertThread(strGetCalloutAboveDefinition)
                        Call RemoveInsertThread(strGetCalloutBelow)
                        Call RemoveInsertThread(strGetPrefixTextDefinition)
                        Call RemoveInsertThread(strGetSuffixTextDefinition)
                        Call CheckInsertThreadMaterial
                        If blnInsertThreadMaterial = True Then
                            If strThread < 14 Then
                                blnCheckInfoThread = False
                                If InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M2", vbTextCompare) Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M3", vbTextCompare) Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M4", vbTextCompare) And InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "X0.5", vbTextCompare) = 0 Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M5", vbTextCompare) And InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "X0.5", vbTextCompare) = 0 Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M6", vbTextCompare) And InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "X0.75", vbTextCompare) = 0 Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M8", vbTextCompare) And InStr(InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M8", vbTextCompare) + 1, UCase(swDisplayData.GetTextAtIndex(0)), "X", vbTextCompare) = 0 Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M8X1.25", vbTextCompare) Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M10", vbTextCompare) And InStr(InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M10", vbTextCompare) + 1, UCase(swDisplayData.GetTextAtIndex(0)), "X", vbTextCompare) = 0 Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M10X1.5", vbTextCompare) Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M12", vbTextCompare) And InStr(InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M12", vbTextCompare) + 1, UCase(swDisplayData.GetTextAtIndex(0)), "X", vbTextCompare) = 0 Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                ElseIf InStr(1, UCase(swDisplayData.GetTextAtIndex(0)), "M12X1.75", vbTextCompare) Then
                                    blnCheckInfoThread = True
                                    GoTo LineExportInsertThread
                                End If
LineExportInsertThread:
                                If blnCheckInfoThread = True Then
                                    If strGetCalloutAboveDefinition = "" Then
                                        Call ExportInsertThread(strGetPrefixTextDefinition & strGetSuffixTextDefinition)
                                    Else
                                        Call ExportInsertThread(strGetCalloutAboveDefinition)
                                    End If
                                Else
                                    If strGetAddNoteInsertThread <> "" Then swDispDim.SetText swDimensionTextCalloutBelow, strGetAddNoteInsertThread
                                End If
                            End If
                        Else
                            If strGetAddNoteInsertThread <> "" Then swDispDim.SetText swDimensionTextCalloutBelow, strGetAddNoteInsertThread
                        End If
                        strGetAddNoteInsertThread = ""
                    End If
                End If
            End If
            strThreadDepth = "0"
            Set swDispDim = swDispDim.GetNext3
        Loop
        Set swView = swView.GetNextView
    Loop
    'Info insert thread for Hole Table
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotationType_e.swTableAnnotation_HoleChart Then
            Set swTextFormat = swTable.GetCellTextFormat(0, 1)
            Set swAnn = swTable.GetAnnotation
            swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationHidden
            Call CheckInsertThreadMaterial
            For m = 0 To swTable.ColumnCount - 1
                If InStr(1, UCase(swTable.Text(0, m)), "NOTE", vbTextCompare) <> 0 Or InStr(1, UCase(swTable.Text(0, m)), "GHI CH" & ChrW(769), vbTextCompare) <> 0 Then
                    For i = 1 To swTable.RowCount - 1
                        If swTable.Text(i, m) <> "" And InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare) Then
                            If InStr(1, swTable.Text(i, m), "M", vbTextCompare) Then
                                If InStr(InStr(1, swTable.Text(i, m), "M", vbTextCompare), swTable.Text(i, m), "x", vbTextCompare) Then
                                    If InStr(InStr(1, swTable.Text(i, m), "x", vbTextCompare), swTable.Text(i, m), "D", vbTextCompare) Then
                                        strGetAddNoteInsertThread = Right(swTable.Text(i, m), Len(swTable.Text(i, m)) - InStr(InStr(1, swTable.Text(i, m), "x", vbTextCompare), swTable.Text(i, m), "D", vbTextCompare))
                                    Else
                                        If InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) Then
                                            strGetAddNoteInsertThread = Right(swTable.Text(i, m), Len(swTable.Text(i, m)) - InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) - 1)
                                        End If
                                    End If
                                Else
                                    If InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) Then
                                        strGetAddNoteInsertThread = Right(swTable.Text(i, m), Len(swTable.Text(i, m)) - InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) - 1)
                                    End If
                                End If
                            Else
                                If InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) Then
                                    strGetAddNoteInsertThread = Right(swTable.Text(i, m), Len(swTable.Text(i, m)) - InStr(InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare), UCase(swTable.Text(i, m)), vbCrLf, vbTextCompare) - 1)
                                End If
                            End If
                        ElseIf swTable.Text(i, m) <> "" And InStr(1, UCase(swTable.Text(i, m)), "INSERT REN", vbTextCompare) = 0 Then
                            strGetAddNoteInsertThread = swTable.Text(i, m)
                        End If
                        swTable.Text(i, m) = ""
                        If blnInsertThreadMaterial = True Then
                            For j = 1 To Len(swTable.Text(i, 1))
                                intThreadSymbolPosition = InStr(j, swTable.Text(i, 1), "M", vbTextCompare)
                                If intThreadSymbolPosition <> 0 Then
                                    If IsNumeric(Mid(swTable.Text(i, 1), intThreadSymbolPosition + 1, 4)) Then
                                        strThread = Mid(swTable.Text(i, 1), intThreadSymbolPosition + 1, 4)
                                    ElseIf IsNumeric(Mid(swTable.Text(i, 1), intThreadSymbolPosition + 1, 3)) Then
                                        strThread = Mid(swTable.Text(i, 1), intThreadSymbolPosition + 1, 3)
                                    ElseIf IsNumeric(Mid(swTable.Text(i, 1), intThreadSymbolPosition + 1, 2)) Then
                                        strThread = Mid(swTable.Text(i, 1), intThreadSymbolPosition + 1, 2)
                                    ElseIf IsNumeric(Mid(swTable.Text(i, 1), intThreadSymbolPosition + 1, 1)) Then
                                        strThread = Mid(swTable.Text(i, 1), intThreadSymbolPosition + 1, 1)
                                    End If
                                End If
                            Next j
                            If strThread <> "" Then
                                If strThread < 14 Then
                                    blnCheckInfoThread = False
                                    If InStr(1, UCase(swTable.Text(i, 1)), "M2", vbTextCompare) Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M3", vbTextCompare) Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M4", vbTextCompare) And InStr(1, UCase(swTable.Text(i, 1)), "X0.5", vbTextCompare) = 0 Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M5", vbTextCompare) And InStr(1, UCase(swTable.Text(i, 1)), "X0.5", vbTextCompare) = 0 Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M6", vbTextCompare) And InStr(1, UCase(swTable.Text(i, 1)), "X0.75", vbTextCompare) = 0 Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M8", vbTextCompare) And InStr(InStr(1, UCase(swTable.Text(i, 1)), "M8", vbTextCompare) + 1, UCase(swTable.Text(i, 1)), "X", vbTextCompare) = 0 Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M8X1.25", vbTextCompare) Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M10", vbTextCompare) And InStr(InStr(1, UCase(swTable.Text(i, 1)), "M10", vbTextCompare) + 1, UCase(swTable.Text(i, 1)), "X", vbTextCompare) = 0 Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M10X1.5", vbTextCompare) Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M12", vbTextCompare) And InStr(InStr(1, UCase(swTable.Text(i, 1)), "M12", vbTextCompare) + 1, UCase(swTable.Text(i, 1)), "X", vbTextCompare) = 0 Then
                                        blnCheckInfoThread = True
                                    ElseIf InStr(1, UCase(swTable.Text(i, 1)), "M12X1.75", vbTextCompare) Then
                                        blnCheckInfoThread = True
                                    End If
                                    If blnCheckInfoThread = True Then
                                        intHoleDepthPosition = InStr(1, swTable.Text(i, 1), "<HOLE-DEPTH>", vbTextCompare)
                                        Dim strGetInsertThreadInfo As String
                                        If intHoleDepthPosition <> 0 Then
                                            If InStr(intHoleDepthPosition + 1, swTable.Text(i, 1), "<HOLE-DEPTH>", vbTextCompare) Then
                                                intHoleDepthPosition = InStr(intHoleDepthPosition + 1, swTable.Text(i, 1), "<HOLE-DEPTH>", vbTextCompare)
                                                If InStr(intHoleDepthPosition + 1, swTable.Text(i, 1), vbCrLf, vbTextCompare) Then
                                                    strGetInsertThreadInfo = Left(swTable.Text(i, 1), InStr(intHoleDepthPosition + 1, swTable.Text(i, 1), vbCrLf, vbTextCompare) - 1)
                                                Else
                                                    strGetInsertThreadInfo = swTable.Text(i, 1)
                                                End If
                                                On Error Resume Next
                                                strThreadDepth = Right(strGetInsertThreadInfo, Len(strGetInsertThreadInfo) - intHoleDepthPosition - 11)
                                                Call CalculateInsertThread(Abs(strThread), strThreadDepth)
                                                Call FilterThreadStringHoleTable
                                                If InStr(1, strGetAddNoteInsertThread, "(", vbTextCompare) And InStr(1, strGetAddNoteInsertThread, ")", vbTextCompare) Then
                                                    swTable.Text(i, m) = "Insert ren" & vbCrLf & "M" & strThread & "x" & singleInsertThread & "D" & strGetAddNoteInsertThread
                                                Else
                                                    swTable.Text(i, m) = "Insert ren" & vbCrLf & "M" & strThread & "x" & singleInsertThread & "D" & " (" & strGetAddNoteInsertThread & ")"
                                                End If
                                                swTable.SetCellTextFormat i, m, False, swTextFormat
                                            Else
                                                If InStr(1, swTable.Text(i, 1), vbCrLf, vbTextCompare) Then
                                                    strGetInsertThreadInfo = Left(swTable.Text(i, 1), InStr(1, swTable.Text(i, 1), vbCrLf, vbTextCompare) - 1)
                                                Else
                                                    strGetInsertThreadInfo = swTable.Text(i, 1)
                                                End If
                                                On Error Resume Next
                                                strThreadDepth = Right(strGetInsertThreadInfo, Len(strGetInsertThreadInfo) - intHoleDepthPosition - 11)
                                                Call CalculateInsertThread(Abs(strThread), strThreadDepth)
                                                Call FilterThreadStringHoleTable
                                                If InStr(1, strGetAddNoteInsertThread, "(", vbTextCompare) And InStr(1, strGetAddNoteInsertThread, ")", vbTextCompare) Then
                                                    swTable.Text(i, m) = "Insert ren" & vbCrLf & "M" & strThread & "x" & singleInsertThread & "D" & strGetAddNoteInsertThread
                                                Else
                                                    swTable.Text(i, m) = "Insert ren" & vbCrLf & "M" & strThread & "x" & singleInsertThread & "D" & " (" & strGetAddNoteInsertThread & ")"
                                                End If
                                                swTable.SetCellTextFormat i, m, False, swTextFormat
                                            End If
                                        Else
                                            If strGetAddNoteInsertThread <> "" Then
                                                If InStr(1, strGetAddNoteInsertThread, "(", vbTextCompare) And InStr(1, strGetAddNoteInsertThread, ")", vbTextCompare) Then
                                                    swTable.Text(i, m) = "Insert ren" & vbCrLf & strGetAddNoteInsertThread
                                                Else
                                                    swTable.Text(i, m) = "Insert ren" & vbCrLf & " (" & strGetAddNoteInsertThread & ")"
                                                End If
                                            Else
                                                swTable.Text(i, m) = "Insert ren"
                                            End If
                                            swTable.SetCellTextFormat i, m, False, swTextFormat
                                        End If
                                        strGetAddNoteInsertThread = ""
                                    Else
                                        If strGetAddNoteInsertThread <> "" Then swTable.Text(i, m) = strGetAddNoteInsertThread
                                    End If
                                End If
                            End If
                            strThread = ""
                        Else
                            If strGetAddNoteInsertThread <> "" Then swTable.Text(i, m) = strGetAddNoteInsertThread
                        End If
                        strGetAddNoteInsertThread = ""
                    Next i
                End If
            Next m
            swAnn.Visible = swAnnotationVisibilityState_e.swAnnotationVisible
        End If
        Set swTable = swTable.GetNext
    Wend
End Sub

Private Sub ExportInsertThread(strThreadInfo As String)
    intHoleDepthPosition = InStr(1, strThreadInfo, "<HOLE-DEPTH>", vbTextCompare)
    If intInsertThreadPosition = 0 Then
        If intHoleDepthPosition <> 0 Then
            strThreadDepth = Right(strThreadInfo, Len(strThreadInfo) - intHoleDepthPosition - 11)
            Call CalculateInsertThread(strThread, strThreadDepth)
            swDispDim.SetText swDimensionTextCalloutBelow, "(Insert ren M" & strThread & "x" & singleInsertThread & "D)" & vbCrLf & strGetAddNoteInsertThread
        Else
            swDispDim.SetText swDimensionTextCalloutBelow, "(Insert ren)" & vbCrLf & strGetAddNoteInsertThread
        End If
    End If
End Sub

Private Sub CalculateInsertThread(strThread As String, ByVal strThreadDepth As String)
    If InStr(1, strThread, "M", vbTextCompare) <> 0 Then
        If InStr(1, UCase(strThread), "X", vbTextCompare) = 0 Then
            strThreadSize = Right(Trim(strThread), Len(Trim(strThread)) - 1)
        Else
            strThreadSize = Mid(Trim(strThread), InStr(1, strThread, "M", vbTextCompare) + 1, InStr(1, UCase(strThread), "X", vbTextCompare) - InStr(1, strThread, "M", vbTextCompare) - 1)
        End If
    Else
        strThreadSize = strThread
    End If
    If strThreadDepth = 0 Then
        blnCheckDepthThru = True
    Else
        singleInsertThread = strThreadDepth / strThreadSize
        If (singleInsertThread + 0.5) - CInt(singleInsertThread) <> 0 And CInt(singleInsertThread) - singleInsertThread > 0 Then
            singleInsertThread = CInt(singleInsertThread - 1)
        ElseIf CInt(singleInsertThread) - singleInsertThread < 0 Then
            singleInsertThread = CInt(singleInsertThread)
        End If
        If ((singleInsertThread + 0.5) - CInt(singleInsertThread) <> 0 And (2 * (strThreadDepth - strThreadSize * singleInsertThread) >= strThreadSize) Or singleInsertThread = 0) Then singleInsertThread = singleInsertThread + 0.5
        Select Case strThreadSize
        Case 2
        If singleInsertThread > 4 Then singleInsertThread = 4
        Case 3
        If singleInsertThread > 3.5 Then singleInsertThread = 3.5
        Case 4
        If singleInsertThread > 3.5 Then singleInsertThread = 3.5
        Case 5
        If singleInsertThread > 3 Then singleInsertThread = 3
        Case 6
        If singleInsertThread > 3 Then singleInsertThread = 3
        Case 8
        If singleInsertThread > 3 Then singleInsertThread = 3
        Case 10
        If singleInsertThread > 3 Then singleInsertThread = 3
        Case 12
        If singleInsertThread > 3 Then singleInsertThread = 3
        End Select
        blnCheckDepthThru = False
    End If
End Sub

Private Sub cmdExportExcel_Click()
    Call CheckRun
    blnStatusExportExcel = True
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Call SetNameFirstSheet
    If blnStatusNameFirstSheet = True Then Exit Sub
    Call GetRange(strRangeExportExcel, CInt(UBound(vSheetNames) + txtStartRenameSheet - 1))
    While InStr(1, strTotalSetExportExcel, ".", vbTextCompare) <> 0 Or InStr(1, strTotalSetExportExcel, ",", vbTextCompare) <> 0
        swApp.SendMsgToUser2 "Please input total sets value again!", swMbWarning, swMbOk
        Exit Sub
    Wend
    While strTotalSetExportExcel <= 0 Or strTotalSetExportExcel = "" Or IsNumeric(strTotalSetExportExcel) = False
        swApp.SendMsgToUser2 "Please input total sets value again!", swMbWarning, swMbOk
        Exit Sub
    Wend
    If flag = False Then
        AUTOPARTLIST.Hide
        If chbAllInOne = True Then
            strFilePath = BrowseFolder
            If strFilePath = "" Then Exit Sub
            Call ExportExcel
            blnStatusExportPDF = True
            opbtAllInOneFilePDF = True
            Call EXPORTPDF
            opbtAllToSeparateFilePDF = True
            Call EXPORTPDF
            statusExportCAD2D = True
            If opbtDXF = True Then
                Call EXPORTCAD2D("dxf")
            ElseIf opbtDWG = True Then
                Call EXPORTCAD2D("dwg")
            End If
        Else
            Call ExportExcel
        End If
        blnStatusExportExcel = False
        Unload AUTOPARTLIST
    End If
End Sub

Private Sub ExportExcel()
    Set EndProgram = New Class_ClickESC
    'Check and notice input machine name
    For i = 0 To UBound(vDrawCustPropNames)
        swDrawCustPropMgr.Get3 vDrawCustPropNames(i), False, Empty, strValue(i)
        If UCase(vDrawCustPropNames(i)) = "MACHINE NAME" Then
            Dim strMachineName As String
            strMachineName = strValue(i)
            Exit For
        End If
    Next i
    If strMachineName = Empty Then
        swApp.SendMsgToUser2 "Please input MACHINE NAME information in Drawing Custom Properties and run AUTO EXPORT LIST MANUFACTURING again!", swMbWarning, swMbOk
        End
    End If
    'Open template Excel
    Dim strFolder As String
    strFolder = swApp.GetCurrentMacroPathFolder + "\" + "BM06-QT02TK-Danh sach thiet bi gia cong-OK.xlsx"
    Set xlApp = CreateObject("Excel.Application")
    On Error Resume Next
    Set xlWorkbook = xlApp.Workbooks.Open(strFolder, False, True)
    If xlWorkbook Is Nothing Then
        Set xlWorkbook = xlApp.Workbooks.Open(strFolder)
    End If
    xlApp.Visible = True
    xlApp.WindowState = xlMaximized
    Call ScanWeldedAssemblies
    'Find total necessary sheets of template Excel
    If opbtAllSheetsExportExcel = False And opbtRangeExportExcel = False And opbtCurrentSheetExportExcel = False Then
        opbtAllSheetsExportExcel = True
    End If
    If opbtAllSheetsExportExcel = True Then
        Call CalculateTotalPages(UBound(vSheetNames))
    ElseIf opbtRangeExportExcel = True Then
        Call CalculateTotalPages(iMemory2 - 1)
    ElseIf opbtCurrentSheetExportExcel = True Then
        n = 1
    End If
    strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    'Add sheet template Excel
    Dim xlWorkSheet_Copy As Object
    Set xlWorkSheet = xlWorkbook.Worksheets(1)
    With xlWorkbook.Worksheets(1)
        On Error Resume Next
        .Cells(7, 16).Value = "V" & ChrW(7853) & "t li" & ChrW(7879) & "u"
        On Error Resume Next
        .Cells(7, 17).Value = "Nhi" & ChrW(7879) & "t luy" & ChrW(7879) & "n"
        On Error Resume Next
        .Cells(7, 18).Value = "XLBM"
        .Cells(7, 19).Value = "Kh" & ChrW(7889) & "i l" & ChrW(432) & ChrW(7907) & "ng (kg)"
        On Error Resume Next
        .Cells(7, 20).Value = "Di" & ChrW(7879) & "n tich b" & ChrW(7873) & " m" & ChrW(7863) & "t" & vbCrLf & "(dm^2)"
        On Error Resume Next
        .Cells(7, 21).Value = "CAD"
        On Error Resume Next
        .Range(.Cells(7, 21), .Cells(7, 22)).Merge
        On Error Resume Next
        .Cells(7, 23).Value = "Ma QR"
        On Error Resume Next
        .Cells(7, 24).Value = "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng"
        On Error Resume Next
        .Range("P7:X37").Borders.LineStyle = xlContinuous
        On Error Resume Next
        .Range("P7:X37").Font.Size = 10
        On Error Resume Next
        .Range("P7:X37").Font.name = "Times New Roman"
        On Error Resume Next
        .Range("P7:X7").Font.Bold = True
        On Error Resume Next
        .Range("P7:X7").Interior.Color = RGB(255, 255, 0)
    End With
    If n > 1 Then
        On Error Resume Next
        xlWorkSheet.[_CodeName] = "Sheet1"
        For i = 1 To n - 1
            Dim intBeforeCount As Integer
            intBeforeCount = xlWorkbook.Worksheets.Count
            Set xlWorkSheet_Copy = xlWorkbook.Worksheets(intBeforeCount)
            xlWorkSheet.Copy After:=xlWorkSheet_Copy
            If xlWorkbook.Worksheets.Count = intBeforeCount Then
                Dim xlNewSheet As Object
                Set xlNewSheet = xlWorkbook.Worksheets.Add(After:=xlWorkSheet_Copy)
                xlWorkSheet.Cells.Copy Destination:=xlNewSheet.Cells
            End If
        Next i
    End If
    'Export info into Excel
    For intAddSheet = 1 To n
        xlWorkbook.Worksheets(intAddSheet).Activate
        With xlWorkbook.Worksheets(intAddSheet)
            On Error Resume Next
            .Cells(5, 3).Value = UCase(strMachineName)
            On Error Resume Next
            .Cells(5, 3).WrapText = True
            On Error Resume Next
            .Cells(5, 9).Value = UCase(strPoCode)
            On Error Resume Next
            .Cells(5, 14).Value = "=TODAY()"
            If n > 1 Then
                For i = 1 To 30
                    On Error Resume Next
                    .Cells(i + 7, 1).Value = 30 * (intAddSheet - 1) + i 'add STT
                Next i
            End If
            vSheetNames = swDraw.GetSheetNames()
            If opbtCurrentSheetExportExcel Then
                Dim strCurSheetExcel As String
                strCurSheetExcel = swDraw.GetCurrentSheet.GetName
                For intSheetNumber = 0 To UBound(vSheetNames)
                    If vSheetNames(intSheetNumber) = strCurSheetExcel Then
                        Call ProcessExportExcel(intSheetNumber)
                        Exit For
                    End If
                Next intSheetNumber
            ElseIf opbtAllSheetsExportExcel = True Then
                Dim intStartSheet As Integer
                Dim intEndSheet As Integer
                intStartSheet = 30 * (intAddSheet - 1)
                intEndSheet = 30 * intAddSheet - 1
                If intEndSheet > UBound(vSheetNames) Then intEndSheet = UBound(vSheetNames)
                For intSheetNumber = intStartSheet To intEndSheet
                    Call ProcessExportExcel(intSheetNumber)
                Next intSheetNumber
            ElseIf opbtRangeExportExcel = True Then
                Dim intStartRange As Integer
                Dim intEndRange As Integer
                intStartRange = 30 * (intAddSheet - 1)
                intEndRange = 30 * intAddSheet - 1
                If intEndRange > iMemory2 - 1 Then intEndRange = iMemory2 - 1
                b = 0
                For a = intStartRange To intEndRange
                    intSheetNumber = CInt(intGetRange(a) - txtStartRenameSheet)
                    Call ProcessExportExcel(intSheetNumber)
                    b = b + 1
                Next a
            End If
        End With
        'Set Print Area
        If intAddSheet < n Then xlWorkbook.Worksheets(intAddSheet).PageSetup.PrintArea = "$A$1:$P$37"
    Next intAddSheet
If opbtPartnameExportExcel = True Then
        Call FilterPartNameFromPathName(swModel.GetPathName)
        If cbPrefixExportExcel = "LIST TBGC " Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strGetCurrentPath & "LIST TBGC " & strPartName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ElseIf cbPrefixExportExcel = "(VDTT)TBGC " Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strGetCurrentPath & "(VDTT)TBGC " & strPartName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ElseIf cbPrefixExportExcel = "" Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strGetCurrentPath & strPartName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ElseIf cbPrefixExportExcel <> "" And cbPrefixExportExcel <> "LIST TBGC " And cbPrefixExportExcel <> "(VDTT)TBGC " Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strGetCurrentPath & cbPrefixExportExcel & " " & strPartName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        End If
    ElseIf opbtDrawingNoExportExcel = True Then
        If cbPrefixExportExcel = "LIST TBGC " Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strGetCurrentPath & "LIST TBGC " & UCase(strPoCode) & "-" & strSoMaySoCum & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ElseIf cbPrefixExportExcel = "(VDTT)TBGC " Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strGetCurrentPath & "(VDTT)TBGC " & UCase(strPoCode) & "-" & strSoMaySoCum & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ElseIf cbPrefixExportExcel = "" Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strGetCurrentPath & UCase(strPoCode) & "-" & strSoMaySoCum & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ElseIf cbPrefixExportExcel <> "" And cbPrefixExportExcel <> "LIST TBGC " And cbPrefixExportExcel <> "(VDTT)TBGC " Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strGetCurrentPath & cbPrefixExportExcel & " " & UCase(strPoCode) & "-" & strSoMaySoCum & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        End If
    ElseIf chbAllInOne = True Then
        If cbPrefixExportExcel <> "" Then
            On Error Resume Next
            xlWorkbook.SaveAs FileName:=strFilePath & cbPrefixExportExcel & " " & UCase(strPoCode) & "-" & strSoMaySoCum & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        End If
    End If
    xlApp.WindowState = xlMinimized
    swApp.SendMsgToUser2 "AUTO EXPORT LIST MANUFACTURING COMPLETED!", swMbInformation, swMbOk
    xlApp.WindowState = xlMaximized
End Sub

Private Sub ProcessExportExcel(intSheetNumber As Integer)
    EndProgram.EndByESCKey
    swDraw.ActivateSheet (vSheetNames(intSheetNumber))
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    If swView Is Nothing Then Exit Sub
    strPathName = swView.GetReferencedModelName
    Call OpenModel(strPathName)
    If opbtAllSheetsExportExcel = True Then
        intRow = intSheetNumber Mod 30
    ElseIf opbtRangeExportExcel = True Then
        intRow = b Mod 30
    ElseIf opbtCurrentSheetExportExcel = True Then
        intRow = 0
    End If
    On Error Resume Next
    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 8).Value = "TCI"
    On Error Resume Next
    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 10).Value = strTotalSetExportExcel
    ReDim strExport(swModelConfigMgr.Count) As String
    If Not swModelConfigMgr Is Nothing Then
        For i = 0 To swModelConfigMgr.Count
            On Error Resume Next
            swModelConfigMgr.Get3 vModelConfigMgr(i), False, Empty, strExport(i)
            Select Case UCase(vModelConfigMgr(i))
            Case "PART NAME"
                Call InfoPartname(swOpenModel)
                Call InfoDrawingNo
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 2).NumberFormat = "@"
                If opbtAllSheetsExportExcel = True And intSheetNumber = 0 Then
                    On Error Resume Next
                    xlWorkbook.Worksheets(intAddSheet).Hyperlinks.Add xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 2), strGetCurrentPath & UCase(strPoCode) & "-" & strSoMaySoCum & ".pdf", TextToDisplay:=UCase(strPartName)
                    On Error Resume Next
                    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 2).Font.Size = 10
                    On Error Resume Next
                    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 2).Font.name = "Times New Roman"
                    On Error Resume Next
                    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 2).Font.Underline = False
                    On Error Resume Next
                    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 2).Font.Color = RGB(0, 0, 0)
                Else
                    On Error Resume Next
                    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 2) = UCase(strPartName)
                End If
            Case "DRAWING NO."
                Call InfoDrawingNo
                swModelConfigMgr.Set "DRAWING NO.", strDrawingNo
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Hyperlinks.Add xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 5), strGetCurrentPath & strDrawingNo & ".pdf", TextToDisplay:=UCase(strDrawingNo)
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 5).Font.Size = 10
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 5).Font.name = "Times New Roman"
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 5).Font.Underline = False
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 5).Font.Color = RGB(0, 0, 0)
            Case "QTY"
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 9).Value = CStr(strExport(i))
                On Error Resume Next
                If strExport(i) = "0" Or strExport(i) = "" Then xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 9).Interior.Color = RGB(255, 0, 0)
            Case "MATERIAL"
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 16).Value = strExport(i)
                Call CheckTBTC
                If strExport(i) = "" And strExtension = "SLDPRT" And blnCheckTBTC = False Then
                    On Error Resume Next
                    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 16).Value = "???"
                    On Error Resume Next
                    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 16).Interior.Color = RGB(255, 0, 0)
                End If
            Case "TREATMENT"
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 17).Value = UCase(strExport(i))
            Case "FINISH"
                On Error Resume Next
                xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 18).Value = strExport(i)
            End Select
        Next i
    End If
    On Error Resume Next
    xlWorkbook.Worksheets(intAddSheet).Hyperlinks.Add xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 21), strGetCurrentPath & strDrawingNo & ".dxf", TextToDisplay:="DXF"
    On Error Resume Next
    xlWorkbook.Worksheets(intAddSheet).Hyperlinks.Add xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 22), strGetCurrentPath & strDrawingNo & ".dwg", TextToDisplay:="DWG"
    Call InfoMass
    On Error Resume Next
    xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 13).Value = GetWeldingRemark(strPathName, strExtension)
    'Call RemoveUnsuitFinish
    xlWorkbook.Worksheets(intAddSheet).Columns("P:V").AutoFit
    On Error Resume Next
    xlWorkbook.Worksheets(intAddSheet).Columns("P:V").HorizontalAlignment = xlCenter
    On Error Resume Next
    xlWorkbook.Worksheets(intAddSheet).Columns("P:V").VerticalAlignment = xlCenter
    Call SaveAndCloseModel3D(strPathName)
End Sub

Private Sub CalculateTotalPages(x As Integer)
    Dim totalItems As Long
    totalItems = CLng(x) + 1
    If totalItems <= 0 Then
        n = 1
        m = 0
    Else
        n = CInt((totalItems - 1) \ 30 + 1)
        m = CInt(totalItems Mod 30)
    End If
End Sub

Private Sub CheckTBTC()
    Dim swV As View
    Dim swN As Note
    Dim vPos As Variant
    Dim bInSheet As Boolean
    Dim sNoteText As String

    blnCheckTBTC = False
    Set swV = swDraw.GetFirstView
    While Not swV Is Nothing
        Set swN = swV.GetFirstNote
        While Not swN Is Nothing
            vPos = swN.GetTextPoint2
            bInSheet = False
            If Not IsEmpty(vPos) Then
                If vPos(0) > 0 And vPos(0) < 0.42 And vPos(1) > 0 And vPos(1) < 0.297 Then bInSheet = True
            End If
            If bInSheet Then
                sNoteText = UCase(swN.GetText)
                If InStr(1, sNoteText, "KE MUA", vbTextCompare) <> 0 Or _
                   InStr(1, sNoteText, "K" & ChrW(202) & " MUA", vbTextCompare) <> 0 Or _
                   InStr(1, sNoteText, "TBTC", vbTextCompare) <> 0 Then
                    blnCheckTBTC = True
                End If
            End If
            Set swN = swN.GetNext
        Wend
        Set swV = swV.GetNextView
    Wend
End Sub

Private Sub InfoMass()
    If Not swModelConfigMgr Is Nothing Then
        For i = 0 To swModelConfigMgr.Count
            On Error Resume Next
            If UCase(vModelConfigMgr(i)) = "MATERIAL" Then swModelConfigMgr.Get3 "MATERIAL", False, Empty, strMaterial
        Next i
    End If
    Set swMassProperty = swOpenModel.Extension.CreateMassProperty
    dblVolume = swMassProperty.Volume * 1000
    dblSurfaceArea = swMassProperty.SurfaceArea * 100
    If strMaterial <> "" Then
        Call CalculateMass(strMaterial, dblVolume)
        If Not swModelConfigMgr Is Nothing Then
            For i = 0 To swModelConfigMgr.Count
                On Error Resume Next
                Select Case UCase(vModelConfigMgr(i))
                Case "MASS SW"
                    If dblMass > 0 Then swModelConfigMgr.Set "MASS SW", CStr(Round(dblMass, 4))
                Case "DENSITY"
                    If singleDensity > 0 Then swModelConfigMgr.Set "DENSITY", CStr(singleDensity)
                Case "VOLUME SW"
                    swModelConfigMgr.Set "VOLUME SW", CStr(Round(dblVolume, 6))
                Case "SURFACE AREA SW"
                    swModelConfigMgr.Set "SURFACE AREA SW", CStr(Round(dblSurfaceArea, 4))
                End Select
            Next i
        End If
        If blnStatusExportExcel = True Then
            On Error Resume Next
            xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 19).Value = "0" + CStr(Round(dblMass, 4))
            On Error Resume Next
            xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 20).Value = "0" + CStr(Round(dblSurfaceArea, 4))
        End If
        singleDensity = 0#
    Else
        Call CalculateMass(strMaterial, dblVolume)
        If Not swModelConfigMgr Is Nothing Then
            For i = 0 To swModelConfigMgr.Count
                On Error Resume Next
                Select Case UCase(vModelConfigMgr(i))
                Case "DENSITY"
                    If singleDensity > 0 Then swModelConfigMgr.Set "DENSITY", ""
                Case "VOLUME SW"
                    swModelConfigMgr.Set "VOLUME SW", CDbl(Round(dblVolume, 6))
                Case "SURFACE AREA SW"
                    swModelConfigMgr.Set "SURFACE AREA SW", CDbl(Round(dblSurfaceArea, 4))
                End Select
            Next i
        End If
        On Error Resume Next
        If blnStatusExportExcel = True Then xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 19).Value = ""
    End If
End Sub

Private Sub cbPrefixExportExcel_Change()
    If cbPrefixExportExcel <> "" Then
        opbtCustomExportExcel.Enabled = False
        opbtPartnameExportExcel.Value = True
    Else
        opbtCustomExportExcel.Enabled = True
        opbtCustomExportExcel.Value = True
    End If
End Sub

Private Sub strRangeExportExcel_Change()
    opbtRangeExportExcel.Value = True
End Sub

Private Sub cmdViewMass_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        MsgBox "Please open a PART or an ASSEMBLY!", vbCritical
        End
    ElseIf swModel.GetType = swDocDRAWING Then
        MsgBox "Please open a PART or an ASSEMBLY!", vbCritical
        End
    End If
    Set swOpenModel = swModel
    If swOpenModel.GetType = swDocPART Then
        Call SetConfigModel
        Call InfoMass
        swOpenModel.EditRebuild3
        swApp.ActivateTaskPane (5)
    Else
        Dim swCompChildren As Component2
        Dim vCompChildren As Variant
        Dim swCompModel As ModelDoc2
        dblMassTotal = 0#
        Set swAsm = swOpenModel
        swAsm.ResolveAllLightWeightComponents False
        vComps = swAsm.GetComponents(False)
        Set swFeat = swOpenModel.FirstFeature
        Call CollectStandardDevice
        For j = LBound(vComps) To UBound(vComps)
            Set swComp = vComps(j)
            strExtension = UCase(Right(swComp.GetPathName, 6))
            If InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.IsSuppressed() = False And strExtension = "SLDASM" And swComp.ExcludeFromBOM = False Then
                Set swOpenModel = swComp.GetModelDoc2
                If Not swOpenModel Is Nothing Then
                    Call SetConfigModel
                    If Not swModelConfigMgr Is Nothing Then
                        For i = 0 To swModelConfigMgr.Count
                            On Error Resume Next
                            If UCase(vModelConfigMgr(i)) = "MATERIAL" Then swModelConfigMgr.Get3 "MATERIAL", False, Empty, strMaterial
                        Next i
                    End If
                    If strMaterial <> "" Then
                        vCompChildren = swComp.GetChildren()
                        For i = 0 To UBound(vCompChildren)
                            Set swCompChildren = vCompChildren(i)
                            If InStr(1, UCase(swCompChildren.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swCompChildren.IsSuppressed() = False Then
                                Set swOpenModel = swCompChildren.GetModelDoc2
                                If Not swOpenModel Is Nothing Then
                                    Call SetConfigModel
                                    If Not swModelConfigMgr Is Nothing Then
                                        For k = 0 To swModelConfigMgr.Count
                                            On Error Resume Next
                                            If UCase(vModelConfigMgr(k)) = "MATERIAL" Then swModelConfigMgr.Set2 "MATERIAL", strMaterial
                                        Next k
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
                Call MassStandardDevice
            ElseIf InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.IsSuppressed() = False And strExtension = "SLDASM" And swComp.ExcludeFromBOM Then
                Set swCompModel = swComp.GetModelDoc2
                Set swAsm = swCompModel
                vCompChildren = swAsm.GetComponents(False)
                For k = 0 To UBound(vCompChildren)
                    'Debug.Print swComp.GetPathName
                    Set swCompChildren = vCompChildren(k)
                    If UCase(Right(swCompChildren.GetPathName, 6)) = "SLDPRT" And swCompChildren.IsSuppressed() = False Then
                        Set swOpenModel = swCompChildren.GetModelDoc2
                        If Not swOpenModel Is Nothing Then
                            Call SetConfigModel
                            Call InfoMass
                            If dblMass <> 0 Then dblMassTotal = dblMassTotal - dblMass
                            dblMass = 0
                        End If
                    End If
                Next k
                Call MassStandardDevice
            ElseIf InStr(1, UCase(swComp.GetPathName), UCase("C:\Solidworks Data"), vbTextCompare) = 0 And swComp.IsSuppressed() = False And strExtension = "SLDPRT" And swComp.ExcludeFromBOM = False Then
                'Add the mass of components
                Set swOpenModel = swComp.GetModelDoc2
                If Not swOpenModel Is Nothing Then
                    Call SetConfigModel
                    Call InfoMass
                    If dblMass <> 0 Then dblMassTotal = dblMassTotal + dblMass
                    dblMass = 0
                End If
                'Add the mass of standard device
                Call MassStandardDevice
            End If
        Next
        'Export info mass for assembly
        Set swOpenModel = swModel
        Call SetConfigModel
LoopClickMouseViewMass:
        Call SetClickMouse
        If Not swModelConfigMgr Is Nothing Then
            For i = 0 To swModelConfigMgr.Count
                On Error Resume Next
                If UCase(vModelConfigMgr(i)) = "MASS SW" Then
                    On Error Resume Next
                    swModelConfigMgr.Set "MASS SW", CStr(Round(dblMassTotal, 4))
                    flag = True
                End If
            Next i
        End If
        If flag = False Then
            If Xposition <= DisplayWidth Then
                TheMouse.Move DisplayWidth * 2 - 260, 160, swMouse_MouseMove
                GoTo LoopClickMouseViewMass
            Else
                TheMouse.Move DisplayWidth - 260, 160, swMouse_MouseMove
                GoTo LoopClickMouseViewMass
            End If
        End If
        swOpenModel.EditRebuild3
        swOpenModel.ClearSelection2 True
        swApp.ActivateTaskPane (5)
    End If
    Set swFeat = swOpenModel.FeatureManager.InsertCenterOfMass
    swOpenModel.SetUserPreferenceToggle swUserPreferenceToggle_e.swDisplayCenterOfMassSymbol, True
    swApp.SendMsgToUser2 "INFO MASS COMPLETED!", swMbInformation, swMbOk
    End
End Sub

Private Sub MassStandardDevice()
    Call FilterPartNameFromPathName(swComp.GetPathName)
    For i = 0 To intCountStandardDevice - 1
        If strStandardDevice(i) = strPartName Then
            Set swOpenModel = swComp.GetModelDoc2
            If Not swOpenModel Is Nothing Then
                Call SetConfigModel
                If Not swModelConfigMgr Is Nothing Then
                    For k = 0 To swModelConfigMgr.Count
                        On Error Resume Next
                        If UCase(vModelConfigMgr(k)) = "MASS SW" Then
                            On Error Resume Next
                            swModelConfigMgr.Get3 "MASS SW", False, Empty, strGetMass
                        Else
                            On Error Resume Next
                            swModelConfigMgr.Get3 "MATERIAL", False, Empty, strMaterial
                        End If
                    Next k
                End If
                If strGetMass <> "" And strGetMass <> "0" And strMaterial = "" Then
                    If swComp.ExcludeFromBOM = False Then
                        If strGetMass <> "" Then dblMassTotal = dblMassTotal + strGetMass
                    End If
                End If
                strGetMass = ""
            End If
        End If
    Next i
End Sub

Private Sub cmdAutoSetupLibrary_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    AutoSetupLibrary.Show
End Sub

Private Sub DetermineLoop2()
    If intCountEmptySheet > 0 Then
        swApp.SendMsgToUser2 "This program has deleted empty sheet and auto run AUTO PARTLIST 2D again!", swMbWarning, swMbOk
    End If
    If intCountDuplicateSheet > 0 Or intCountEmptySheet > 0 Then
        swModel.EditRebuild3
        blnStatusMainLoop = False
        vSheetNames = swDraw.GetSheetNames()
        Call GetPoCodeAndSoMaySoCum
        Call MainLoop
    End If
    If opbtAutoSort = True Or opbtAutoNote = True Then
        If swApp.SendMsgToUser2("Have you sure exported all drawings of this assembly?", swMbWarning, swMbYesNo) = swMbHitYes Then
            Call Loop2
        Else
            End
        End If
    End If
    blnStatusMainLoop = False
    If blnStatusSetDrawingNo Then swApp.SendMsgToUser2 "AUTO PARTLIST COMPLETED!", swMbInformation, swMbOk
End Sub

Private Sub Loop2()
    vSheetNames = swDraw.GetSheetNames()
    If opbtCurrentSheetAutoPartlist = True Then
        Dim strCurSheetL2 As String
        strCurSheetL2 = swDraw.GetCurrentSheet.GetName
        For intSheetNumber = 0 To UBound(vSheetNames)
            If vSheetNames(intSheetNumber) = strCurSheetL2 Then
                Call ProcessLoop2(intSheetNumber)
                Exit For
            End If
        Next intSheetNumber
    ElseIf opbtAllSheetsAutoPartlist = True Then
        For intSheetNumber = 0 To UBound(vSheetNames)
            Call ProcessLoop2(intSheetNumber)
        Next intSheetNumber
    ElseIf opbtRangeAutoPartlist = True Then
        For a = 0 To iMemory2 - 1
            intSheetNumber = CInt(intGetRange(a) - txtStartRenameSheet)
            Call ProcessLoop2(intSheetNumber)
        Next a
    End If
End Sub

Private Sub ProcessLoop2(intSheetNumber As Integer)
    EndProgram.EndByESCKey
    swDraw.ActivateSheet (vSheetNames(intSheetNumber))
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    If swView Is Nothing Then Exit Sub
    strExtension = Right(UCase(swView.GetReferencedModelName), 6)
    If strExtension = "SLDASM" Then
        'Refresh Partlist
        chbAutoRefreshPartlist = True
        Call RefreshPartlist
        'Sort Table Annotation
        Call AutoSortPartlist
        'Auto Note
        Call AutoNote
        'Delete empty row
        Call DeleteEmptyRowAndEditItemNumber
    End If
    'Rename sheet
    If InStr(1, txtStartRenameSheet, "0", vbTextCompare) = 1 And intSheetNumber + txtStartRenameSheet < 10 Then
        swDraw.GetCurrentSheet.SetName "0" & intSheetNumber + txtStartRenameSheet & " "
    Else
        swDraw.GetCurrentSheet.SetName " " & intSheetNumber + txtStartRenameSheet & " "
    End If
End Sub

Private Sub AutoSortPartlist()
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotation_BillOfMaterials Then
            Dim swSortData As BomTableSortData
            Set swBOMTable = swTable
            Set swSortData = swBOMTable.GetBomTableSortData
            swSortData.ColumnIndex(0) = 2
            swSortData.Ascending(0) = 1
            swSortData.ColumnIndex(1) = -1
            swSortData.Ascending(1) = 1
            swSortData.ColumnIndex(2) = -1
            swSortData.Ascending(2) = 1
            swSortData.ItemGroups = swBomTableSortItemGroup_e.swBomTableSortItemGroup_None
            swSortData.SortMethod = swBomTableSortMethod_Numeric
            swSortData.DoNotChangeItemNumber = True
            swSortData.SaveCurrentSortParameters = False
            swBOMTable.Sort swSortData
        End If
        Set swTable = swTable.GetNext
    Wend
End Sub

Private Sub AutoNote()
    Set swView = swDraw.GetFirstView
    Set swTable = swView.GetFirstTableAnnotation
    While Not swTable Is Nothing
        If swTable.Type = swTableAnnotation_BillOfMaterials Then
            Dim vAttachedEnts As Variant
            Dim swEnt As Entity
            While Not swView Is Nothing
                vNote = swView.GetNotes
                For i = 0 To UBound(vNote)
                    Set swNote = vNote(i)
                    If swNote.HasBalloon() = False And swNote.GetLeaderCount() > 0 And swNote.IsBomBalloon() Then
                        Set swAnn = swNote.GetAnnotation
                        vAttachedEnts = swAnn.GetAttachedEntities3
                        If IsEmpty(vAttachedEnts) = False Then
                            swAnn.Select3 False, Nothing
                            If opbtAutoNote = True Then
                                Set swNote = swModel.Extension.EditBalloonProperties2(swBS_Circular, swBF_2Chars, swBalloonTextItemNumber, Empty, Empty, Empty, Empty, False, Empty, Empty, Empty)
                            End If
                        End If
                    End If
                    If swNote.HasBalloon() And swNote.GetLeaderCount() > 0 Then
                        Set swAnn = swNote.GetAnnotation
                        If swNote.GetText = "?" Then
                            swAnn.Select3 False, Nothing
                            swModel.EditDelete
                        Else
                            vAttachedEnts = swAnn.GetAttachedEntities3
                            On Error Resume Next
                            If IsEmpty(vAttachedEnts) = False Then
                                Dim vAttEntType As Variant
                                vAttEntType = swAnn.GetAttachedEntityTypes
                                If Not IsEmpty(vAttEntType) Then
                                    For j = 0 To UBound(vAttEntType)
                                        If vAttEntType(j) = swSelSILHOUETTES Then
                                            swAnn.Select3 False, Nothing
                                            swApp.SendMsgToUser2 "Balloon number: " & swNote.GetBomBalloonText(1) & vbCrLf & "This balloon associates with a SILHOUETTE is not supported." & vbCrLf & "Please drag this balloon associate with an EDGE or a FACE on the component and run program again!", swMbWarning, swMbOk
                                            End
                                        End If
                                    Next j
                                End If
                                Set swEnt = vAttachedEnts(0)
                                Set swComp = swEnt.GetComponent
                                'Filter part name in front of the first slashes symbol for standard view
                                Set swBOMTable = swTable
                                Dim swBOMFeature As BomFeature
                                Set swBOMFeature = swBOMTable.BomFeature
                                If swBOMFeature.TableType = swBomType_TopLevelOnly Then
                                    Call FilterPartNameTopLevelTemp(swComp.Name2)
                                    'Filter last part name in front of the first horizontal mark
                                    Call FilterPartNameTopLevelLast(strPartNameTopLevel_Temp)
                                    'Check last part name
                                    FilterPartNameFromPathName swView.GetReferencedModelName
                                    If strPartName = strPartNameTopLevel_Last Then
                                        On Error Resume Next
                                        strPathName = Right(swComp.Name2, Len(swComp.Name2) - intFirstSlashesPosition)
                                        Call FilterPartNameTopLevelTemp(strPathName)
                                        Call FilterPartNameTopLevelLast(strPartNameTopLevel_Temp)
                                    End If
                                Else
                                    If InStrRev(swComp.Name2, "/", -1, vbTextCompare) <> 0 Then
                                        strPartNameTopLevel_Temp = Right(swComp.Name2, Len(swComp.Name2) - InStrRev(swComp.Name2, "/", -1, vbTextCompare))
                                    Else
                                        strPartNameTopLevel_Temp = swComp.Name2
                                    End If
                                    Call FilterPartNameTopLevelLast(strPartNameTopLevel_Temp)
                                End If
                                'Filter partname for last name
                                Call EditPartNameStandardDevice
                                'Auto note
                                For j = 0 To swTable.RowCount - 3
                                    If StrComp(swNote.GetBomBalloonText(1), swTable.Text(j, 0), vbTextCompare) = 0 And swTable.Text(j, 2) = "" Then
                                        swAnn.Select3 False, Nothing
                                        Dim ItemNumber As String
                                        Dim PartNumber As String
                                        Dim GetPathComponentBOM As Variant
                                        Set swBOMTable = swTable
                                        GetPathComponentBOM = swBOMTable.GetModelPathNames(j, ItemNumber, PartNumber)
                                        strExtension = UCase(Right(GetPathComponentBOM(0), 6))
                                        If opbtAutoNote = True Or blnCheckMissingBalloon = True Then
                                            If InStr(1, swComp.ReferencedConfiguration, "Default", vbTextCompare) = 0 And IsNumeric(swComp.ReferencedConfiguration) = False Then
                                                If strExtension = "SLDPRT" Then
                                                    Set swNote = swModel.Extension.EditBalloonProperties2(swBS_Underline, swBF_Tightest, swBalloonTextCustom, swComp.ReferencedConfiguration & vbCrLf & "SL: " & swTable.Text(j, 3) & " pcs", Empty, Empty, Empty, False, Empty, Empty, Empty)
                                                Else
                                                    Set swNote = swModel.Extension.EditBalloonProperties2(swBS_Underline, swBF_Tightest, swBalloonTextCustom, swComp.ReferencedConfiguration & vbCrLf & "SL: " & swTable.Text(j, 3) & " set", Empty, Empty, Empty, False, Empty, Empty, Empty)
                                                End If
                                            Else
                                                If strExtension = "SLDPRT" Then
                                                    Set swNote = swModel.Extension.EditBalloonProperties2(swBS_Underline, swBF_Tightest, swBalloonTextCustom, strPartNameTopLevel_Last & vbCrLf & "SL: " & swTable.Text(j, 3) & " pcs", Empty, Empty, Empty, False, Empty, Empty, Empty)
                                                Else
                                                    Set swNote = swModel.Extension.EditBalloonProperties2(swBS_Underline, swBF_Tightest, swBalloonTextCustom, strPartNameTopLevel_Last & vbCrLf & "SL: " & swTable.Text(j, 3) & " set", Empty, Empty, Empty, False, Empty, Empty, Empty)
                                                End If
                                            End If
                                        Else
                                            swModel.EditDelete 'Delete balloon lead on component is not been exported drawing
                                        End If
                                    End If
                                Next j
                            Else
                                If swNote.GetBomBalloonText(1) <> "" Then
                                    swApp.SendMsgToUser2 "There're a problem with attached balloon in this drawing: Balloon number " & swNote.GetBomBalloonText(1) & vbCrLf & "Please check carefully and run program again!", swMbWarning, swMbOk
                                    End
                                End If
                            End If
                        End If
                    End If
                Next i
                Set swView = swView.GetNextView
            Wend
        End If
        Set swTable = swTable.GetNext
    Wend
End Sub

Private Sub EditPartNameStandardDevice()
    While InStr(1, UCase(strPartNameTopLevel_Last), ".ASM", vbTextCompare) <> 0
        strPartNameTopLevel_Last = Left(strPartNameTopLevel_Last, InStr(1, UCase(strPartNameTopLevel_Last), ".ASM", vbTextCompare) - 1)
    Wend
    While InStr(1, UCase(strPartNameTopLevel_Last), ".PRT", vbTextCompare) <> 0
        strPartNameTopLevel_Last = Left(strPartNameTopLevel_Last, InStr(1, UCase(strPartNameTopLevel_Last), ".PRT", vbTextCompare) - 1)
    Wend
    While InStr(1, UCase(strPartNameTopLevel_Last), "_", vbTextCompare) <> 0
        strPartNameTopLevel_Last = Replace(strPartNameTopLevel_Last, "_", "-", 1, -1, vbTextCompare)
    Wend
End Sub

Private Sub opbtAutoNote_Click()
    opbtAutoSort.Value = True
    opbtAutoSort.Enabled = False
    If opbtAutoNote.Value = False Then
        opbtAutoSort.Value = False
        opbtAutoSort.Enabled = True
    End If
End Sub

Private Sub cmdCheckMissingBalloons_Click()
Call CheckRun
    AUTOPARTLIST.Hide
    Call ErrorHandling
    blnCheckMissingBalloon = True
    Dim strGetBalloonText() As String
    Dim strGetItemNumber() As String
    Dim intTotalBalloonText As Integer
    Dim intTotalItemNumber As Integer
    Dim intCheckBalloon As Integer
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView
    strExtension = Right(UCase(swView.GetReferencedModelName), 6)
    If strExtension = "SLDASM" Then
        Set swView = swDraw.GetFirstView
        Set swTable = swView.GetFirstTableAnnotation
        While Not swTable Is Nothing
            If swTable.Type = swTableAnnotation_BillOfMaterials Then
                'Refresh partlist
                chbAutoRefreshPartlist = True
                Call RefreshPartlist
                'Undo auto note
                Dim vAttachedEnts As Variant
                Dim swEnt As Entity
                Set swView = swDraw.GetFirstView
                While Not swView Is Nothing
                    vNote = swView.GetNotes
                    For i = 0 To UBound(vNote)
                        Set swNote = vNote(i)
                        If swNote.HasBalloon() = False And swNote.GetLeaderCount() = 1 And swNote.IsBomBalloon() Then
                            Set swAnn = swNote.GetAnnotation
                            vAttachedEnts = swAnn.GetAttachedEntities3
                            If IsEmpty(vAttachedEnts) = False Then
                                swAnn.Select3 False, Nothing
                                If swNote.GetBomBalloonText(1) <> "" Then flag = True
                                Set swNote = swModel.Extension.EditBalloonProperties2(swBS_Circular, swBF_2Chars, swBalloonTextItemNumber, Empty, Empty, Empty, Empty, False, Empty, Empty, Empty)
                            End If
                        End If
                    Next i
                    Set swView = swView.GetNextView
                Wend
                'Get info of new BOM
                Set swView = swDraw.GetFirstView
                Set swTable = swView.GetFirstTableAnnotation
                While Not swTable Is Nothing
                    If swTable.Type = swTableAnnotation_BillOfMaterials Then
                        On Error Resume Next
                        ReDim strGetItemNumber(swTable.RowCount - 3) As String
                        Dim ItemNumber As String
                        Dim PartNumber As String
                        Dim GetPathComponentBOM As Variant
                        Dim strGetNameComponentBOMTable() As String
                        On Error Resume Next
                        ReDim strGetNameComponentBOMTable(swTable.RowCount - 3) As String
                        j = 0
                        For i = 0 To swTable.RowCount - 3
                            Set swBOMTable = swTable
                            On Error Resume Next
                            GetPathComponentBOM = swBOMTable.GetModelPathNames(swTable.RowCount - 3 - i, ItemNumber, PartNumber)
                            On Error Resume Next
                            strGetItemNumber(j) = ItemNumber
                            On Error Resume Next
                            strGetNameComponentBOMTable(j) = PartNumber
                            j = j + 1
                        Next i
                        intTotalItemNumber = j
                        'Get info of all balloons
                        Set swView = swDraw.GetFirstView
                        Set swView = swView.GetNextView
                        FilterPartNameFromPathName swView.GetReferencedModelName
                        On Error Resume Next
                        ReDim strGetBalloonText(swTable.RowCount * 5) As String
                        j = 0
                        While Not swView Is Nothing
                            vNote = swView.GetNotes
                            For i = 0 To UBound(vNote)
                                Set swNote = vNote(i)
                                If swNote.HasBalloon() Then
                                    strGetBalloonText(j) = swNote.GetText
                                    j = j + 1
                                End If
                            Next i
                            Set swView = swView.GetNextView
                        Wend
                        intTotalBalloonText = j
                    End If
                    Set swTable = swTable.GetNext
                Wend
                'Restore initialize status
                If flag = True Then
                    Call AutoSortPartlist
                    Call AutoNote
                    Call DeleteEmptyRowAndEditItemNumber
                End If
                'Export result
                Set xlApp = CreateObject("Excel.Application")
                xlApp.Visible = True
                xlApp.WindowState = xlMaximized
                Set xlWorkbook = xlApp.Workbooks.Add
                With xlWorkbook.Worksheets(1)
                    On Error Resume Next
                    .Columns("B").NumberFormat = "@"
                    On Error Resume Next
                    .Columns("C").NumberFormat = "@"
                    On Error Resume Next
                    .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH B" & ChrW(211) & "NG CH" & ChrW(7880) & " M" & ChrW(7908) & "C THI" & ChrW(7870) & "U"
                    On Error Resume Next
                    .Range(.Cells(1, 1), .Cells(1, 3)).Merge
                    On Error Resume Next
                    .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
                    On Error Resume Next
                    .Range(.Cells(2, 1), .Cells(2, 3)).Merge
                    On Error Resume Next
                    .Cells(1, 1).Font.Size = 10
                    On Error Resume Next
                    .Cells(3, 1).Value = "STT"
                    On Error Resume Next
                    .Cells(3, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
                    On Error Resume Next
                    .Cells(3, 3).Value = "B" & ChrW(211) & "NG CH" & ChrW(7880) & " M" & ChrW(7908) & "C"
                    On Error Resume Next
                    .Range("A1:C3").Font.Bold = True
                    k = 0
                    For i = 0 To intTotalItemNumber - 1
                        intCheckBalloon = 0
                        For j = 0 To intTotalBalloonText - 1
                            If strGetBalloonText(j) = strGetItemNumber(i) Then intCheckBalloon = 1
                        Next j
                        If intCheckBalloon = 0 Then
                            On Error Resume Next
                            .Cells(k + 4, 1).Value = k + 1
                            On Error Resume Next
                            .Cells(k + 4, 2).Value = strGetNameComponentBOMTable(i)
                            On Error Resume Next
                            .Cells(k + 4, 3).Value = strGetItemNumber(i)
                            k = k + 1
                        End If
                    Next i
                    If k = 0 Then
                        On Error Resume Next
                        .Range(.Cells(4, 1), .Cells(4, 3)).Merge
                        On Error Resume Next
                        .Cells(4, 1).Value = "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y ch" & ChrW(7881) & " m" & ChrW(7909) & "c  thi" & ChrW(7871) & "u!"
                        On Error Resume Next
                        .Cells(4, 1).Font.Bold = True
                        On Error Resume Next
                        .Cells(4, 1).Font.Italic = True
                        On Error Resume Next
                        .Cells(4, 1).Font.Color = RGB(255, 0, 0)
                    End If
                    On Error Resume Next
                    .Columns("A:C").AutoFit
                    On Error Resume Next
                    .Columns("A:C").HorizontalAlignment = xlCenter
                    On Error Resume Next
                    .Range(.Cells(1, 1), .Cells(k + 3, 3)).Borders.LineStyle = xlContinuous
                End With
            End If
            Set swTable = swTable.GetNext
        Wend
        Unload Me
    Else
        Set swView = swDraw.GetFirstView
        Set swTable = swView.GetFirstTableAnnotation
        While Not swTable Is Nothing
            If swTable.Type = swTableAnnotation_BillOfMaterials Or swTable.Type = swTableAnnotation_WeldmentCutList Then
                'Get info of all balloons
                Set swView = swDraw.GetFirstView
                Set swView = swView.GetNextView
                FilterPartNameFromPathName swView.GetReferencedModelName
                On Error Resume Next
                ReDim strGetBalloonText(swTable.RowCount * 5) As String
                j = 0
                While Not swView Is Nothing
                    vNote = swView.GetNotes
                    For i = 0 To UBound(vNote)
                        Set swNote = vNote(i)
                        If swNote.HasBalloon() Then
                            strGetBalloonText(j) = swNote.GetText
                            j = j + 1
                        End If
                    Next i
                    Set swView = swView.GetNextView
                Wend
                intTotalBalloonText = j
                ReDim strGetItemNumber(swTable.RowCount - 3) As String
                j = 0
                For i = 0 To swTable.RowCount - 3
                    strGetItemNumber(j) = swTable.RowCount - 2 - i
                    j = j + 1
                Next i
                intTotalItemNumber = j
                'Export result
                Set xlApp = CreateObject("Excel.Application")
                xlApp.Visible = True
                xlApp.WindowState = xlMaximized
                Set xlWorkbook = xlApp.Workbooks.Add
                With xlWorkbook.Worksheets(1)
                    On Error Resume Next
                    .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH B" & ChrW(211) & "NG CH" & ChrW(7880) & " M" & ChrW(7908) & "C THI" & ChrW(7870) & "U"
                    On Error Resume Next
                    .Range(.Cells(1, 1), .Cells(1, 2)).Merge
                    On Error Resume Next
                    .Columns("B").ColumnWidth = 20
                    On Error Resume Next
                    .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
                    On Error Resume Next
                    .Range(.Cells(2, 1), .Cells(2, 2)).Merge
                    On Error Resume Next
                    .Cells(1, 1).Font.Size = 10
                    On Error Resume Next
                    .Cells(3, 1).Value = "STT"
                    On Error Resume Next
                    .Cells(3, 2).Value = "B" & ChrW(211) & "NG CH" & ChrW(7880) & " M" & ChrW(7908) & "C"
                    On Error Resume Next
                    .Range("A1:B3").Font.Bold = True
                    k = 0
                    For i = 0 To intTotalItemNumber - 1
                        intCheckBalloon = 0
                        For j = 0 To intTotalBalloonText - 1
                            If strGetBalloonText(j) = strGetItemNumber(i) Then intCheckBalloon = 1
                        Next j
                        If intCheckBalloon = 0 Then
                            On Error Resume Next
                            .Cells(k + 4, 1).Value = k + 1
                            On Error Resume Next
                            .Cells(k + 4, 2).Value = strGetItemNumber(i)
                            k = k + 1
                        End If
                    Next i
                    If k = 0 Then
                        On Error Resume Next
                        .Range(.Cells(4, 1), .Cells(4, 2)).Merge
                        On Error Resume Next
                        .Cells(4, 1).Value = "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y ch" & ChrW(7881) & " m" & ChrW(7909) & "c  thi" & ChrW(7871) & "u!"
                        On Error Resume Next
                        .Cells(4, 1).Font.Bold = True
                        On Error Resume Next
                        .Cells(4, 1).Font.Italic = True
                        On Error Resume Next
                        .Cells(4, 1).Font.Color = RGB(255, 0, 0)
                    End If
                    On Error Resume Next
                    '.Columns("A:B").AutoFit
                    On Error Resume Next
                    .Columns("A:B").HorizontalAlignment = xlCenter
                    On Error Resume Next
                    .Range(.Cells(1, 1), .Cells(k + 3, 2)).Borders.LineStyle = xlContinuous
                End With
            End If
            Set swTable = swTable.GetNext
        Wend
    End If
End Sub

Private Sub RemoveUnsuitFinish()
    If Not swModelConfigMgr Is Nothing Then
        For i = 0 To swModelConfigMgr.Count
            On Error Resume Next
            If UCase(vModelConfigMgr(i)) = "FINISH" Then swModelConfigMgr.Get3 "FINISH", False, Empty, strFinish
        Next i
    End If
    Dim blnRemoveUnsuitFinish As Boolean
    blnRemoveUnsuitFinish = False
    If strFinish <> "" Then
        Dim vCheckFinishMaterial As Variant
        vCheckFinishMaterial = Array("ACRYLIC", "PET", "EPDM", "MICA", "MIKA", "MC NYLON", "NYLON", "POM", "TEFLON", "HDPE", "PP", "PPE", "PE", "PPS", "PEEK", "PTFE", "PEH", "POLYETHYLENE", "POLYPROPYLENE", "POLYURETHANE", "POLYCARBONATE", "URETHANE", "UHMWPE", "UHMW", "BAKELIT", "BAKELITE", "PU", "ABS", "EPOXY", "SILICON", "RUBBER", "CAO SU", "CARBON")
        For i = LBound(vCheckFinishMaterial) To UBound(vCheckFinishMaterial)
            If InStr(1, UCase(strMaterial), vCheckFinishMaterial(i), vbTextCompare) <> 0 Then
                blnRemoveUnsuitFinish = True
                GoTo ProcessRemoveUnsuitFinish
                Exit For
            End If
        Next
        vCheckFinishMaterial = Array("AL", "A2017", "A2024", "A5052", "A6061", "A7075", "ANP79")
        For i = LBound(vCheckFinishMaterial) To UBound(vCheckFinishMaterial)
            If InStr(1, UCase(strMaterial), vCheckFinishMaterial(i), vbTextCompare) <> 0 And UCase(strFinish) <> "T" & ChrW(7848) & "Y TR" & ChrW(7854) & "NG" And InStr(1, UCase(strFinish), "ANOT", vbTextCompare) = 0 And InStr(1, UCase(strFinish), "ALUMITE", vbTextCompare) = 0 Then
                blnRemoveUnsuitFinish = True
                GoTo ProcessRemoveUnsuitFinish
                Exit For
            End If
        Next
        If (InStr(1, UCase(strMaterial), "SUS", vbTextCompare) = 1 Or InStr(1, UCase(strMaterial), "INOX", vbTextCompare) = 1) And InStr(1, UCase(strFinish), "TEFLON", vbTextCompare) = 0 And InStr(1, UCase(strFinish), "PU", vbTextCompare) = 0 And InStr(1, UCase(strFinish), "PHUN", vbTextCompare) = 0 And InStr(1, UCase(strFinish), "R" & ChrW(7916) & "A", vbTextCompare) = 0 And InStr(1, UCase(strFinish), "R" & ChrW(431) & ChrW(777) & "A", vbTextCompare) = 0 And InStr(1, UCase(strFinish), "HIGHLIGHT", vbTextCompare) = 0 And InStr(1, UCase(strFinish), "X" & ChrW(431) & ChrW(7898) & "C M" & ChrW(7900), vbTextCompare) = 0 Then
            blnRemoveUnsuitFinish = True
            GoTo ProcessRemoveUnsuitFinish
        End If
        vCheckFinishMaterial = Array("AL", "A2017", "A2024", "A5052", "A6061", "A7075", "ANP79")
        Dim intCheckMaterial As Integer
        intCheckMaterial = 0
        For i = LBound(vCheckFinishMaterial) To UBound(vCheckFinishMaterial)
            If InStr(1, UCase(strMaterial), vCheckFinishMaterial(i), vbTextCompare) <> 0 Then intCheckMaterial = 1
        Next
        If intCheckMaterial = 0 And (UCase(strFinish) = "T" & ChrW(7848) & "Y TR" & ChrW(7854) & "NG" Or InStr(1, UCase(strFinish), "ANOT", vbTextCompare) <> 0) Then
            swModelConfigMgr.Set "FINISH", "???"
            If blnStatusExportExcel = True Then xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 18).Value = "???"
            strFinish = ""
        End If
        If strMaterial <> "" And blnCheckWeldment_1 = True And strExtension = "SLDPRT" Then
            blnRemoveUnsuitFinish = True
            GoTo ProcessRemoveUnsuitFinish
        End If
ProcessRemoveUnsuitFinish:
        If blnRemoveUnsuitFinish = True Then
            swModelConfigMgr.Set "FINISH", ""
            On Error Resume Next
            If blnStatusExportExcel = True Then xlWorkbook.Worksheets(intAddSheet).Cells(intRow + 8, 18).Value = ""
            strFinish = ""
            blnRemoveUnsuitFinish = False
        End If
    End If
End Sub

Private Sub CalculateMass(strMaterial As String, dblVolume As Double)
    Dim blnPart As Boolean
    blnPart = False
    If UCase(Right(swOpenModel.GetPathName, 6)) = "SLDPRT" Then
        Dim swPart As PartDoc
        Set swPart = swOpenModel
        Dim vSetColorManual As Variant
        vSetColorManual = swPart.MaterialPropertyValues
        blnPart = True
    End If
    Call TrimSpacePhrase(strMaterial)
    Select Case UCase(strMaterial)
    Case "AL"
        singleDensity = 2.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "6061 Alloy"
    Case "A2024"
        singleDensity = 2.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "2024 Alloy"
    Case "A2017"
        singleDensity = 2.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "2014 Alloy"
    Case "A5052"
        singleDensity = 2.68
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "5052-H32 Alloy"
    Case "A6061"
        singleDensity = 2.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "6061 Alloy"
    Case "A7075"
        singleDensity = 2.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "7075-O (SS)"
    Case "ANP79"
        singleDensity = 2.77
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "7079 Alloy"
    Case "ACRYLIC"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "PVC"
        singleDensity = 1.3
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PVC Rigid"
    Case "PVC V" & ChrW(192) & "NG"
        singleDensity = 1.3
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PVC Rigid"
    Case "PET"
        singleDensity = 1.42
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PET"
    Case "EPDM"
        singleDensity = 0.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "EPDM"
    Case "MICA"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MIKA"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MICA " & ChrW(272) & "EN"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MIKA " & ChrW(272) & "EN"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MICA TRONG CH" & ChrW(212) & ChrW(769) & "NG TI" & ChrW(771) & "NH " & ChrW(272) & "I" & ChrW(202) & ChrW(803) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MIKA TRONG CH" & ChrW(212) & ChrW(769) & "NG TI" & ChrW(771) & "NH " & ChrW(272) & "I" & ChrW(202) & ChrW(803) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MICA TRONG CH" & ChrW(7888) & "NG T" & ChrW(296) & "NH " & ChrW(272) & "I" & ChrW(7878) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MIKA TRONG CH" & ChrW(7888) & "NG T" & ChrW(296) & "NH " & ChrW(272) & "I" & ChrW(7878) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MICA " & ChrW(272) & "EN CH" & ChrW(212) & ChrW(769) & "NG TI" & ChrW(771) & "NH " & ChrW(272) & "I" & ChrW(202) & ChrW(803) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MIKA " & ChrW(272) & "EN CH" & ChrW(212) & ChrW(769) & "NG TI" & ChrW(771) & "NH " & ChrW(272) & "I" & ChrW(202) & ChrW(803) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MICA " & ChrW(272) & "EN CH" & ChrW(7888) & "NG T" & ChrW(296) & "NH " & ChrW(272) & "I" & ChrW(7878) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MIKA " & ChrW(272) & "EN CH" & ChrW(7888) & "NG T" & ChrW(296) & "NH " & ChrW(272) & "I" & ChrW(7878) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MICA CH" & ChrW(212) & ChrW(769) & "NG TI" & ChrW(771) & "NH " & ChrW(272) & "I" & ChrW(202) & ChrW(803) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MIKA CH" & ChrW(212) & ChrW(769) & "NG TI" & ChrW(771) & "NH " & ChrW(272) & "I" & ChrW(202) & ChrW(803) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MICA CH" & ChrW(7888) & "NG T" & ChrW(296) & "NH " & ChrW(272) & "I" & ChrW(7878) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MIKA CH" & ChrW(7888) & "NG T" & ChrW(296) & "NH " & ChrW(272) & "I" & ChrW(7878) & "N"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Acrylic (Medium-high impact)"
    Case "MC NYLON"
        singleDensity = 1.15
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Nylon 101"
    Case "MC NYLON " & ChrW(272) & "EN"
        singleDensity = 1.15
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Nylon 101"
    Case "NYLON"
        singleDensity = 1.15
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Nylon 101"
    Case "POM"
        singleDensity = 1.39
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "POM Acetal Copolymer"
    Case "POM " & ChrW(272) & "EN"
        singleDensity = 1.39
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "POM Acetal Copolymer"
    Case "POM TR" & ChrW(7854) & "NG"
        singleDensity = 1.39
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "POM Acetal Copolymer"
    Case "TEFLON"
        singleDensity = 2.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PTFE (general)"
    Case "TEFLON TR" & ChrW(7854) & "NG"
        singleDensity = 2.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PTFE (general)"
    Case "HDPE"
        singleDensity = 0.952
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PE High Density"
    Case "PP"
        singleDensity = 0.89
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PP Copolymer"
    Case "PPE"
        singleDensity = 1.06
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PPE"
    Case "PE"
        singleDensity = 0.917
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PE Low/Medium Density"
    Case "PPS"
        singleDensity = 1.57
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PPS"
    Case "PEEK"
        singleDensity = 1.31
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Polyetheretherketone (PEEK)"
    Case "PTFE"
        singleDensity = 2.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PTFE (general)"
    Case "PEH"
        singleDensity = 0.952
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PE High Density"
    Case "POLYETHYLENE"
        singleDensity = 0.91
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PE Low/Medium Density"
    Case "POLYPROPYLENE"
        singleDensity = 0.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PP Copolymer"
    Case "POLYURETHANE"
        singleDensity = 1.26
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PUR"
    Case "POLYCARBONATE"
        singleDensity = 1.19
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PC High Viscosity"
    Case "URETHANE"
        singleDensity = 1.26
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PUR"
    Case "UHMWPE"
        singleDensity = 0.93
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PE High Density"
    Case "UHMW"
        singleDensity = 0.93
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PE High Density"
    Case "BAKELIT"
        singleDensity = 1.36
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PBT General Purpose"
    Case "BAKELITE"
        singleDensity = 1.36
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PBT General Purpose"
    Case "PU"
        singleDensity = 1.26
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PUR"
    Case "PU V" & ChrW(192) & "NG"
        singleDensity = 1.26
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PUR"
    Case "PU V" & ChrW(192) & "NG"
        singleDensity = 1.26
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PUR"
    Case "PU XANH"
        singleDensity = 1.26
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "PUR"
    Case "ABS"
        singleDensity = 1.02
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "ABS"
    Case "EPOXY"
        singleDensity = 1.56
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Epoxy, Unfilled"
    Case "SILICON RUBBER"
        singleDensity = 1.25
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Silicon Rubber"
    Case "SILICON"
        singleDensity = 2.33
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Silicon"
    Case "RUBBER"
        singleDensity = 1#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Rubber"
    Case "CAO SU"
        singleDensity = 1#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Rubber"
    Case "CAO SU " & ChrW(272) & "EN"
        singleDensity = 1#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Rubber"
    Case "CAO SU " & ChrW(208) & "EN"
        singleDensity = 1#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Rubber"
    Case "CARBON FIBER"
        singleDensity = 1.78
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Hexcel AS4C (3000 Filaments)"
    Case "S" & ChrW(7906) & "I CARBON"
        singleDensity = 1.78
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Hexcel AS4C (3000 Filaments)"
    Case "CEMENT"
        singleDensity = 2.7
    Case "XI M" & ChrW(258) & "NG"
        singleDensity = 2.7
    Case "CERAMIC"
        singleDensity = 2.3
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Ceramic Porcelain"
    Case "G" & ChrW(7894) & "M"
        singleDensity = 2.3
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Ceramic Porcelain"
    Case "BRICK"
        singleDensity = 2.15
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Bricks"
    Case "G" & ChrW(7840) & "CH"
        singleDensity = 2.15
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Bricks"
    Case "GLASS"
        singleDensity = 2.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Glass"
    Case "K" & ChrW(205) & "NH"
        singleDensity = 2.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Glass"
    Case "K" & ChrW(205) & "NH"
        singleDensity = 2.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Glass"
    Case "K" & ChrW(205) & "NH C" & ChrW(431) & ChrW(7900) & "NG L" & ChrW(7920) & "C"
        singleDensity = 2.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Glass"
    Case "K" & ChrW(205) & "NH CL"
        singleDensity = 2.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Glass"
    Case "KINH CL"
        singleDensity = 2.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Glass"
    Case "K" & ChrW(205) & "NH CL"
        singleDensity = 2.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Glass"
    Case "SUS301"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 304"
    Case "SUS304"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 304"
    Case "SUS304 X" & ChrW(431) & ChrW(7898) & "C M" & ChrW(7900)
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 304"
    Case "SUS304 X" & ChrW(431) & ChrW(416) & ChrW(769) & "C M" & ChrW(416) & ChrW(768)
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 304"
    Case "SUS304L"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 304"
    Case "SUS201"
        singleDensity = 7.86
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "201 Annealed Stainless Steel (SS)"
    Case "SUS201 X" & ChrW(431) & ChrW(7898) & "C M" & ChrW(7900)
        singleDensity = 7.86
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "201 Annealed Stainless Steel (SS)"
    Case "SUS201 X" & ChrW(431) & ChrW(416) & ChrW(769) & "C M" & ChrW(416) & ChrW(768)
        singleDensity = 7.86
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "201 Annealed Stainless Steel (SS)"
    Case "SUS303"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 316 Stainless Steel Sheet (SS)"
    Case "SUS303 X" & ChrW(431) & ChrW(7898) & "C M" & ChrW(7900)
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 316 Stainless Steel Sheet (SS)"
    Case "SUS303 X" & ChrW(431) & ChrW(416) & ChrW(769) & "C M" & ChrW(416) & ChrW(768)
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 316 Stainless Steel Sheet (SS)"
    Case "SUS316"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 316 Stainless Steel Sheet (SS)"
    Case "SUS316 X" & ChrW(431) & ChrW(7898) & "C M" & ChrW(7900)
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 304"
    Case "SUS316 X" & ChrW(431) & ChrW(416) & ChrW(769) & "C M" & ChrW(416) & ChrW(768)
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 304"
    Case "SUS316L"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI Type 316L stainless steel"
    Case "SUS317"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Stainless Steel"
    Case "SUS321"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "AISI 321 Annealed Stainless Steel (SS)"
    Case "SUS347"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Stainless Steel"
    Case "SUS309S"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Stainless Steel"
    Case "SUS310S"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Stainless Steel"
    Case "SUS403"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS405"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS410"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS430"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS430LX"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS430J1L"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS434"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS436L"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS440C"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS443J1"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS444"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS409"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "SUS409L"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Stainless Steel (ferritic)"
    Case "CT3"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "C45"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SS400"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SAE1008"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHC"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "O8 Kp"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "Q235"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "Q295"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "Q345"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "Q390"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "Q420"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "Q460"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "Q345B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "Q460D"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "ASTM A36"
        singleDensity = 7.85
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "ASTM A36 Steel"
    Case "ASTM A572"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SS490"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SS50"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SS330"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SS34"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SS41"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SS540"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SS55"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM400A"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM400B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM400C"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM490A"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM490B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM490C"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM490YA"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM490YB"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM520B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM520C"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM570"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM41A"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SM41B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SN400A"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SN400B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SN490B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SWRM10"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SWRM20"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SCM415"
        singleDensity = 7.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "SCM440"
        singleDensity = 7.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "SKS93"
        singleDensity = 7.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "SKS3"
        singleDensity = 7.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "SK3"
        singleDensity = 7.85
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "SKS51"
        singleDensity = 7.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "SKD11"
        singleDensity = 7.85
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "DC53"
        singleDensity = 7.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "SKH51"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Wrought Stainless Steel"
    Case "SKD61"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Wrought Stainless Steel"
    Case "HPM-MG"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Wrought Stainless Steel"
    Case "HPM-37"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Wrought Stainless Steel"
    Case "DH2F"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "ASTM A36 Steel"
    Case "S45C"
        singleDensity = 7.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "1023 Carbon Steel Sheet (SS)"
    Case "S50C"
        singleDensity = 7.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "1023 Carbon Steel Sheet (SS)"
    Case "S55CN"
        singleDensity = 7.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "1023 Carbon Steel Sheet (SS)"
    Case "S55C"
        singleDensity = 7.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "1023 Carbon Steel Sheet (SS)"
    Case "G-STAR"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Carbon Steel"
    Case "S-STAR"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Carbon Steel"
    Case "PX5"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Carbon Steel"
    Case "NAK55"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Carbon Steel"
    Case "NAK80"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Carbon Steel"
    Case "FC"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Carbon Steel"
    Case ChrW(272) & ChrW(7890) & "NG " & ChrW(272) & ChrW(7886)
        singleDensity = 8.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cooper"
    Case ChrW(272) & ChrW(7890) & "NG V" & ChrW(192) & "NG"
        singleDensity = 8.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Brass"
    Case ChrW(272) & ChrW(7890) & "NG"
        singleDensity = 8.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Brass"
    Case "Copper"
        singleDensity = 8.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Copper"
    Case "Brass"
        singleDensity = 8.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Brass"
    Case "CU"
        singleDensity = 8.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Brass"
    Case "C1100P"
        singleDensity = 8.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cooper"
    Case "C1020P"
        singleDensity = 8.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cooper"
    Case "Z3234"
        singleDensity = 8.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cooper"
    Case "C2801P"
        singleDensity = 8.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Brass"
    Case "TP340H"
        singleDensity = 4.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Titanium"
    Case "TITANIUM"
        singleDensity = 4.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Titanium"
    Case "WATER"
        singleDensity = 1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Water"
    Case "N" & ChrW(431) & ChrW(7898) & "C"
        singleDensity = 1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Water"
    Case "AIR"
        singleDensity = 0.0011
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Air"
    Case "KH" & ChrW(212) & "NG KH" & ChrW(205)
        singleDensity = 0.0011
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Air"
    Case "SILVER"
        singleDensity = 10.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Pure Silver"
    Case "B" & ChrW(7840) & "C"
        singleDensity = 10.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Pure Silver"
    Case "LEAD"
        singleDensity = 11.34
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Pure Lead"
    Case "CH" & ChrW(204)
        singleDensity = 11.34
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Pure Lead"
    Case "GOLD"
        singleDensity = 19.32
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Pure Gold"
    Case "V" & ChrW(192) & "NG"
        singleDensity = 19.32
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Pure Gold"
    Case "NAM CH" & ChrW(194) & "M"
        singleDensity = 7.5
    Case "NAM CH" & ChrW(194) & "M"
        singleDensity = 7.5
    Case "NAM CH" & ChrW(194) & "M " & ChrW(272) & ChrW(7844) & "T HI" & ChrW(7870) & "M"
        singleDensity = 7.5
    Case "MAGNET"
        singleDensity = 7.5
    Case "TINH TH" & ChrW(7874) & " TH" & ChrW(7840) & "CH ANH"
        singleDensity = 2.65
    Case "TH" & ChrW(7840) & "CH ANH"
        singleDensity = 2.65
    Case "QUARTZ"
        singleDensity = 2.65
    Case "G" & ChrW(7894)
        singleDensity = 0.56
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Oak"
    Case "WOOD"
        singleDensity = 0.56
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Oak"
    Case ChrW(272) & ChrW(7842)
        singleDensity = 2.6
    Case "STONE"
        singleDensity = 2.6
    Case "SPCC"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-1"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-2"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-4"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-8"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-A"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-D"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-8D"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-08B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-4B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-2B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCC-1B"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCD"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCD-CD"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCE"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCE-SD"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPCEN-SD"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPFC"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPFC-340"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPFC-370"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPFC-390"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "08K"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "08YU"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHT1"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHT2"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHT3"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHT4"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHT1"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHC"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHD"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHE"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SPHF"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK2"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK2"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK4-D"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK4"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK95"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK5"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK85"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK6"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK75"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK7"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "SK65"
        singleDensity = 7.84
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Plain Carbon Steel"
    Case "WC"
        singleDensity = 15.63
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "ZINC"
        singleDensity = 7.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Zinc AC41A Alloy, As Cast"
    Case "URANIUM"
        singleDensity = 18.8
    Case "PLUTONIUM"
        singleDensity = 19.84
    Case "PLATINUM"
        singleDensity = 21.45
    Case "MERCURY"
        singleDensity = 13.54
    Case "MANGANESE"
        singleDensity = 7.325
    Case "MAGNESIUM"
        singleDensity = 1.74
    Case "LITHIUM"
        singleDensity = 0.54
    Case "HYDROGEN"
        singleDensity = 0.07
    Case "ICE"
        singleDensity = 0.9
    Case "DIAMOND"
        singleDensity = 3.5
    Case "CHROMIUM"
        singleDensity = 7.2
    Case "CARBON DIOXIDE"
        singleDensity = 0.3
    Case "G1A"
        singleDensity = 15.63
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G1"
        singleDensity = 15.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G2"
        singleDensity = 14.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G3"
        singleDensity = 14.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G4"
        singleDensity = 14.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G5"
        singleDensity = 14.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G6"
        singleDensity = 13.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G6.5"
        singleDensity = 14#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "S6"
        singleDensity = 13.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G7"
        singleDensity = 13.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "S7"
        singleDensity = 13.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G8"
        singleDensity = 13.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G9"
        singleDensity = 13.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "G10"
        singleDensity = 12.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "5F"
        singleDensity = 14.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "6F"
        singleDensity = 14.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "F65"
        singleDensity = 13.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "7F"
        singleDensity = 13.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "8F"
        singleDensity = 13.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RG1"
        singleDensity = 15#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RG2"
        singleDensity = 14.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RG3"
        singleDensity = 14.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RG4"
        singleDensity = 14.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RG5"
        singleDensity = 14.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RT51"
        singleDensity = 14.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RT52"
        singleDensity = 13.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RT525"
        singleDensity = 13.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RT53"
        singleDensity = 13.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RT54"
        singleDensity = 13.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RT55"
        singleDensity = 13.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RT56"
        singleDensity = 12.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "FF05"
        singleDensity = 14.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "FF10"
        singleDensity = 14.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "FF20"
        singleDensity = 14.3
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "FF40"
        singleDensity = 14#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RFF05"
        singleDensity = 14.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RFF10"
        singleDensity = 14.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RFF20"
        singleDensity = 14.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RFF40"
        singleDensity = 13.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RFF60"
        singleDensity = 13.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RFF70"
        singleDensity = 13#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RFF80"
        singleDensity = 12.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SF05"
        singleDensity = 14.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SF20"
        singleDensity = 14.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SF25"
        singleDensity = 13.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SF30"
        singleDensity = 13.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SF-X"
        singleDensity = 13.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RSF05"
        singleDensity = 14.4
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RSF20"
        singleDensity = 14.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RSF25"
        singleDensity = 13.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RSF30"
        singleDensity = 13.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "RSF-X"
        singleDensity = 13#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SH10"
        singleDensity = 13.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SH15"
        singleDensity = 13.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SH20"
        singleDensity = 13.3
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SH25"
        singleDensity = 13.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SH30"
        singleDensity = 12.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "SH35"
        singleDensity = 13#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "QS15"
        singleDensity = 14.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "QS25"
        singleDensity = 14.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "QS35"
        singleDensity = 14.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "QS65"
        singleDensity = 13.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "Q2"
        singleDensity = 14#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "Q3"
        singleDensity = 13.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "Q5"
        singleDensity = 13.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "Q7"
        singleDensity = 13.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "WS10"
        singleDensity = 14.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "WS20"
        singleDensity = 14#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "GT20"
        singleDensity = 13.6
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "GT26"
        singleDensity = 13.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "DS28"
        singleDensity = 14.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "DS40"
        singleDensity = 13.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "DS60"
        singleDensity = 12.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "TN3"
        singleDensity = 6.7
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Alloy Steel"
    Case "TW4"
        singleDensity = 7.1
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Alloy Steel"
    Case "TW5"
        singleDensity = 7.5
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Alloy Steel"
    Case "CU-W"
        singleDensity = 13.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "CW"
        singleDensity = 13.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "CW-2"
        singleDensity = 13.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "AW-2"
        singleDensity = 14.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "AG-W"
        singleDensity = 14.9
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "HV18"
        singleDensity = 18.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "CV18"
        singleDensity = 18#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "HV28"
        singleDensity = 17.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "MV17"
        singleDensity = 17.2
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Tungsten"
    Case "65MN"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "16MN"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "40XH"
        singleDensity = 8#
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Alloy Steel"
    Case "SUJ2"
        singleDensity = 7.8
        If blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", "Cast Carbon Steel"
    End Select
    If strMaterial = "" And blnPart = True Then swPart.SetMaterialPropertyName2 swConfig.name, "SOLIDWORKS Materials.sldmat", Empty
    If blnPart = True Then swPart.MaterialPropertyValues = vSetColorManual
    dblMass = CDbl(singleDensity) * dblVolume
End Sub

Private Sub Up_QR_Click()
'Call AddQR
'Call QR_offline_Click
                        If opbtCheckInternet = True Then
                            Call DelQR_offline
                            Call AddQR
                        ElseIf opbtCheckInternet = False Then
                            Call Del_QR_RS
                            Call QR_offline_Click 'Call AddQR
                            End If
End Sub

Private Sub UserForm_Initialize()
'    AUTOPARTLIST.ListBox1.AddItem "Author: LE TRUNG HIEU"
    AUTOPARTLIST.ListBox1.AddItem "Company: TRI CUONG INDUSTRIAL CO.,LTD"
    AUTOPARTLIST.ListBox1.AddItem ""
    AUTOPARTLIST.ListBox1.AddItem "Embedded software: AUTO PARTLIST."
    AUTOPARTLIST.ListBox1.AddItem "This application is only developed for Solidworks."
    AUTOPARTLIST.ListBox1.AddItem "Platform: Microsoft Visual Basic for Applications."
    AUTOPARTLIST.ListBox1.AddItem "Release 2019 for free."
    AUTOPARTLIST.ListBox1.AddItem ""
    AUTOPARTLIST.ListBox1.AddItem "Version 1.0's features:"
    AUTOPARTLIST.ListBox1.AddItem "- Auto add info Partname and Drawing No."
    AUTOPARTLIST.ListBox1.AddItem "- Auto sort BOM ascending."
    AUTOPARTLIST.ListBox1.AddItem "- Auto delete standard device's rows."
    AUTOPARTLIST.ListBox1.AddItem "- Auto delete empty sheet."
    AUTOPARTLIST.ListBox1.AddItem "- Auto rename sheet number."
    AUTOPARTLIST.ListBox1.AddItem "- Auto export partlist template Excel."
    AUTOPARTLIST.ListBox1.AddItem "- Backup Custom Properties."
    AUTOPARTLIST.ListBox1.AddItem "- Recovery Custom Properties."
    AUTOPARTLIST.ListBox1.AddItem ""
    AUTOPARTLIST.ListBox1.AddItem "Version 1.1's features:"
    AUTOPARTLIST.ListBox1.AddItem "- Auto delete excess balloons."
    AUTOPARTLIST.ListBox1.AddItem "- Auto add info Remark in BOM."
    AUTOPARTLIST.ListBox1.AddItem "- Auto add info weldment in partlist Excel."
    AUTOPARTLIST.ListBox1.AddItem "- Sort BOM later."
    AUTOPARTLIST.ListBox1.AddItem "- Edit partlist's item number ascending by 1."
    AUTOPARTLIST.ListBox1.AddItem "- Optinally install (include unzip folder)."
    AUTOPARTLIST.ListBox1.AddItem "- Debug with only one drawing."
    AUTOPARTLIST.ListBox1.AddItem ""
    AUTOPARTLIST.ListBox1.AddItem "Version 1.2's features:"
    AUTOPARTLIST.ListBox1.AddItem "- Auto note for standard devices."
    AUTOPARTLIST.ListBox1.AddItem "- Auto set the print area of partlist Excel."
    AUTOPARTLIST.ListBox1.AddItem "- Auto check duplicate sheets in drawing."
    AUTOPARTLIST.ListBox1.AddItem "- Separate info for different configurations."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade backup and recovery tool."
    AUTOPARTLIST.ListBox1.AddItem "- Remove auto set class 2."
    AUTOPARTLIST.ListBox1.AddItem ""
    AUTOPARTLIST.ListBox1.AddItem "Version 2.0's features:"
    AUTOPARTLIST.ListBox1.AddItem "- Auto quantity for model 3D."
    AUTOPARTLIST.ListBox1.AddItem "- Export Excel manufacturing devices's quantities."
    AUTOPARTLIST.ListBox1.AddItem "- Export Excel standard devices's quantities."
    AUTOPARTLIST.ListBox1.AddItem "- Auto export PDF for drawing."
    AUTOPARTLIST.ListBox1.AddItem "- Auto export DXF/DWG for drawing."
    AUTOPARTLIST.ListBox1.AddItem "- Auto add insert thread in drawing."
    AUTOPARTLIST.ListBox1.AddItem "- Auto delete checksheet box in assembly drawing."
    AUTOPARTLIST.ListBox1.AddItem "- Auto PO Code for external components."
    AUTOPARTLIST.ListBox1.AddItem "- Auto edit BOM cut list partlist."
    AUTOPARTLIST.ListBox1.AddItem "- Export data in ranges."
    AUTOPARTLIST.ListBox1.AddItem "- Export more info of partlist Excel form."
    AUTOPARTLIST.ListBox1.AddItem "- Auto add info editing drawing in partlist Excel."
    AUTOPARTLIST.ListBox1.AddItem "- Rename sheet by random input number."
    AUTOPARTLIST.ListBox1.AddItem "- Other tools:"
    AUTOPARTLIST.ListBox1.AddItem "  > Set sheet scale"
    AUTOPARTLIST.ListBox1.AddItem "  > Auto update sheet format"
    AUTOPARTLIST.ListBox1.AddItem "  > Auto check excess drawings"
    AUTOPARTLIST.ListBox1.AddItem "  > Auto check missing drawings"
    AUTOPARTLIST.ListBox1.AddItem "  > Check/uncheck exclude from BOM"
    AUTOPARTLIST.ListBox1.AddItem "  > View mass"
    AUTOPARTLIST.ListBox1.AddItem "  > Auto remove envelope"
    AUTOPARTLIST.ListBox1.AddItem "  > Auto remove insert thread"
    AUTOPARTLIST.ListBox1.AddItem "  > Auto remove exist Drawing No."
    AUTOPARTLIST.ListBox1.AddItem ""
    AUTOPARTLIST.ListBox1.AddItem "Version 2.1's features:"
    AUTOPARTLIST.ListBox1.AddItem "- Summary of semi-standard devices."
    AUTOPARTLIST.ListBox1.AddItem "- Summary of mechanical standard devices."
    AUTOPARTLIST.ListBox1.AddItem "- Summary of electrical standard devices."
    AUTOPARTLIST.ListBox1.AddItem "- Debug partname of index number(000E-01...vv)."
    AUTOPARTLIST.ListBox1.AddItem "- Debug zero quantity value of AUTO QTY feature."
    AUTOPARTLIST.ListBox1.AddItem "- Exit program execution by ESC button."
    AUTOPARTLIST.ListBox1.AddItem "- Update material library (100+)."
    AUTOPARTLIST.ListBox1.AddItem "- Tool Export Excel for Check Excess Drawings."
    AUTOPARTLIST.ListBox1.AddItem "- Limited the ranges of insert thread tool."
    AUTOPARTLIST.ListBox1.AddItem "- Auto remove unsuitable FINISH of material."
    AUTOPARTLIST.ListBox1.AddItem "- Checkbox for total sheets."
    AUTOPARTLIST.ListBox1.AddItem "- More other tools:"
    AUTOPARTLIST.ListBox1.AddItem "  > Check components excluded from BOM"
    AUTOPARTLIST.ListBox1.AddItem "  > Delete/undelete roughness box"
    AUTOPARTLIST.ListBox1.AddItem "  > Undo auto note"
    AUTOPARTLIST.ListBox1.AddItem "  > Check missing balloons"
    AUTOPARTLIST.ListBox1.AddItem "  > Rename (drawing) components"
    AUTOPARTLIST.ListBox1.AddItem ""
    AUTOPARTLIST.ListBox1.AddItem "Version 2.2's features:"
    AUTOPARTLIST.ListBox1.AddItem "- Tool compare exist Q'TY with calculated Q'TY."
    AUTOPARTLIST.ListBox1.AddItem "- Auto refresh partlist."
    AUTOPARTLIST.ListBox1.AddItem "- Auto add info item number of editing drawing."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Auto note."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Insert thread remain additional info."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Auto update sheet format."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Check excess drawings."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Check missing drawings."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Check missing balloons."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Auto remove insert thread."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Auto remove unsuitable FINISH."
    AUTOPARTLIST.ListBox1.AddItem "- Upgrade tool Auto edit BOM cut list."
    AUTOPARTLIST.ListBox1.AddItem "- Link custom material with materials library."
    AUTOPARTLIST.ListBox1.AddItem "- Hyperlinks forward 3D components of lists Excel."
    AUTOPARTLIST.ListBox1.AddItem "- More other tools:"
    AUTOPARTLIST.ListBox1.AddItem "  > Auto setup library"
    AUTOPARTLIST.ListBox1.AddItem "  > Rename (3D) components"
    AUTOPARTLIST.ListBox1.AddItem ""
    AUTOPARTLIST.ListBox1.AddItem "Version 2.3's features:"
    AUTOPARTLIST.ListBox1.AddItem "- Export All In One for list manufacturing."
    AUTOPARTLIST.ListBox1.AddItem "- Debug click mouse Auto Q'TY."
    AUTOPARTLIST.ListBox1.AddItem "- Debug click mouse Auto edit BOM cut list."
    AUTOPARTLIST.ListBox1.AddItem "- Check missing balloons for cut lists."
    AUTOPARTLIST.ListBox1.AddItem "- Summary of hydraulics & pneumatics devices."
    AUTOPARTLIST.ListBox1.AddItem "- Rename (3D) components as 3D order."
    AUTOPARTLIST.ListBox1.AddItem ""
        AUTOPARTLIST.ListBox1.AddItem "Version 2.4's features:"
    AUTOPARTLIST.ListBox1.AddItem "- Add QR code each Part sheet drawing."
    AUTOPARTLIST.ListBox1.AddItem "- Save PDF/DXF-DWG to spec-Folder in Model Folder."
    AUTOPARTLIST.ListBox1.AddItem "- Save All Part to Step to spec-Folder in Model Folder."
    AUTOPARTLIST.ListBox1.AddItem ""
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If Not swModel Is Nothing Then
        If swModel.GetType = swDocDRAWING Then
            Set swDraw = swModel
            vSheetNames = swDraw.GetSheetNames()
            txtStartRenameSheet.Value = vSheetNames(0)
        End If
    End If
    cbExportCAD2D.AddItem "R2007"
    cbExportCAD2D.AddItem "R2010"
    cbExportCAD2D.AddItem "R2013"
    cbPrefixExportExcel.AddItem ""
    cbPrefixExportExcel.AddItem "LIST TBGC "
    cbPrefixExportExcel.AddItem "(VDTT) LIST TBGC "
End Sub

'þ======== TAO FOLDER QR

Private Sub CheckFolderExists()
Dim path As String
Dim strFolderExists As String
    path = ThisWorkbook.path & Application.PathSeparator & "\QRCode"
    strFolderExists = Dir(path, vbDirectory)
    If strFolderExists = "" Then
        MkDir path
        Call InsertQR
        Call PutTheQR
    Else
        Call InsertQR
        Call PutTheQR
    End If
End Sub

'------------- CHECK ACTIVE

Private Sub CheckRun()
    Set swApp = Application.SldWorks
    Dim fso As Object
    Dim bActive As Boolean
    Dim fldSW As Object
    Dim subFld As Object
    Dim checkPath As String

    bActive = False
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FolderExists("C:\ProgramData\SOLIDWORKS") Then
            Set fldSW = fso.GetFolder("C:\ProgramData\SOLIDWORKS")
            For Each subFld In fldSW.SubFolders
                If Left(UCase(Trim(subFld.name)), 13) = "SOLIDWORKS 20" Then
                    checkPath = subFld.path & "\lang\turkish\xmls.chemax"
                    If fso.FolderExists(checkPath) Then
                        bActive = True
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    On Error GoTo 0

    If Not bActive Then
        AUTOPARTLIST.Hide
        ActiveAP.Show
        End
    End If
End Sub


Private Sub AddQR()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    Call Del_QR_RS ' Hàm này c?a anh không có trong code g?i kèm, gi? nguyên g?i hàm
    Call CheckQRFolderExists
    Call AddQR2
End Sub

' '--=======================================================================================

Private Sub AddQR2()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    Set swSheet = swDraw.GetCurrentSheet()
    Dim GetSheetProperties As Variant
    Dim sFile As String
    Dim strFileExists As String
    Dim TL1 As Double
    Dim TL2 As Double
    Dim TL As Double
    strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    sFile = strGetCurrentPath & "\QRCode\" & strDrawingNo & ".bmp"
    ' FIX 1: T?m d?ng 0.5s d? Windows hoàn t?t vi?c luu file .bmp và m? khóa file
    Dim t As Single
    t = Timer
    Do While Timer < t + 0.5 And Timer >= t
        DoEvents
    Loop
    strFileExists = Dir(sFile)
    If strFileExists = "" Then
        MsgBox "M" & ChrW(227) & " QRCode: " & strDrawingNo & " KHONG TON TAI. Kiem tra lai va tao QRCode voi ten: " & strDrawingNo & " vao Folder: " & strGetCurrentPath & "\QRCode **** VA CHAY LAI ** AUTO PARTLIST **"
    Else
        GetSheetProperties = swSheet.GetProperties2()
        TL1 = GetSheetProperties(2)
        TL2 = GetSheetProperties(3)
        TL = TL2 / TL1
        Set Part = swApp.ActiveDoc
        ' FIX 2: Ðua b?n v? v? tr?ng thái chu?n c?a Sheet thay vì dùng InsertSketch
        Part.ClearSelection2 True
        swDraw.EditSheet
        Dim SkPicture As Object
        Set SkPicture = Part.SketchManager.InsertSketchPicture(sFile)
        Dim vSheetProps As Variant
        vSheetProps = swSheet.GetProperties
        If Not SkPicture Is Nothing Then
            If vSheetProps(0) = 8 Then
                SkPicture.SetOrigin (0.372 * TL), (0.245 * TL)
                SkPicture.SetSize (0.036 * TL), (0.036 * TL), 1
            ElseIf vSheetProps(0) = 9 Then
                SkPicture.SetOrigin (0.543 * TL), (0.365 * TL)
                SkPicture.SetSize (0.036 * TL), (0.036 * TL), 1
            End If
        End If
        ' FIX 3: Gi?i phóng hoàn toàn d?i tu?ng ?nh kh?i RAM tru?c khi Rebuild
        Set SkPicture = Nothing
    End If
    ' FIX 4: Nhu?ng quy?n CPU 0.5s cho d? h?a SolidWorks k?p load hình lên màn hình
    t = Timer
    Do While Timer < t + 0.5 And Timer >= t
        DoEvents
    Loop
    ' FIX 5: Dùng ForceRebuild thay vì EditRebuild3 d? tránh crash sâu
    swModel.ForceRebuild3 False
End Sub

' '--=======================================================================================

Private Sub InsertQR()
    Dim name As String
    name = strDrawingNo
    Set swApp = Application.SldWorks
    Dim userInput As String
    userInput = name
    Set Logo = QRCodegenBarcode(userInput) ' Hàm này c?a anh t? d?nh nghia
    AUTOPARTLIST.Image2.Picture = Logo
    SavePicture AUTOPARTLIST.Image2.Picture, strGetCurrentPath & "\QRCode\" & name & ".bmp"
    SavePicture AUTOPARTLIST.Image2.Picture, "C:\Temp\pic3.bmp"
End Sub

Function ShowPic(PICPIC As String) As String
    Dim AC As Range
    Dim PicFile As String
    PICPIC = PicFile & ".bmp"
    On Error GoTo Done
    Set AC = Application.Caller
    ActiveSheet.Shapes.AddPicture(ThisWorkbook.path & Application.PathSeparator & PICPIC, False, True, AC.Left, AC.Top, 20, 20).name = "QR"
    ShowPic = ""
    Exit Function
Done:
    ShowPic = "Error"
End Function

Private Sub PutTheQR()
    Dim val As String
    Worksheets("TBGC").Cells(8, 23).Select
    val = ActiveCell.Offset(0, -18).Value
    Do While val <> ""
        ActiveCell.FormulaR1C1 = "=ShowPic(RC[-18])"
        ActiveCell.ClearContents
        ActiveCell.Offset(1, 0).Activate
        val = ActiveCell.Offset(0, -18).Value
    Loop
End Sub

'þ==================================
'Xoa QR code cu (Ðã FIX l?nh l?i th?i)

Private Sub DelQR()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    swModel.ClearSelection2 True
    ' FIX 3: Dùng SelectByID2 thay cho SelectByName
    Dim bRet As Boolean
    bRet = swModel.Extension.SelectByID2("Sketch Picture1", "SKETCHBITMAP", 0, 0, 0, False, 0, Nothing, 0)
    If bRet Then swModel.EditDelete
End Sub

'þ======== Tao folder QR

Private Sub CheckQRFolderExists()
    AUTOPARTLIST.Hide
    Set EndProgram = New Class_ClickESC
    Dim PathQR As String
    Dim strFolderExists As String
    strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    PathQR = strGetCurrentPath & "\QRCode\"
    strFolderExists = Dir(PathQR, vbDirectory)
    If strFolderExists = "" Then
        MkDir PathQR
    End If
    Call InsertQR
End Sub

'========================= TAO QR OFFLINE

Private Sub QR_offline_Click()
    Call DelQR_offline
    Call InsertQR_ofline
End Sub

Private Sub Quit_word_Click()
    Dim x As Variant
    x = Shell("powershell.exe kill -processname winword", 1)
End Sub

Private Sub DeleteAllOleItemsInSheet(swApp As Object, swModel As Object, swSheet As Object)
    Dim nOleCnt As Long
    Dim i As Long
    Dim nSize As Long
    Dim nAspect As Long
    Dim bRet As Boolean
    nOleCnt = swSheet.GetOLEObjectCount
    If 0 = nOleCnt Then Exit Sub
    For i = 0 To nOleCnt - 1
        bRet = swModel.Extension.SelectByID2("", "OLEITEM", 0.5 * (0.372 + 0.41), 0.5 * (0.282 + 0.23), 0, False, 0, Nothing, 0)
        bRet = swModel.DeleteSelection(False)
    Next i
End Sub

Private Sub DelQR_offline()
    Dim swModel As Object
    Dim swDraw As Object
    Dim swSheet As Object
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    Set swSheet = swDraw.GetCurrentSheet
    DeleteAllOleItemsInSheet swApp, swModel, swSheet
End Sub

'================================
' Tao QR offline

Private Sub InsertQR_ofline()
    Dim name1 As String
    name1 = strDrawingNo
    Const BarcodeWidth As Integer = 156
    Dim WdApp As Object
    Set WdApp = CreateObject("Word.Application")
    WdApp.Visible = True
    WdApp.WindowState = 1 ' wdWindowStateMaximize
    With WdApp.Documents.Add
        .PageSetup.RightMargin = .PageSetup.PageWidth - .PageSetup.LeftMargin - BarcodeWidth + 52
        .Fields.Add(Range:=.Range, Type:=-1, Text:="DISPLAYBARCODE " & name1 & " QR \q 3", PreserveFormatting:=False).Copy
        ' Ð?i Windows dua d? li?u lên Clipboard
        Dim t As Single
        t = Timer
        Do While Timer < t + 1
            DoEvents
        Loop
        Call AddQR_offline1
        WdApp.Quit SaveChanges:=0 ' wdDoNotSaveChanges
    End With
End Sub

Private Sub AddQR_offline1()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    Set Part = swApp.ActiveDoc
    Dim boolstatus As Boolean
    Dim sheetName As String
    sheetName = swDraw.GetCurrentSheet().GetName()
    Part.ClearSelection2 True
    boolstatus = Part.Extension.SelectByID2(sheetName, "SHEET", 0.372, 0.282, 0, False, 0, Nothing, 0)
    ' Ch?c ch?n dang thao tác trên Sheet
    swDraw.EditSheet
    ' Th?c hi?n dán OLE item
    Part.Paste
    boolstatus = Part.Extension.SelectByID2("Document", "OLEITEM", 0.372, 0.282, 0, False, 0, Nothing, 0)
    ' Ch? d? h?a OLE load xong tru?c khi cho phép ngu?i dùng thao tác ti?p
    Dim t As Single
    t = Timer
    Do While Timer < t + 0.5 And Timer >= t
        DoEvents
    Loop
    swModel.ClearSelection2 True
End Sub

'þ========TAO FOLDER PDF

Private Sub CheckPDFFolderExists()
    AUTOPARTLIST.Hide
    Set EndProgram = New Class_ClickESC
'    Call ErrorHandling1
Dim PathPDF As String
Dim strFolderExists As String
strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    PathPDF = strGetCurrentPath + "\PDF\"
    strFolderExists = Dir(PathPDF, vbDirectory)
    If strFolderExists = "" Then
        MkDir PathPDF
        Call cmdEXPORTPDF_Click_1
    Else
        Call cmdEXPORTPDF_Click_1
    End If
End Sub

'þ========TAO FOLDER DXF/DWG

Private Sub CheckDXFFolderExists()
    AUTOPARTLIST.Hide
    Set EndProgram = New Class_ClickESC
'    Call ErrorHandling1
Dim Pathdxf As String
Dim strFolderExists As String
strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    Pathdxf = strGetCurrentPath + "\DXF-DWG\"
    strFolderExists = Dir(Pathdxf, vbDirectory)
    If strFolderExists = "" Then
        MkDir Pathdxf
        Call cmdEXPORTDXF_DWG_1
    Else
        Call cmdEXPORTDXF_DWG_1
    End If
End Sub

'þ======== Tao folder STEP

Private Sub CheckSTPFolderExists()
    AUTOPARTLIST.Hide
    Set EndProgram = New Class_ClickESC
    Call ErrorHandling2
Dim pathSTP As String
Dim strFolderExists As String
strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    pathSTP = strGetCurrentPath + "\STEP\"
    strFolderExists = Dir(pathSTP, vbDirectory)
    If strFolderExists = "" Then
        MkDir pathSTP
        Call STEP_expt
    Else
        Call STEP_expt
    End If
End Sub

Private Sub STEP_exp_Click()
Call CheckSTPFolderExists
End Sub

Private Sub STEP_expt()
Dim swApp As SldWorks.SldWorks
Dim swAssy As SldWorks.AssemblyDoc
Dim swConf As SldWorks.Configuration
Dim swRootComp As SldWorks.Component2
Dim retVal As Boolean
Dim errors As Long, warnings As Long
Dim revision As String
Dim vaultPath As String
Dim name As String
Dim pathSTP As String
Dim longstatus As Long, longwarnings As Long
Set swApp = Application.SldWorks
Set swAssy = swApp.ActiveDoc
Set swConf = swAssy.GetActiveConfiguration
Set swRootComp = swConf.GetRootComponent3(True)
strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    pathSTP = strGetCurrentPath + "\STEP\"
vaultPath = pathSTP '"C:\Users\RD-11\Desktop\111\" 'set folder for vault (change this later)
TraverseComponent swRootComp, 1, vaultPath
'Call TraverseComponent
 End Sub

Private Sub TraverseComponent(swComp As SldWorks.Component2, nLevel As Long, vaultPath As String)
'Dim swApp As SldWorks.SldWorks
'Dim swAssy As SldWorks.AssemblyDoc
'Dim swConf As SldWorks.Configuration
'Dim swRootComp As SldWorks.Component2
'Dim retVal As Boolean
Dim errors As Long, warnings As Long
'Dim revision As String
Dim name As String
Dim pathSTP As String
Dim longstatus As Long, longwarnings As Long
    Dim vChilds As Variant, vChild As Variant
    Dim swChildComp As SldWorks.Component2
    Dim MyString As String
    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Set swApp = Application.SldWorks
    vChilds = swComp.GetChildren
    For Each vChild In vChilds
        Set swChildComp = vChild
        Dim FileName As String
        FileName = swChildComp.GetPathName
        FileName = Left(FileName, InStr(FileName, ".") - 1)
        FileName = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
        Debug.Print "Part Name    : " & FileName
        MyString = FileName
        Dim ActiveConfig As String
        ActiveConfig = swChildComp.ReferencedConfiguration
        Debug.Print "Configuration: " & ActiveConfig
        FileName = swChildComp.GetPathName
        If UCase(Right(FileName, 6)) = "SLDPRT" Then
            'MsgBox ("part found")
            Dim swPart As SldWorks.ModelDoc2
            Set swPart = swApp.OpenDoc6(swChildComp.GetPathName, 1, 0, "", longstatus, longwarnings)
            'Dim name As String 'I tried adding this but it made no difference
            name = swPart.GetTitle 'get the title of the active document
            'chop the extension off if present
            If Right(name, 7) = ".SLDPRT" Or Right(name, 7) = ".SLDasm" Then
                name = Left(name, Len(name) - 7)
            End If
            Set swCustPropMgr = swPart.Extension.CustomPropertyManager("") 'get properties
            revision = swCustPropMgr.Get("Revision") 'get revision
            retVal = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP, 214) 'change the step file options
            'save with revision if present
            If revision = "" Or revision = Null Then
            retVal = swPart.Extension.SaveAs(vaultPath & name & ".step", swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errors, warnings)
            Else
            retVal = swPart.Extension.SaveAs(vaultPath & name & " - Rev " & revision & ".step", swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errors, warnings)
            End If
            swApp.CloseDoc swPart.GetTitle
        End If
        Debug.Print
        TraverseComponent swChildComp, nLevel + 1, vaultPath
    Next
End Sub

'=======================================================================
'þ======== Tao folder STEP

Private Sub CheckSTPFolderExists2()
    AUTOPARTLIST.Hide
    Set EndProgram = New Class_ClickESC
    'Call ErrorHandling2
Dim pathSTP2 As String
Dim strFolderExists2 As String
strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    pathSTP2 = strGetCurrentPath + "\STEP\"
    strFolderExists = Dir(pathSTP2, vbDirectory)
    If strFolderExists = "" Then
        MkDir pathSTP2
        Call cmdEXPORTstp2_Click_1
    Else
        Call cmdEXPORTstp2_Click_1
    End If
End Sub

'=====
'Option Explicit
'    Public swApp        As SldWorks.SldWorks
'    Public swModDoc     As SldWorks.ModelDoc2
'    Dim swView          As SldWorks.View
'    Dim swPart          As PartDoc
'    Dim swAss           As AssemblyDoc
'    Dim boolstatus      As Boolean
'    Dim lErrors         As Long 'Varaible to collect Errors
'    Dim lWarnings       As Long 'Varaible to collect Errors

Private Sub StepFDrw(pathSTP3 As String)
    Dim swView          As SldWorks.View
    Dim swPart          As PartDoc
    Dim swAss           As AssemblyDoc
    Dim boolstatus      As Boolean
    Dim lErrors         As Long 'Varaible to collect Errors
    Dim lWarnings       As Long 'Varaible to collect Errors
    Dim nameS As String
    Dim pathSTP2 As String
    pathSTP2 = pathSTP3
    Set swApp = Application.SldWorks
    Set swModDoc = swApp.ActiveDoc
    Set swView = swModDoc.GetFirstView
    Set swView = swView.GetNextView
    If swView.ReferencedDocument.GetType = 1 Then
        Set swPart = swView.ReferencedDocument
        'boolstatus = swPart.Save3(swSaveAsOptions_Silent, lErrors, lWarnings)
        boolstatus = swPart.SaveAs3(pathSTP2 & ".STEP", 0, 0)
    ElseIf swView.ReferencedDocument.GetType = 2 Then
        Set swAss = swView.ReferencedDocument
        boolstatus = swAss.SaveAs3(pathSTP2 & ".STEP", 0, 0)
    End If
End Sub

Private Sub cmdEXPORTstp2_Click()
Call CheckSTPFolderExists2
End Sub

Private Sub cmdEXPORTstp2_Click_1()
Call CheckRun
    blnStatusExportPDF = True
    AUTOPARTLIST.Hide
    Call ErrorHandling
    Call GetPoCodeAndSoMaySoCum
    Call GetRange(strRangeSTP, CInt(UBound(vSheetNames) + txtStartRenameSheet - 1))
    If flag = False Then
'        If swApp.SendMsgToUser2("Please choose a created folder to save file PDF!", swMbInformation, swMbYesNo) = swMbHitYes Then
            Call EXPORTstp2
            End
'        Else
'            GoTo EndExportPDF
'        End If
    End If
    blnStatusExportStp2 = False
EndExportPDF:
    AUTOPARTLIST.Show
End Sub

Private Sub EXPORTstp2()
Dim Pathdwg As String
    strGetCurrentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\", -1, vbTextCompare))
    PathSTEP2 = strGetCurrentPath + "\STEP\"
    Set EndProgram = New Class_ClickESC
    Set swExportSTPData = swApp.GetExportFileData(1)
    If blnStatusExportExcel = False Then strFilePath = PathSTEP2 'BrowseFolder
    If strFilePath <> "" Then
        If opbtAllToSeparateFileSTP Then
            For intSheetNumber = 0 To UBound(vSheetNames)
                EndProgram.EndByESCKey
                ProcessExportSTP intSheetNumber
            Next intSheetNumber
        ElseIf opbtAllInOneFileSTP Then
            blnExportExcel = swExportSTPData.SetSheets(swExportData_ExportAllSheet, vSheetNames)
            If opbtDrawingNoSTP = True Then
                blnExportExcel = swModel.Extension.SaveAs(strFilePath & "\" & UCase(strPoCode) & "-" & strSoMaySoCum & ".step", 0, 0, swExportSTPData, Empty, Empty)
'                pathSTP3 = strFilePath & "\" & strDrawingNo
'                Call StepFDrw(pathSTP3)
            ElseIf opbtPartnameSTP = True Then
                Call FilterPartNameFromPathName(swModel.GetPathName)
                blnExportExcel = swModel.Extension.SaveAs(strFilePath & "\" & strPartName & ".step", 0, 0, swExportSTPData, Empty, Empty)
'                pathSTP3 = strFilePath & "\" & strPartName
'                Call StepFDrw(pathSTP3)
'
            End If
        ElseIf opbtRangeSTP Then
            For i = 0 To iMemory2 - 1
                EndProgram.EndByESCKey
                intSheetNumber = CInt(intGetRange(i) - txtStartRenameSheet)
'                Call ProcessExportSTP(intSheetNumber)
                ProcessExportSTP intSheetNumber
            Next i
        ElseIf opbtCurrentSheetSTP Then
            For intSheetNumber = 0 To UBound(vSheetNames)
                If vSheetNames(intSheetNumber) = swDraw.GetCurrentSheet.GetName Then Call ProcessExportSTP(intSheetNumber)
            Next intSheetNumber
        End If
        swApp.SendMsgToUser2 "EXPORT FILE STEP COMPLETED!", swMbInformation, swMbOk
    Else
        Exit Sub
    End If
End Sub

Private Sub ProcessExportSTP(intSheetNumber As Integer)
'    Dim pathSTP3 As String
    swDraw.ActivateSheet vSheetNames(intSheetNumber)
    blnExportExcel = swExportSTPData.SetSheets(swExportData_ExportCurrentSheet, vSheetNames)
    If opbtDrawingNoSTP = True Then
        Call InfoDrawingNo
        'blnExportExcel = swModel.Extension.SaveAs(strFilePath & "\" & strDrawingNo & ".step", 0, 0, swExportSTPData, Empty, Empty)
        pathSTP3 = strFilePath & "\" & strDrawingNo
        Call StepFDrw(pathSTP3)
    ElseIf opbtPartnameSTP = True Then
        Set swView = swDraw.GetFirstView
        Set swView = swView.GetNextView
        strPathName = swView.GetReferencedModelName
        Call OpenModel(strPathName)
        Call InfoPartname(swOpenModel)
        Call SaveAndCloseModel3D(strPathName)
        'blnExportExcel = swModel.Extension.SaveAs(strFilePath & "\" & strPartName & ".step", 0, 0, swExportSTPData, Empty, Empty)
        pathSTP3 = strFilePath & "\" & strPartName
        Call StepFDrw(pathSTP3)
    End If
End Sub

Private Sub strRangeSTP_Change()
    opbtRangeSTP.Value = True
End Sub

'========================================================================
' HE THONG DEM VA XUAT DANH SACH BULONG CHUYEN DUNG (HO TRO TOOLBOX VA KHOI FOLDER)
' HO TRO LOAD DAY DU CAC BO LAP CON (SUB-ASSEMBLIES) VA QUET TOAN BO FOLDER CUNG LOAI

'========================================================================

Private Function IsBulongFolderName(ByVal strName As String) As Boolean
    Dim s As String
    s = UCase(Trim(strName))
    Dim sClean As String
    sClean = Replace(s, " ", "")
    sClean = Replace(sClean, "-", "")
    sClean = Replace(sClean, "_", "")
    sClean = Replace(sClean, "&", "")
    If sClean = "BULONG" Or sClean = "BOLT" Or sClean = "BOLTNUT" Or _
       sClean = "BOLTS" Or sClean = "BULON" Or sClean = "FASTENER" Or _
       sClean = "FASTENERS" Or _
       InStr(1, s, "BULONG", vbTextCompare) > 0 Or _
       InStr(1, s, "BOLT", vbTextCompare) > 0 Then
        IsBulongFolderName = True
    Else
        IsBulongFolderName = False
    End If
End Function

Private Function IsBulongChecked() As Boolean
    Dim b As Boolean
    b = False
    On Error Resume Next
    If Not b Then b = (opbtBoltStandardDevice.Value = True)
    If Not b Then b = (chkBulong.Value = True)
    If Not b Then b = (opbtBulong.Value = True)
    If Not b Then b = (chkBolt.Value = True)
    If Not b Then b = (opbtBolt.Value = True)
    If Not b Then b = (CheckBox1.Value = True)
    On Error GoTo 0
    IsBulongChecked = b
End Function

Private Function IsBulongCumChecked() As Boolean
    Dim b As Boolean
    b = False
    On Error Resume Next
    If Not b Then b = (chkBulongCum.Value = True)
    If Not b Then b = (opbtBulongCum.Value = True)
    If Not b Then b = (chkBulong_Cum.Value = True)
    If Not b Then b = (opbtBulong_Cum.Value = True)
    If Not b Then b = (chkBoltSubAsm.Value = True)
    If Not b Then b = (opbtBoltSubAsm.Value = True)
    If Not b Then b = (chkBulongTheoCum.Value = True)
    If Not b Then b = (opbtBulongTheoCum.Value = True)
    If Not b Then b = (CheckBox2.Value = True)
    On Error GoTo 0
    IsBulongCumChecked = b
End Function

Private Sub chkBulongCum_Click()
    On Error Resume Next
    If IsBulongCumChecked() Then
        opbtBoltStandardDevice.Value = True
        chkBulong.Value = True
    End If
    On Error GoTo 0
End Sub

Private Sub opbtBulongCum_Click()
    On Error Resume Next
    If IsBulongCumChecked() Then
        opbtBoltStandardDevice.Value = True
        chkBulong.Value = True
    End If
    On Error GoTo 0
End Sub

Private Sub CheckBox2_Click()
    On Error Resume Next
    If IsBulongCumChecked() Then
        opbtBoltStandardDevice.Value = True
        CheckBox1.Value = True
    End If
    On Error GoTo 0
End Sub

Private Sub ProcessExportBulong()
    Dim swAsmDoc As SldWorks.AssemblyDoc
    Set swAsmDoc = swModel
    If swAsmDoc Is Nothing Then
        MsgBox "Please open an assembly document!", vbCritical
        Exit Sub
    End If

    ' 0. Tu dong load va giai nen toan bo linh kien Lightweight trong mo hinh
    On Error Resume Next
    swAsmDoc.ResolveAllLightWeightComponents False
    On Error GoTo 0

    Dim totalMachines As Long
    totalMachines = 1
    If IsNumeric(strTotalSetAutoQTY) And val(strTotalSetAutoQTY) > 0 Then
        totalMachines = CLng(strTotalSetAutoQTY)
    End If

    Dim blnExportByCluster As Boolean
    blnExportByCluster = IsBulongCumChecked()

    '========================================================================
    ' CH? Ð? 1: CH? CH?N BULONG -> T?NG H?P TOÀN B? THEO C?M L?N
    '========================================================================
    If Not blnExportByCluster Then
        Dim strBoltName() As String
        Dim strBoltPath() As String
        Dim strBoltConfig() As String
        Dim intBoltCount() As Long
        Dim intTotalBoltTypes As Long

        ReDim strBoltName(1000) As String
        ReDim strBoltPath(1000) As String
        ReDim strBoltConfig(1000) As String
        ReDim intBoltCount(1000) As Long
        intTotalBoltTypes = 0

        ' 1. Quet folder BULONG / BOLT ngay tai Assembly chinh (Top-Level)
        Call ScanBulongFolder(swModel, 1, strBoltName, strBoltPath, strBoltConfig, intBoltCount, intTotalBoltTypes)

        ' 2. Nap va quet toan bo cac Bo lap con (Sub-Assemblies) o moi cap
        Dim vAllComps As Variant
        vAllComps = swAsmDoc.GetComponents(False)

        If Not IsEmpty(vAllComps) Then
            Dim strProcessedSubAsm() As String
            Dim intProcessedCount As Long
            ReDim strProcessedSubAsm(UBound(vAllComps)) As String
            intProcessedCount = 0

            Dim iComp As Long
            For iComp = LBound(vAllComps) To UBound(vAllComps)
                Dim swC As SldWorks.Component2
                Set swC = vAllComps(iComp)
                If Not swC Is Nothing Then
                    If swC.GetSuppression <> 0 And UCase(Right(swC.GetPathName, 6)) = "SLDASM" Then
                        Dim subPath As String
                        subPath = swC.GetPathName

                        Dim bProcessed As Boolean
                        bProcessed = False
                        Dim p As Long
                        For p = 0 To intProcessedCount - 1
                            If StrComp(strProcessedSubAsm(p), subPath, vbTextCompare) = 0 Then
                                bProcessed = True
                                Exit For
                            End If
                        Next p

                        If Not bProcessed Then
                            strProcessedSubAsm(intProcessedCount) = subPath
                            intProcessedCount = intProcessedCount + 1

                            Dim subCount As Long
                            subCount = 0
                            Dim q As Long
                            For q = LBound(vAllComps) To UBound(vAllComps)
                                If StrComp(vAllComps(q).GetPathName, subPath, vbTextCompare) = 0 And vAllComps(q).GetSuppression <> 0 Then
                                    subCount = subCount + 1
                                End If
                            Next q

                            Dim swSubModel As SldWorks.ModelDoc2
                            Set swSubModel = swC.GetModelDoc2
                            If swSubModel Is Nothing Then
                                Dim lErrors As Long, lWarnings As Long
                                Set swSubModel = swApp.OpenDoc6(subPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", lErrors, lWarnings)
                            End If

                            If Not swSubModel Is Nothing Then
                                If swSubModel.GetType = swDocASSEMBLY Then
                                    Dim swSubAsmDoc As SldWorks.AssemblyDoc
                                    Set swSubAsmDoc = swSubModel
                                    On Error Resume Next
                                    swSubAsmDoc.ResolveAllLightWeightComponents False
                                    On Error GoTo 0
                                End If
                                Call ScanBulongFolder(swSubModel, subCount, strBoltName, strBoltPath, strBoltConfig, intBoltCount, intTotalBoltTypes)
                            End If
                        End If
                    End If
                End If
            Next iComp
        End If

        If intTotalBoltTypes = 0 Then
            swApp.SendMsgToUser2 "Khong tim thay Folder BULONG / BOLT hoac khong co chi tiet nao ben trong cac folder nay!", swMbInformation, swMbOk
            Exit Sub
        End If

        ' Xuat 1 bang duy nhat tong hop theo cum lon
        Dim xlApp As Object
        Dim xlWorkbook As Object
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        xlApp.WindowState = xlMaximized
        Set xlWorkbook = xlApp.Workbooks.Add

        With xlWorkbook.Worksheets(1)
            On Error Resume Next
            .Columns("B").NumberFormat = "@"

            .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH BULONG"
            .Range(.Cells(1, 1), .Cells(1, 3)).Merge

            Call FilterPartNameFromPathName(swModel.GetPathName)
            .Cells(2, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            .Range(.Cells(2, 1), .Cells(2, 3)).Merge
            .Range(.Cells(1, 1), .Cells(3, 3)).Font.Size = 10

            .Cells(3, 1).Value = "STT"
            .Cells(3, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
            .Cells(3, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            .Range("A1:C3").Font.Bold = True

            Dim j As Long
            For j = 0 To intTotalBoltTypes - 1
                .Cells(j + 4, 1).Value = j + 1
                If strBoltPath(j) <> "" Then
                    .Hyperlinks.Add .Cells(j + 4, 2), strBoltPath(j), TextToDisplay:=strBoltName(j)
                Else
                    .Cells(j + 4, 2).Value = strBoltName(j)
                End If
                .Cells(j + 4, 3).Value = CStr(intBoltCount(j) * totalMachines)
            Next j

            .Columns("A:C").AutoFit
            .Columns("A:C").HorizontalAlignment = xlCenter
            .Range(.Cells(1, 1), .Cells(intTotalBoltTypes + 3, 3)).Borders.LineStyle = xlContinuous
        End With
        Exit Sub
    End If

    '========================================================================
    ' CH? Ð? 2: CH?N C? BULONG VÀ BULONG C?M -> LIST RIÊNG THEO T?NG C?M
    '========================================================================
    Dim xlApp2 As Object
    Dim xlWorkbook2 As Object
    Set xlApp2 = CreateObject("Excel.Application")
    xlApp2.Visible = True
    xlApp2.WindowState = xlMaximized
    Set xlWorkbook2 = xlApp2.Workbooks.Add

    Dim hasAnyClusterBolts As Boolean
    hasAnyClusterBolts = False

    With xlWorkbook2.Worksheets(1)
        On Error Resume Next
        .Columns("B").NumberFormat = "@"

        ' Tieu de chung toan bang
        .Cells(1, 1).Value = "DANH S" & ChrW(193) & "CH BULONG"
        .Range(.Cells(1, 1), .Cells(1, 3)).Merge
        .Range(.Cells(1, 1), .Cells(1, 3)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 3)).Font.Size = 10
        .Range(.Cells(1, 1), .Cells(1, 3)).HorizontalAlignment = xlCenter

        Dim curRow As Long
        curRow = 2

        ' 1. Kiem tra Bulong tai Cum chinh (Top-Level) neu co
        Dim strClustBoltName() As String
        Dim strClustBoltPath() As String
        Dim strClustBoltConfig() As String
        Dim intClustBoltCount() As Long
        Dim intClustTotal As Long

        ReDim strClustBoltName(500) As String
        ReDim strClustBoltPath(500) As String
        ReDim strClustBoltConfig(500) As String
        ReDim intClustBoltCount(500) As Long
        intClustTotal = 0

        Call ScanBulongFolder(swModel, 1, strClustBoltName, strClustBoltPath, strClustBoltConfig, intClustBoltCount, intClustTotal)

        If intClustTotal > 0 Then
            hasAnyClusterBolts = True
            Call FilterPartNameFromPathName(swModel.GetPathName)

            .Cells(curRow, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
            .Range(.Cells(curRow, 1), .Cells(curRow, 3)).Merge
            .Range(.Cells(curRow, 1), .Cells(curRow, 3)).Font.Bold = True
            .Range(.Cells(curRow, 1), .Cells(curRow, 3)).Font.Size = 10
            .Range(.Cells(curRow, 1), .Cells(curRow, 3)).HorizontalAlignment = xlCenter

            .Cells(curRow + 1, 1).Value = "STT"
            .Cells(curRow + 1, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
            .Cells(curRow + 1, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
            .Range(.Cells(curRow + 1, 1), .Cells(curRow + 1, 3)).Font.Bold = True
            .Range(.Cells(curRow + 1, 1), .Cells(curRow + 1, 3)).HorizontalAlignment = xlCenter

            Dim rowIdx As Long
            For rowIdx = 0 To intClustTotal - 1
                Dim dRow As Long
                dRow = curRow + 2 + rowIdx
                .Cells(dRow, 1).Value = rowIdx + 1
                If strClustBoltPath(rowIdx) <> "" Then
                    .Hyperlinks.Add .Cells(dRow, 2), strClustBoltPath(rowIdx), TextToDisplay:=strClustBoltName(rowIdx)
                Else
                    .Cells(dRow, 2).Value = strClustBoltName(rowIdx)
                End If
                .Cells(dRow, 3).Value = CStr(intClustBoltCount(rowIdx) * totalMachines)
            Next rowIdx

            .Range(.Cells(curRow, 1), .Cells(curRow + 1 + intClustTotal, 3)).Borders.LineStyle = xlContinuous
            curRow = curRow + 2 + intClustTotal + 1
        End If

        ' 2. Kiem tra lan luot tung Cum con (Sub-Assemblies)
        Dim vSubComps As Variant
        vSubComps = swAsmDoc.GetComponents(False)

        If Not IsEmpty(vSubComps) Then
            Dim strDoneSub() As String
            Dim intDoneCount As Long
            ReDim strDoneSub(UBound(vSubComps)) As String
            intDoneCount = 0

            Dim sIdx As Long
            For sIdx = LBound(vSubComps) To UBound(vSubComps)
                Dim swSubC As SldWorks.Component2
                Set swSubC = vSubComps(sIdx)
                If Not swSubC Is Nothing Then
                    If swSubC.GetSuppression <> 0 And UCase(Right(swSubC.GetPathName, 6)) = "SLDASM" Then
                        Dim sPath As String
                        sPath = swSubC.GetPathName

                        Dim isDone As Boolean
                        isDone = False
                        Dim d As Long
                        For d = 0 To intDoneCount - 1
                            If StrComp(strDoneSub(d), sPath, vbTextCompare) = 0 Then
                                isDone = True
                                Exit For
                            End If
                        Next d

                        If Not isDone Then
                            strDoneSub(intDoneCount) = sPath
                            intDoneCount = intDoneCount + 1

                            Dim swSubDoc As SldWorks.ModelDoc2
                            Set swSubDoc = swSubC.GetModelDoc2
                            If swSubDoc Is Nothing Then
                                Dim lErr2 As Long, lWarn2 As Long
                                Set swSubDoc = swApp.OpenDoc6(sPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", lErr2, lWarn2)
                            End If

                            If Not swSubDoc Is Nothing Then
                                If swSubDoc.GetType = swDocASSEMBLY Then
                                    Dim swSubAsmDoc2 As SldWorks.AssemblyDoc
                                    Set swSubAsmDoc2 = swSubDoc
                                    On Error Resume Next
                                    swSubAsmDoc2.ResolveAllLightWeightComponents False
                                    On Error GoTo 0
                                End If

                                ReDim strClustBoltName(500) As String
                                ReDim strClustBoltPath(500) As String
                                ReDim strClustBoltConfig(500) As String
                                ReDim intClustBoltCount(500) As Long
                                intClustTotal = 0

                                Call ScanBulongFolder(swSubDoc, 1, strClustBoltName, strClustBoltPath, strClustBoltConfig, intClustBoltCount, intClustTotal)

                                If intClustTotal > 0 Then
                                    hasAnyClusterBolts = True
                                    Call FilterPartNameFromPathName(sPath)

                                    .Cells(curRow, 1).Value = "C" & ChrW(7908) & "M: " & strPartName
                                    .Range(.Cells(curRow, 1), .Cells(curRow, 3)).Merge
                                    .Range(.Cells(curRow, 1), .Cells(curRow, 3)).Font.Bold = True
                                    .Range(.Cells(curRow, 1), .Cells(curRow, 3)).Font.Size = 10
                                    .Range(.Cells(curRow, 1), .Cells(curRow, 3)).HorizontalAlignment = xlCenter

                                    .Cells(curRow + 1, 1).Value = "STT"
                                    .Cells(curRow + 1, 2).Value = "T" & ChrW(202) & "N THI" & ChrW(7870) & "T B" & ChrW(7882)
                                    .Cells(curRow + 1, 3).Value = "S" & ChrW(7888) & " L" & ChrW(431) & ChrW(7906) & "NG TI" & ChrW(769) & "NH TO" & ChrW(193) & "N"
                                    .Range(.Cells(curRow + 1, 1), .Cells(curRow + 1, 3)).Font.Bold = True
                                    .Range(.Cells(curRow + 1, 1), .Cells(curRow + 1, 3)).HorizontalAlignment = xlCenter

                                    Dim r As Long
                                    For r = 0 To intClustTotal - 1
                                        Dim dataRow2 As Long
                                        dataRow2 = curRow + 2 + r
                                        .Cells(dataRow2, 1).Value = r + 1
                                        If strClustBoltPath(r) <> "" Then
                                            .Hyperlinks.Add .Cells(dataRow2, 2), strClustBoltPath(r), TextToDisplay:=strClustBoltName(r)
                                        Else
                                            .Cells(dataRow2, 2).Value = strClustBoltName(r)
                                        End If
                                        .Cells(dataRow2, 3).Value = CStr(intClustBoltCount(r) * totalMachines)
                                    Next r

                                    .Range(.Cells(curRow, 1), .Cells(curRow + 1 + intClustTotal, 3)).Borders.LineStyle = xlContinuous
                                    curRow = curRow + 2 + intClustTotal + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Next sIdx
        End If

        If Not hasAnyClusterBolts Then
            xlWorkbook2.Close False
            xlApp2.Quit
            swApp.SendMsgToUser2 "Khong tim thay Folder BULONG / BOLT nao trong cac cum lap!", swMbInformation, swMbOk
            Exit Sub
        End If

        .Columns("A:C").AutoFit
        .Columns("A:C").HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub ScanBulongFolder(swTargetModel As SldWorks.ModelDoc2, intMultiplier As Long, _
                             ByRef strBoltName() As String, ByRef strBoltPath() As String, _
                             ByRef strBoltConfig() As String, ByRef intBoltCount() As Long, _
                             ByRef intTotalBoltTypes As Long)
    If swTargetModel Is Nothing Then Exit Sub
    Dim swFeat As SldWorks.Feature
    Dim swFeatFolder As SldWorks.FeatureFolder
    Dim vFeatures As Variant
    Dim i As Long

    Set swFeat = swTargetModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "FtrFolder" And InStr(1, swFeat.name, "EndTag", vbTextCompare) = 0 Then
            If IsBulongFolderName(swFeat.name) Then
                Set swFeatFolder = swFeat.GetSpecificFeature2
                If Not swFeatFolder Is Nothing Then
                    vFeatures = swFeatFolder.GetFeatures
                    If Not IsEmpty(vFeatures) Then
                        For i = 0 To (swFeatFolder.GetFeatureCount - 1)
                            Dim swSubFeat As SldWorks.Feature
                            Set swSubFeat = vFeatures(i)
                            If Not swSubFeat Is Nothing Then
                                Call ProcessBulongItem(swTargetModel, swSubFeat, intMultiplier, _
                                                       strBoltName, strBoltPath, strBoltConfig, _
                                                       intBoltCount, intTotalBoltTypes)
                            End If
                        Next i
                    End If
                End If
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
End Sub

Private Sub ProcessBulongItem(swTargetModel As SldWorks.ModelDoc2, swSubFeat As SldWorks.Feature, intMultiplier As Long, _
                             ByRef strBoltName() As String, ByRef strBoltPath() As String, _
                             ByRef strBoltConfig() As String, ByRef intBoltCount() As Long, _
                             ByRef intTotalBoltTypes As Long)
    Dim swComp As SldWorks.Component2
    Set swComp = Nothing

    On Error Resume Next
    Set swComp = swSubFeat.GetSpecificFeature2

    If swComp Is Nothing Then
        Dim swEntity As SldWorks.Entity
        Set swEntity = swSubFeat
        If Not swEntity Is Nothing Then
            Set swComp = swEntity.GetComponent
        End If
    End If

    If swComp Is Nothing And swTargetModel.GetType = swDocASSEMBLY Then
        Dim swAsmDocTarget As SldWorks.AssemblyDoc
        Set swAsmDocTarget = swTargetModel
        Set swComp = swAsmDocTarget.GetComponentByName(swSubFeat.name)
        If swComp Is Nothing Then
            Dim strCleanName As String
            strCleanName = Replace(swSubFeat.name, "<", "-")
            strCleanName = Replace(strCleanName, ">", "")
            Set swComp = swAsmDocTarget.GetComponentByName(strCleanName)
        End If
    End If

    Dim strPath As String
    Dim strConfig As String
    Dim strPart As String
    Dim strDisplay As String

    If Not swComp Is Nothing Then
        If swComp.GetSuppression = 0 Then Exit Sub
        strPath = swComp.GetPathName
        strConfig = swComp.ReferencedConfiguration
        Call FilterPartNameFromPathName(strPath)
        strPart = strPartName

        If strConfig <> "" And InStr(1, strConfig, "Default", vbTextCompare) = 0 Then
            If InStr(1, strConfig, strPart, vbTextCompare) = 0 And InStr(1, UCase(strPath), "SOLIDWORKS DATA") = 0 Then
                strDisplay = strPart & " - " & strConfig
            Else
                strDisplay = strConfig
            End If
        Else
            strDisplay = strPart
            strPartNameTopLevel_Last = strDisplay
            Call EditPartNameStandardDevice
            strDisplay = strPartNameTopLevel_Last
        End If
    Else
        strPath = ""
        strConfig = ""
        strDisplay = swSubFeat.name
        If InStrRev(strDisplay, "-", -1, vbTextCompare) <> 0 Then
            strDisplay = Left(strDisplay, InStrRev(strDisplay, "-", -1, vbTextCompare) - 1)
        ElseIf InStrRev(strDisplay, "<", -1, vbTextCompare) <> 0 Then
            strDisplay = Left(strDisplay, InStrRev(strDisplay, "<", -1, vbTextCompare) - 1)
        End If
    End If

    If Trim(strDisplay) = "" Then Exit Sub

    Dim foundIndex As Long
    foundIndex = -1
    Dim k As Long
    For k = 0 To intTotalBoltTypes - 1
        If StrComp(strBoltName(k), strDisplay, vbTextCompare) = 0 Then
            foundIndex = k
            Exit For
        End If
    Next k

    If foundIndex >= 0 Then
        intBoltCount(foundIndex) = intBoltCount(foundIndex) + intMultiplier
    Else
        If intTotalBoltTypes > UBound(strBoltName) Then
            ReDim Preserve strBoltName(UBound(strBoltName) + 500) As String
            ReDim Preserve strBoltPath(UBound(strBoltPath) + 500) As String
            ReDim Preserve strBoltConfig(UBound(strBoltConfig) + 500) As String
            ReDim Preserve intBoltCount(UBound(intBoltCount) + 500) As Long
        End If
        strBoltName(intTotalBoltTypes) = strDisplay
        strBoltPath(intTotalBoltTypes) = strPath
        strBoltConfig(intTotalBoltTypes) = strConfig
        intBoltCount(intTotalBoltTypes) = intMultiplier
        intTotalBoltTypes = intTotalBoltTypes + 1
    End If
End Sub


'=========================================================================
'
' QR Code generator library (VB6/VBA)
'
' Copyright (c) Project Nayuki. (MIT License)
' https://www.nayuki.io/page/qr-code-generator-library
'
' Copyright (c) wqweto@gmail.com (MIT License)
'
'=========================================================================
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0) Or (TWINBASIC <> 0)

'=========================================================================
' Public enums
'=========================================================================

Public Enum QRCodegenEcc
    QRCodegenEcc_LOW = 0  ' The QR Code can tolerate about  7% erroneous codewords
    QRCodegenEcc_MEDIUM   ' The QR Code can tolerate about 15% erroneous codewords
    QRCodegenEcc_QUARTILE ' The QR Code can tolerate about 25% erroneous codewords
    QRCodegenEcc_HIGH     ' The QR Code can tolerate about 30% erroneous codewords
End Enum

Public Enum QRCodegenMask
    QRCodegenMask_AUTO = -1
    QRCodegenMask_0 = 0
    QRCodegenMask_1
    QRCodegenMask_2
    QRCodegenMask_3
    QRCodegenMask_4
    QRCodegenMask_5
    QRCodegenMask_6
    QRCodegenMask_7
End Enum

Public Enum QRCodegenMode
    QRCodegenMode_NUMERIC = &H1
    QRCodegenMode_ALPHANUMERIC = &H2
    QRCodegenMode_BYTE = &H4
    QRCodegenMode_KANJI = &H8
    QRCodegenMode_ECI = &H7
End Enum

Public Type QRCodegenSegment
    Mode            As QRCodegenMode
    NumChars        As Long
    Data()          As Byte
    BitLength       As Long
End Type

'=========================================================================
' API
'=========================================================================

#If HasPtrSafe Then
Private Declare PtrSafe Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileW" (ByVal hdcRef As LongPtr, ByVal lpFileName As LongPtr, ByVal lpRect As LongPtr, ByVal lpDescription As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, lpvObj As IPicture) As Long
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As LongPtr, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function SHCreateMemStream Lib "shlwapi" Alias "#12" (pInit As Any, ByVal cbInit As Long) As stdole.IUnknown
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, lpBitsInfo As Any, ByVal wUsage As Long, lpBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare PtrSafe Function SetStretchBltMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nStretchMode As Long) As Long
Private Declare PtrSafe Function StretchBlt Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function PolyPolygon Lib "gdi32" (ByVal hDC As LongPtr, lpPoint As Any, lpPolyCounts As Any, ByVal nCount As Long) As Long
Private Declare PtrSafe Function SetMapMode Lib "gdi32" (ByVal hCD As LongPtr, ByVal nMapMode As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileW" (ByVal hdcRef As LongPtr, ByVal lpFileName As LongPtr, ByVal lpRect As LongPtr, ByVal lpDescription As LongPtr) As LongPtr
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, lpvObj As IPicture) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As LongPtr, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare Function SHCreateMemStream Lib "shlwapi" Alias "#12" (pInit As Any, ByVal cbInit As Long) As stdole.IUnknown
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, lpBitsInfo As Any, ByVal wUsage As Long, lpBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare Function PolyPolygon Lib "gdi32" (ByVal hDC As LongPtr, lpPoint As Any, lpPolyCounts As Any, ByVal nCount As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hCD As LongPtr, ByVal nMapMode As Long) As Long
#End If

Private Type POINTAPI
    x                   As Long
    Y                   As Long
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type PICTDESC
    Size                As Long
    Type                As Long
    hBmpOrIcon          As LongPtr
    hPal                As LongPtr
End Type

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const VERSION_MIN               As Long = 1
Private Const VERSION_MAX               As Long = 40
Private Const PENALTY_N1                As Long = 3
Private Const PENALTY_N2                As Long = 3
Private Const PENALTY_N3                As Long = 40
Private Const PENALTY_N4                As Long = 10
Private Const INT16_MAX                 As Long = 32767
Private Const LONG_MAX                  As Long = 2147483647
Private Const ALPHANUMERIC_CHARSET      As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:"
Private Const M_PI_2                    As Double = 3.14159265358979 / 2
Private Const LINE_TO                   As Long = 1
Private Const TURN_LEFT                 As Long = 2
Private Const TURN_RIGHT                As Long = 3
Private Const WIDE_LEFT                 As Long = 4
Private Const WIDE_RIGHT                As Long = 5
Private Const SUPERWIDE_LEFT            As Long = 6 '--- not used
Private Const SUPERWIDE_RIGHT           As Long = 7
Private Const SKIP_TO                   As Long = 8

Private LNG_POW2(0 To 31)                            As Long
Private ECC_CODEWORDS_PER_BLOCK(0 To 3, 0 To 40)     As Long
Private NUM_ERROR_CORRECTION_BLOCKS(0 To 3, 0 To 40) As Long
Private IID_IPicture(0 To 3)                         As Long '--- 7BF80980-BF32-101A-8BBB-00AA00300CAB
Private IID_IPersistStream(0 To 3)                   As Long '--- 00000109-0000-0000-C000-000000000046

'=========================================================================
' Functions
'=========================================================================

Public Function QRCodegenBarcode(TextOrByteArray As Variant, _
            Optional ByVal ForeColor As OLE_COLOR = vbBlack, _
            Optional ByVal ModuleSize As Long = 120, _
            Optional ByVal SquareModules As Boolean, _
            Optional ByVal Ecl As QRCodegenEcc = QRCodegenEcc_LOW, _
            Optional ByVal MinVersion As Long = VERSION_MIN, _
            Optional ByVal MaxVersion As Long = VERSION_MAX, _
            Optional ByVal Mask As QRCodegenMask = QRCodegenMask_AUTO, _
            Optional ByVal BoostEcl As Boolean = True) As StdPicture
    Dim baQrCode()      As Byte
    
    If QRCodegenEncode(TextOrByteArray, baQrCode, Ecl, MinVersion, MaxVersion, Mask, BoostEcl) Then
        Set QRCodegenBarcode = QRCodegenConvertToPicture(baQrCode, ForeColor, ModuleSize, SquareModules)
    End If
End Function

Public Function QRCodegenEncode(TextOrByteArray As Variant, baQrCode() As Byte, _
            Optional ByVal Ecl As QRCodegenEcc = QRCodegenEcc_LOW, _
            Optional ByVal MinVersion As Long = VERSION_MIN, _
            Optional ByVal MaxVersion As Long = VERSION_MAX, _
            Optional ByVal Mask As QRCodegenMask = QRCodegenMask_AUTO, _
            Optional ByVal BoostEcl As Boolean = True) As Boolean
    Dim baData()        As Byte
    Dim lDataLen        As Long
    Dim sText           As String
    Dim lBufLen         As Long
    Dim uSegments()     As QRCodegenSegment
    
    pvInit
    If IsArray(TextOrByteArray) Then
        baData = TextOrByteArray
        lDataLen = UBound(baData) + 1
    Else
        sText = TextOrByteArray
        lDataLen = Len(sText)
    End If
    If lDataLen = 0 Then
        ReDim uSegments(-1 To -1) As QRCodegenSegment
    Else
        ReDim uSegments(0 To 0) As QRCodegenSegment
        lBufLen = pvGetBufferLenForVersion(MaxVersion)
        If IsArray(TextOrByteArray) Then
            If QRCodegenCalcSegmentBufferSize(QRCodegenMode_BYTE, lDataLen) > lBufLen Then
                GoTo QH
            End If
            uSegments(0) = QRCodegenMakeBytes(baData)
        ElseIf QRCodegenIsNumeric(sText) Then
            If QRCodegenCalcSegmentBufferSize(QRCodegenMode_NUMERIC, lDataLen) > lBufLen Then
                GoTo QH
            End If
            uSegments(0) = QRCodegenMakeNumeric(sText)
        ElseIf QRCodegenIsAlphanumeric(sText) Then
            If QRCodegenCalcSegmentBufferSize(QRCodegenMode_ALPHANUMERIC, lDataLen) > lBufLen Then
                GoTo QH
            End If
            uSegments(0) = QRCodegenMakeAlphanumeric(sText)
        Else
            baData = pvToUtf8Array(sText)
            lDataLen = UBound(baData) + 1
            If QRCodegenCalcSegmentBufferSize(QRCodegenMode_BYTE, lDataLen) > lBufLen Then
                GoTo QH
            End If
            uSegments(0) = QRCodegenMakeBytes(baData)
        End If
    End If
    QRCodegenEncode = QRCodegenEncodeSegments(uSegments, baQrCode, Ecl, MinVersion, MaxVersion, Mask, BoostEcl)
    Exit Function
QH:
    ReDim baQrCode(0 To 0) As Byte
End Function

Public Function QRCodegenEncodeSegments(uSegments() As QRCodegenSegment, baQrCode() As Byte, _
            Optional ByVal Ecl As QRCodegenEcc = QRCodegenEcc_LOW, _
            Optional ByVal MinVersion As Long = VERSION_MIN, _
            Optional ByVal MaxVersion As Long = VERSION_MAX, _
            Optional ByVal Mask As QRCodegenMask = QRCodegenMask_AUTO, _
            Optional ByVal BoostEcl As Boolean = True) As Boolean
    Dim lVersion        As Long
    Dim lDataUsedBits   As Long
    Dim lDataCapacityBits As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lBitLen         As Long
    Dim lBit            As Long
    Dim lTerminatorBits As Long
    Dim lPadByte        As Long
    Dim baTemp()        As Byte
    Dim lMinPenalty     As Long
    Dim lPenalty        As Long
    
    pvInit
    '--- Find the minimal version number to use
    For lVersion = MinVersion To MaxVersion
        lDataCapacityBits = pvGetNumDataCodewords(lVersion, Ecl) * 8
        lDataUsedBits = pvGetTotalBits(uSegments, lVersion)
        If lDataUsedBits <> -1 And lDataUsedBits <= lDataCapacityBits Then
            Exit For
        End If
        If lVersion >= MaxVersion Then
            baQrCode = vbNullString
            GoTo QH
        End If
    Next
    Debug.Assert lDataUsedBits <> -1
    '--- Increase the error correction level while the data still fits in the current version number
    If BoostEcl Then
        For lIdx = QRCodegenEcc_MEDIUM To QRCodegenEcc_HIGH
            If lDataUsedBits <= pvGetNumDataCodewords(lVersion, lIdx) * 8 Then
                Ecl = lIdx
            End If
        Next
    End If
    '--- Concatenate all segments to create the data bit string
    ReDim baQrCode(0 To pvGetBufferLenForVersion(lVersion) - 1) As Byte
    For lIdx = 0 To UBound(uSegments)
        With uSegments(lIdx)
            pvAppendBitsToBuffer .Mode, 4, baQrCode, lBitLen
            pvAppendBitsToBuffer .NumChars, pvNumCharCountBits(.Mode, lVersion), baQrCode, lBitLen
            For lJdx = 0 To .BitLength - 1
                lBit = -((.Data(lJdx \ 8) And LNG_POW2(7 - (lJdx And 7))) <> 0)
                pvAppendBitsToBuffer lBit, 1, baQrCode, lBitLen
            Next
        End With
    Next
    Debug.Assert lBitLen = lDataUsedBits
    '--- Add terminator and pad up to a byte if applicable
    lDataCapacityBits = pvGetNumDataCodewords(lVersion, Ecl) * 8
    Debug.Assert lBitLen <= lDataCapacityBits
    lTerminatorBits = lDataCapacityBits - lBitLen
    If lTerminatorBits > 4 Then
        lTerminatorBits = 4
    End If
    pvAppendBitsToBuffer 0, lTerminatorBits, baQrCode, lBitLen
    pvAppendBitsToBuffer 0, (8 - (lBitLen And 7)) And 7, baQrCode, lBitLen
    Debug.Assert lBitLen Mod 8 = 0
    '--- Pad with alternating bytes until data capacity is reached
    lPadByte = &HEC
    Do While lBitLen < lDataCapacityBits
        pvAppendBitsToBuffer lPadByte, 8, baQrCode, lBitLen
        lPadByte = lPadByte Xor (&HEC Xor &H11)
    Loop
    '--- Compute ECC, draw modules
    pvAddEccAndInterleave baQrCode, lVersion, Ecl, baTemp
    pvInitializeFunctionModules lVersion, baQrCode
    Debug.Assert UBound(baTemp) + 1 = pvGetNumRawDataModules(lVersion) \ 8
    pvDrawCodewords baTemp, baQrCode
    pvDrawLightFunctionModules lVersion, baQrCode
    pvInitializeFunctionModules lVersion, baTemp
    '--- Do masking
    If Mask = QRCodegenMask_AUTO Then
        lMinPenalty = LONG_MAX
        For lIdx = QRCodegenMask_0 To QRCodegenMask_7
            pvApplyMask baTemp, baQrCode, lIdx
            pvDrawFormatBits Ecl, lIdx, baQrCode
            lPenalty = pvGetPenaltyScore(baQrCode)
            If lPenalty < lMinPenalty Then
                Mask = lIdx
                lMinPenalty = lPenalty
            End If
            pvApplyMask baTemp, baQrCode, lIdx '--- Undoes the mask due to XOR
        Next
    End If
    Debug.Assert QRCodegenMask_0 <= Mask And Mask <= QRCodegenMask_7
    pvApplyMask baTemp, baQrCode, Mask
    pvDrawFormatBits Ecl, Mask, baQrCode
    '--- success
    QRCodegenEncodeSegments = True
QH:
End Function

Public Function QRCodegenConvertToPicture(baQrCode() As Byte, _
            Optional ByVal ForeColor As OLE_COLOR = vbBlack, _
            Optional ByVal ModuleSize As Long = 120, _
            Optional ByVal SquareModules As Boolean) As StdPicture
    Const WHITE_BRUSH   As Long = 0
    Const NULL_PEN      As Long = 8
    Const MM_TWIPS      As Long = 6
    Const vbPicTypeEMetafile As Long = 4
    Dim uVectors()      As RECT
    Dim uPoints()       As POINTAPI
    Dim aSizes()        As Long
    Dim lQrSize         As Long
    Dim hDC             As LongPtr
    Dim uRect           As RECT
    Dim hBlackBrush     As LongPtr
    Dim hPrevBrush      As LongPtr
    Dim hPrevPen        As LongPtr
    Dim lIdx            As Long
    Dim uDesc           As PICTDESC
    Dim hResult         As Long
    Dim vErr            As Variant
    
    On Error GoTo EH
    pvConstructVectors baQrCode, SquareModules, uVectors
    pvConstructPolygons uVectors, ModuleSize, uPoints, aSizes
    '--- draw polygons to enhanced metafile
    lQrSize = QRCodegenGetSize(baQrCode)
    hDC = CreateEnhMetaFile(0, 0, 0, 0)
    Call SetMapMode(hDC, MM_TWIPS)
    uRect.Right = (lQrSize + 1) * ModuleSize
    uRect.Bottom = (lQrSize + 1) * ModuleSize
    Call FillRect(hDC, uRect, GetStockObject(WHITE_BRUSH))
    hBlackBrush = CreateSolidBrush(ForeColor)
    hPrevBrush = SelectObject(hDC, hBlackBrush)
    hPrevPen = SelectObject(hDC, GetStockObject(NULL_PEN))
    For lIdx = 0 To UBound(uPoints)
        With uPoints(lIdx)
            .x = ModuleSize \ 2 + .x - 1
            .Y = ModuleSize \ 2 + (lQrSize) * ModuleSize - .Y - 1
        End With
    Next
    Call PolyPolygon(hDC, uPoints(0), aSizes(0), UBound(aSizes) + 1)
    If hPrevBrush <> 0 Then
        Call SelectObject(hDC, hPrevBrush)
        hPrevBrush = 0
    End If
    If hPrevPen <> 0 Then
        Call SelectObject(hDC, hPrevPen)
        hPrevPen = 0
    End If
    With uDesc
        .Size = LenB(uDesc)
        .Type = vbPicTypeEMetafile
        .hBmpOrIcon = CloseEnhMetaFile(hDC)
    End With
    If IID_IPicture(0) = 0 Then
        IID_IPicture(0) = &H7BF80980: IID_IPicture(1) = &H101ABF32: IID_IPicture(2) = &HAA00BB8B: IID_IPicture(3) = &HAB0C3000
    End If
    hResult = OleCreatePictureIndirect(uDesc, IID_IPicture(0), 1, QRCodegenConvertToPicture)
    If hResult < 0 Then
        Err.Raise hResult, "OleCreatePictureIndirect"
    End If
    '--- not to destroy metafile, it's already owned by returned IPicture
    hDC = 0
QH:
    If hPrevBrush <> 0 Then
        Call SelectObject(hDC, hPrevBrush)
        hPrevBrush = 0
    End If
    If hPrevPen <> 0 Then
        Call SelectObject(hDC, hPrevPen)
        hPrevPen = 0
    End If
    If hDC <> 0 Then
        Call CloseEnhMetaFile(hDC)
        hDC = 0
    End If
    If hBlackBrush <> 0 Then
        Call DeleteObject(hBlackBrush)
        hBlackBrush = 0
    End If
    If IsArray(vErr) Then
        On Error GoTo 0
        Err.Raise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    Resume QH
End Function

Public Function QRCodegenConvertToData(ByVal pPicture As IPicture, Optional ByVal NewWidth As Long, Optional ByVal NewHeight As Long) As Byte()
    Const IDX_QUERYINTERFACE As Long = 0
    Const IDX_SAVE          As Long = 6
    Const IDX_SEEK          As Long = 5
    Const IDX_READ          As Long = 3
    Const STREAM_SEEK_SET   As Long = 0
    Const STREAM_SEEK_END   As Long = 2
    Dim pStream         As stdole.IUnknown
    Dim pPersist        As stdole.IUnknown
    Dim cSize           As Currency
    Dim baOutput()      As Byte
    Dim hResult         As Long
    
    If pPicture Is Nothing Then
        baOutput = vbNullString
        GoTo QH
    End If
    If NewWidth > 0 And NewHeight > 0 Then
        '--- super sample to 4x4 for cheap anti-aliasing
        Set pPicture = QRCodegenResizePicture(pPicture, NewWidth * 4, NewHeight * 4)
        Set pPicture = QRCodegenResizePicture(pPicture, NewWidth, NewHeight)
    End If
    Set pStream = SHCreateMemStream(ByVal 0, 0)
    If IID_IPersistStream(0) = 0 Then
        IID_IPersistStream(0) = &H109: IID_IPersistStream(2) = &HC0: IID_IPersistStream(3) = &H46000000
    End If
    hResult = DispCallByVtbl(pPicture, IDX_QUERYINTERFACE, VarPtr(IID_IPersistStream(0)), VarPtr(pPersist))
    If hResult < 0 Then
        Err.Raise hResult, "IUnknown.QueryInterface"
    End If
    hResult = DispCallByVtbl(pPersist, IDX_SAVE, pStream, True)
    If hResult < 0 Then
        Err.Raise hResult, "IPersistStream.Save"
    End If
    hResult = DispCallByVtbl(pStream, IDX_SEEK, 0@, STREAM_SEEK_END, VarPtr(cSize))
    If hResult < 0 Then
        Err.Raise hResult, "IStream.Seek(STREAM_SEEK_END)"
    End If
    If cSize <= 8 Then
        baOutput = vbNullString
        GoTo QH
    End If
    ReDim baOutput(0 To cSize * 10000@ - 9) As Byte
    hResult = DispCallByVtbl(pStream, IDX_SEEK, 0.0008@, STREAM_SEEK_SET, VarPtr(cSize))
    If hResult < 0 Then
        Err.Raise hResult, "IStream.Seek(STREAM_SEEK_SET)"
    End If
    hResult = DispCallByVtbl(pStream, IDX_READ, VarPtr(baOutput(0)), UBound(baOutput) + 1, VarPtr(cSize))
    If hResult < 0 Then
        Err.Raise hResult, "IStream.Read"
    End If
QH:
    QRCodegenConvertToData = baOutput
End Function

Private Function HM2Pix(ByVal Value As Double, ByVal lDpi As Long) As Long
    HM2Pix = Int(Value * lDpi / 2540 + 0.5)
End Function

Public Function QRCodegenResizePicture(pPicture As IPicture, ByVal NewWidth As Long, ByVal NewHeight As Long) As StdPicture
    Const HALFTONE        As Long = 4
    Const DIB_RGB_COLORS  As Long = 0
    Const LOGPIXELSX      As Long = 88
    Const LOGPIXELSY      As Long = 90
    Const vbPicTypeBitmap As Long = 1
    Const vbSrcCopy       As Long = &HCC0020
    Dim hDC             As LongPtr
    Dim uHdr            As BITMAPINFOHEADER
    Dim hDib            As LongPtr
    Dim hPrevDib        As LongPtr
    Dim hSrcDC          As LongPtr
    Dim hSrcPrevBmp     As LongPtr
    Dim uDesc           As PICTDESC
    Dim hResult         As Long
    Dim vErr            As Variant
    
    On Error GoTo EH
    hDC = CreateCompatibleDC(0)
    Call SetStretchBltMode(hDC, HALFTONE)
    With uHdr
        .biSize = LenB(uHdr)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = NewWidth
        .biHeight = -NewHeight
        .biSizeImage = (4 * NewWidth) * NewHeight
    End With
    hDib = CreateDIBSection(hDC, uHdr, DIB_RGB_COLORS, 0, 0, 0)
    hPrevDib = SelectObject(hDC, hDib)
    If pPicture.Type = vbPicTypeBitmap Then
        hSrcDC = CreateCompatibleDC(0)
        hSrcPrevBmp = SelectObject(hSrcDC, pPicture.Handle)
        Call StretchBlt(hDC, 0, 0, NewWidth, NewHeight, hSrcDC, 0, 0, _
            HM2Pix(pPicture.Width, GetDeviceCaps(hDC, LOGPIXELSX)), _
            HM2Pix(pPicture.Height, GetDeviceCaps(hDC, LOGPIXELSY)), vbSrcCopy)
    Else
        pPicture.Render CLng(hDC), 0, 0, NewWidth, NewHeight, 0, pPicture.Height, pPicture.Width, -pPicture.Height, ByVal 0
    End If
    Call SelectObject(hDC, hPrevDib)
    hPrevDib = 0
    With uDesc
        .Size = LenB(uDesc)
        .Type = vbPicTypeBitmap
        .hBmpOrIcon = hDib
    End With
    If IID_IPicture(0) = 0 Then
        IID_IPicture(0) = &H7BF80980: IID_IPicture(1) = &H101ABF32: IID_IPicture(2) = &HAA00BB8B: IID_IPicture(3) = &HAB0C3000
    End If
    hResult = OleCreatePictureIndirect(uDesc, IID_IPicture(0), 1, QRCodegenResizePicture)
    If hResult < 0 Then
        Err.Raise hResult, "OleCreatePictureIndirect"
    End If
    '--- not to destroy DIB, it's already owned by returned StdPicture
    hDib = 0
QH:
    If hSrcPrevBmp <> 0 Then
        Call SelectObject(hSrcDC, hSrcPrevBmp)
        hSrcPrevBmp = 0
    End If
    If hSrcDC <> 0 Then
        Call DeleteDC(hSrcDC)
        hSrcDC = 0
    End If
    If hPrevDib <> 0 Then
        Call SelectObject(hDC, hPrevDib)
        hPrevDib = 0
    End If
    If hDC <> 0 Then
        Call DeleteDC(hDC)
        hDC = 0
    End If
    If hDib <> 0 Then
        Call DeleteObject(hDib)
        hDib = 0
    End If
    If IsArray(vErr) Then
        On Error GoTo 0
        Err.Raise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    Resume QH
End Function

Public Function QRCodegenDebugDump(baQrCode() As Byte) As String
    Dim lQrSize         As Long
    Dim aRows()         As String
    Dim lX              As Long
    Dim lY              As Long
    
    If UBound(baQrCode) >= 0 Then
        lQrSize = QRCodegenGetSize(baQrCode)
        ReDim aRows(0 To lQrSize - 1) As String
        For lY = 0 To lQrSize - 1
            For lX = 0 To lQrSize - 1
                aRows(lY) = aRows(lY) & IIf(QRCodegenGetModule(baQrCode, lX, lY), "##", "  ")
            Next
            aRows(lY) = RTrim$(aRows(lY))
        Next
        QRCodegenDebugDump = Join(aRows, vbCrLf)
    End If
End Function

Public Function QRCodegenGetSize(baQrCode() As Byte) As Long
    If UBound(baQrCode) >= 0 Then
        QRCodegenGetSize = baQrCode(0)
        Debug.Assert VERSION_MIN * 4 + 17 <= QRCodegenGetSize And QRCodegenGetSize <= VERSION_MAX * 4 + 17
    End If
End Function

Public Function QRCodegenGetModule(baQrCode() As Byte, ByVal lX As Long, ByVal lY As Long) As Boolean
    Dim lQrSize         As Long
    
    If UBound(baQrCode) > 0 Then
        lQrSize = baQrCode(0)
        If 0 <= lX And lX < lQrSize And 0 <= lY And lY < lQrSize Then
            QRCodegenGetModule = pvGetModuleBounded(baQrCode, lX, lY)
        End If
    End If
End Function

Public Function QRCodegenIsNumeric(sText As String) As Boolean
    If LenB(sText) <> 0 Then
        QRCodegenIsNumeric = Not (sText Like "*[!0-9]*")
    End If
End Function

Public Function QRCodegenIsAlphanumeric(sText As String) As Boolean
    Dim lIdx            As Long
    
    If LenB(sText) <> 0 Then
        For lIdx = 1 To Len(sText)
            If InStr(ALPHANUMERIC_CHARSET, Mid$(sText, lIdx, 1)) = 0 Then
                Exit Function
            End If
        Next
    End If
    QRCodegenIsAlphanumeric = True
End Function

Public Function QRCodegenCalcSegmentBufferSize(ByVal eMode As QRCodegenMode, ByVal lNumChars As Long) As Long
    Dim lSize           As Long
    
    lSize = pvCalcSegmentBitLength(eMode, lNumChars)
    If lSize = -1 Then
        QRCodegenCalcSegmentBufferSize = LONG_MAX
    Else
        Debug.Assert 0 <= lSize And lSize < INT16_MAX
        QRCodegenCalcSegmentBufferSize = (lSize + 7) \ 8
    End If
End Function

Public Function QRCodegenMakeBytes(baData() As Byte) As QRCodegenSegment
    With QRCodegenMakeBytes
        .Mode = QRCodegenMode_BYTE
        .BitLength = pvCalcSegmentBitLength(.Mode, UBound(baData) + 1)
        Debug.Assert .BitLength <> -1
        .NumChars = UBound(baData) + 1
        .Data = baData
    End With
End Function

Public Function QRCodegenMakeNumeric(sDigits As String) As QRCodegenSegment
    Dim lLen            As Long
    Dim lBitLen         As Long
    Dim lAccumData      As Long
    Dim lAccumCount     As Long
    Dim lIdx            As Long
    Dim lDigit          As Long
    
    With QRCodegenMakeNumeric
        lLen = Len(sDigits)
        .Mode = QRCodegenMode_NUMERIC
        lBitLen = pvCalcSegmentBitLength(.Mode, lLen)
        Debug.Assert lBitLen <> -1
        ReDim .Data(0 To (lBitLen + 7) \ 8 - 1) As Byte
        .NumChars = lLen
        For lIdx = 1 To lLen
            lDigit = Asc(Mid$(sDigits, lIdx, 1)) - 48    '--- Asc("0") = 48
            Debug.Assert 0 <= lDigit And lDigit <= 9
            lAccumData = lAccumData * 10 + lDigit
            lAccumCount = lAccumCount + 1
            If lAccumCount = 3 Then
                pvAppendBitsToBuffer lAccumData, 10, .Data, .BitLength
                lAccumData = 0
                lAccumCount = 0
            End If
        Next
        If lAccumCount > 0 Then
            pvAppendBitsToBuffer lAccumData, lAccumCount * 3 + 1, .Data, .BitLength
        End If
        Debug.Assert lBitLen = .BitLength
    End With
End Function

Public Function QRCodegenMakeAlphanumeric(sText As String) As QRCodegenSegment
    Dim lLen            As Long
    Dim lBitLen         As Long
    Dim lAccumData      As Long
    Dim lAccumCount     As Long
    Dim lIdx            As Long
    Dim lChar          As Long
    
    With QRCodegenMakeAlphanumeric
        lLen = Len(sText)
        .Mode = QRCodegenMode_ALPHANUMERIC
        lBitLen = pvCalcSegmentBitLength(.Mode, lLen)
        Debug.Assert lBitLen <> -1
        ReDim .Data(0 To (lBitLen + 7) \ 8 - 1) As Byte
        .NumChars = lLen
        For lIdx = 1 To lLen
            lChar = InStr(ALPHANUMERIC_CHARSET, Mid$(sText, lIdx, 1)) - 1
            Debug.Assert 0 <= lChar
            lAccumData = lAccumData * 45 + lChar
            lAccumCount = lAccumCount + 1
            If lAccumCount = 2 Then
                pvAppendBitsToBuffer lAccumData, 11, .Data, .BitLength
                lAccumData = 0
                lAccumCount = 0
            End If
        Next
        If lAccumCount > 0 Then
            pvAppendBitsToBuffer lAccumData, 6, .Data, .BitLength
        End If
        Debug.Assert lBitLen = .BitLength
    End With
End Function

Public Function QRCodegenMakeEci(ByVal lAssignVal As Long) As QRCodegenSegment
    With QRCodegenMakeEci
        .Mode = QRCodegenMode_ECI
        ReDim .Data(0 To 2) As Byte
        If lAssignVal < 0 Then
            Debug.Assert False
        ElseIf lAssignVal < LNG_POW2(7) Then
            pvAppendBitsToBuffer lAssignVal, 8, .Data, .BitLength
        ElseIf lAssignVal < LNG_POW2(14) Then
            pvAppendBitsToBuffer 2, 2, .Data, .BitLength
            pvAppendBitsToBuffer lAssignVal, 14, .Data, .BitLength
        ElseIf lAssignVal < 1000000 Then
            pvAppendBitsToBuffer 6, 3, .Data, .BitLength
            pvAppendBitsToBuffer lAssignVal \ LNG_POW2(10), 11, .Data, .BitLength
            pvAppendBitsToBuffer lAssignVal And &H3FF, 10, .Data, .BitLength
        Else
            Debug.Assert False
        End If
    End With
End Function

'= private ===============================================================

Private Sub pvInit()
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    If ECC_CODEWORDS_PER_BLOCK(0, 0) <> 0 Then
        Exit Sub
    End If
    LNG_POW2(0) = 1
    For lIdx = 1 To UBound(LNG_POW2) - 1
        LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
    Next
    LNG_POW2(31) = &H80000000
    vSplit = Split("-1| 7|10|15|20|26|18|20|24|30|18|20|24|26|30|22|24|28|30|28|28|28|28|30|30|26|28|30|30|30|30|30|30|30|30|30|30|30|30|30|30|" & _
                   "-1|10|16|26|18|24|16|18|22|22|26|30|22|22|24|24|28|28|26|26|26|26|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|" & _
                   "-1|13|22|18|26|18|24|18|22|20|24|28|26|24|20|30|24|28|28|26|30|28|30|30|30|30|28|30|30|30|30|30|30|30|30|30|30|30|30|30|30|" & _
                   "-1|17|28|22|16|22|28|26|26|24|28|24|28|22|24|24|30|28|28|26|28|30|24|30|30|30|30|30|30|30|30|30|30|30|30|30|30|30|30|30|30", "|")
    For lIdx = 0 To UBound(vSplit)
        ECC_CODEWORDS_PER_BLOCK(lIdx \ 41, lIdx Mod 41) = vSplit(lIdx)
    Next
    vSplit = Split("-1|1|1|1|1|1|2|2|2|2|4| 4| 4| 4| 4| 6| 6| 6| 6| 7| 8| 8| 9| 9|10|12|12|12|13|14|15|16|17|18|19|19|20|21|22|24|25|" & _
                   "-1|1|1|1|2|2|4|4|4|5|5| 5| 8| 9| 9|10|10|11|13|14|16|17|17|18|20|21|23|25|26|28|29|31|33|35|37|38|40|43|45|47|49|" & _
                   "-1|1|1|2|2|4|4|6|6|8|8| 8|10|12|16|12|17|16|18|21|20|23|23|25|27|29|34|34|35|38|40|43|45|48|51|53|56|59|62|65|68|" & _
                   "-1|1|1|2|4|4|4|5|6|8|8|11|11|16|16|18|16|19|21|25|25|25|34|30|32|35|37|40|42|45|48|51|54|57|60|63|66|70|74|77|81", "|")
    For lIdx = 0 To UBound(vSplit)
        NUM_ERROR_CORRECTION_BLOCKS(lIdx \ 41, lIdx Mod 41) = vSplit(lIdx)
    Next
End Sub

Private Function pvGetNumDataCodewords(ByVal lVersion As Long, ByVal eEcl As QRCodegenEcc) As Long
    Debug.Assert QRCodegenEcc_LOW <= eEcl And eEcl <= QRCodegenEcc_HIGH
    pvGetNumDataCodewords = pvGetNumRawDataModules(lVersion) \ 8 - ECC_CODEWORDS_PER_BLOCK(eEcl, lVersion) * NUM_ERROR_CORRECTION_BLOCKS(eEcl, lVersion)
End Function
 
Private Function pvGetNumRawDataModules(ByVal lVersion As Long) As Long
    Dim lNumAlign       As Long
    
    Debug.Assert VERSION_MIN <= lVersion And lVersion <= VERSION_MAX
    pvGetNumRawDataModules = (16 * lVersion + 128) * lVersion + 64
    If lVersion >= 2 Then
        lNumAlign = lVersion \ 7 + 2
        pvGetNumRawDataModules = pvGetNumRawDataModules - (25 * lNumAlign - 10) * lNumAlign + 55
        If lVersion >= 7 Then
            pvGetNumRawDataModules = pvGetNumRawDataModules - 36
        End If
    End If
    Debug.Assert 208 <= pvGetNumRawDataModules And pvGetNumRawDataModules <= 29648
End Function
 
Private Function pvGetTotalBits(uSegments() As QRCodegenSegment, ByVal lVersion As Long) As Long
    Dim lIdx            As Long
    Dim lCcBits         As Long
    
    For lIdx = 0 To UBound(uSegments)
        lCcBits = pvNumCharCountBits(uSegments(lIdx).Mode, lVersion)
        Debug.Assert 0 <= lCcBits And lCcBits <= 16
        If uSegments(lIdx).NumChars >= LNG_POW2(lCcBits) Then
            pvGetTotalBits = -1
            GoTo QH
        End If
        pvGetTotalBits = pvGetTotalBits + 4 + lCcBits + uSegments(lIdx).BitLength
        If pvGetTotalBits > INT16_MAX Then
            pvGetTotalBits = -1
            GoTo QH
        End If
    Next
    Debug.Assert 0 <= pvGetTotalBits And pvGetTotalBits <= INT16_MAX
QH:
End Function

Private Function pvNumCharCountBits(ByVal eMode As QRCodegenMode, ByVal lVersion As Long) As Long
    Dim lIdx            As Long
    
    lIdx = (lVersion + 7) \ 17 + 1
    Select Case eMode
    Case QRCodegenMode_NUMERIC
        pvNumCharCountBits = Choose(lIdx, 10, 12, 14)
    Case QRCodegenMode_ALPHANUMERIC
        pvNumCharCountBits = Choose(lIdx, 9, 11, 13)
    Case QRCodegenMode_BYTE
        pvNumCharCountBits = Choose(lIdx, 8, 16, 16)
    Case QRCodegenMode_KANJI
        pvNumCharCountBits = Choose(lIdx, 0, 10, 12)
    Case QRCodegenMode_ECI
        pvNumCharCountBits = 0
    Case Else
        Debug.Assert False
    End Select
End Function

Private Sub pvAddEccAndInterleave(baData() As Byte, ByVal lVersion As Long, ByVal eEcl As QRCodegenEcc, baResult() As Byte)
    Dim lNumBlocks      As Long
    Dim lBlockEccLen    As Long
    Dim lRawCodewords   As Long
    Dim lDataLen        As Long
    Dim lNumShortBlocks As Long
    Dim lShortBlockDataLen  As Long
    Dim baDiv()         As Byte
    Dim lIdx            As Long
    Dim lBlockPos       As Long
    Dim lBlockLen       As Long
    Dim baEcc()         As Byte
    Dim lJdx            As Long
    Dim lKdx            As Long
    
    Debug.Assert VERSION_MIN <= lVersion And lVersion <= VERSION_MAX
    Debug.Assert QRCodegenEcc_LOW <= eEcl And eEcl <= QRCodegenEcc_HIGH
    lNumBlocks = NUM_ERROR_CORRECTION_BLOCKS(eEcl, lVersion)
    lBlockEccLen = ECC_CODEWORDS_PER_BLOCK(eEcl, lVersion)
    lRawCodewords = pvGetNumRawDataModules(lVersion) \ 8
    lDataLen = pvGetNumDataCodewords(lVersion, eEcl)
    lNumShortBlocks = lNumBlocks - lRawCodewords Mod lNumBlocks
    lShortBlockDataLen = lRawCodewords \ lNumBlocks - lBlockEccLen
    ReDim baResult(0 To lDataLen + lBlockEccLen * lNumBlocks - 1) As Byte
    pvReedSolomonComputeDivisor lBlockEccLen, baDiv
    For lIdx = 0 To lNumBlocks - 1
        lBlockLen = lShortBlockDataLen + IIf(lIdx < lNumShortBlocks, 0, 1)
        pvReedSolomonComputeRemainder baData, lBlockPos, lBlockLen, baDiv, lBlockEccLen, baEcc
        lKdx = lIdx
        For lJdx = 0 To lBlockLen - 1
            If lJdx = lShortBlockDataLen Then
                lKdx = lKdx - lNumShortBlocks
            End If
            baResult(lKdx) = baData(lBlockPos + lJdx)
            lKdx = lKdx + lNumBlocks
        Next
        lKdx = lDataLen + lIdx
        For lJdx = 0 To lBlockEccLen - 1
            baResult(lKdx) = baEcc(lJdx)
            lKdx = lKdx + lNumBlocks
        Next
        lBlockPos = lBlockPos + lBlockLen
    Next
End Sub

Private Sub pvReedSolomonComputeDivisor(ByVal lDegree As Long, baResult() As Byte)
    Dim lRoot           As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    ReDim baResult(0 To lDegree - 1) As Byte
    baResult(lDegree - 1) = 1
    lRoot = 1
    For lIdx = 0 To lDegree - 1
        For lJdx = 0 To lDegree - 1
            baResult(lJdx) = pvReedSolomonMultiply(baResult(lJdx), lRoot)
            If lJdx + 1 < lDegree Then
                baResult(lJdx) = baResult(lJdx) Xor baResult(lJdx + 1)
            End If
        Next
        lRoot = pvReedSolomonMultiply(lRoot, 2)
    Next
End Sub

Private Sub pvReedSolomonComputeRemainder( _
            baData() As Byte, ByVal lDataPos As Long, ByVal lDataSize As Long, _
            baGen() As Byte, ByVal lDegree As Long, baResult() As Byte)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim bFactor         As Byte
    
    ReDim baResult(0 To lDegree - 1) As Byte
    For lIdx = lDataPos To lDataPos + lDataSize - 1
        bFactor = baData(lIdx) Xor baResult(0)
        For lJdx = 0 To lDegree - 2
            baResult(lJdx) = baResult(lJdx + 1) Xor pvReedSolomonMultiply(baGen(lJdx), bFactor)
        Next
        baResult(lDegree - 1) = pvReedSolomonMultiply(baGen(lJdx), bFactor)
    Next
End Sub

Private Function pvReedSolomonMultiply(ByVal bX As Byte, ByVal bY As Byte) As Byte
    Dim lIdx            As Long
    Dim lTemp           As Long
    
    For lIdx = 7 To 0 Step -1
        If (pvReedSolomonMultiply And &H80) <> 0 Then
            lTemp = &H11D
        Else
            lTemp = 0
        End If
        pvReedSolomonMultiply = ((pvReedSolomonMultiply * 2) Xor lTemp) And &HFF
        If (bY And LNG_POW2(lIdx)) <> 0 Then
            pvReedSolomonMultiply = pvReedSolomonMultiply Xor bX
        End If
    Next
End Function

Private Sub pvInitializeFunctionModules(ByVal lVersion As Long, baQrCode() As Byte)
    Dim lQrSize         As Long
    Dim lNumAlign       As Long
    Dim baAlignPatPos(0 To 6) As Byte
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    lQrSize = lVersion * 4 + 17
    ReDim baQrCode(0 To (lQrSize * lQrSize + 7) \ 8) As Byte
    baQrCode(0) = lQrSize
    '--- Fill horizontal and vertical timing patterns
    pvFillRectangle 6, 0, 1, lQrSize, baQrCode
    pvFillRectangle 0, 6, lQrSize, 1, baQrCode
    '--- Fill 3 finder patterns (all corners except bottom right) and format bits
    pvFillRectangle 0, 0, 9, 9, baQrCode
    pvFillRectangle lQrSize - 8, 0, 8, 9, baQrCode
    pvFillRectangle 0, lQrSize - 8, 9, 8, baQrCode
    '--- Fill numerous alignment patterns
    lNumAlign = pvGetAlignmentPatternPositions(lVersion, baAlignPatPos)
    For lIdx = 0 To lNumAlign - 1
        For lJdx = 0 To lNumAlign - 1
            If (lIdx = 0 And lJdx = 0) Or (lIdx = 0 And lJdx = lNumAlign - 1) Or (lIdx = lNumAlign - 1 And lJdx = 0) Then
                '--- Don't draw on the three finder corners
            Else
                pvFillRectangle baAlignPatPos(lIdx) - 2, baAlignPatPos(lJdx) - 2, 5, 5, baQrCode
            End If
        Next
    Next
    '--- Fill version blocks
    If lVersion >= 7 Then
        pvFillRectangle lQrSize - 11, 0, 3, 6, baQrCode
        pvFillRectangle 0, lQrSize - 11, 6, 3, baQrCode
    End If
End Sub

Private Sub pvDrawLightFunctionModules(ByVal lVersion As Long, baQrCode() As Byte)
    Dim lQrSize         As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lKdx            As Long
    Dim lDy             As Long
    Dim lDx             As Long
    Dim lDist           As Long
    Dim lNumAlign       As Long
    Dim baAlignPatPos(0 To 6) As Byte
    Dim bIsDark         As Boolean
    Dim lRem            As Long
    Dim lBits           As Long
    
    lQrSize = baQrCode(0)
    '--- Draw horizontal and vertical timing patterns
    For lIdx = 7 To lQrSize - 7 Step 2
        pvSetModuleBounded baQrCode, 6, lIdx, False
        pvSetModuleBounded baQrCode, lIdx, 6, False
    Next
    '--- Draw 3 finder patterns (all corners except bottom right; overwrites some timing modules)
    For lDy = -4 To 4
        For lDx = -4 To 4
            lDist = Abs(lDx)
            If Abs(lDy) > lDist Then
                lDist = Abs(lDy)
            End If
            If lDist = 2 Or lDist = 4 Then
                pvSetModuleUnbounded baQrCode, 3 + lDx, 3 + lDy, False
                pvSetModuleUnbounded baQrCode, lQrSize - 4 + lDx, 3 + lDy, False
                pvSetModuleUnbounded baQrCode, 3 + lDx, lQrSize - 4 + lDy, False
            End If
        Next
    Next
    '--- Draw numerous alignment patterns
    lNumAlign = pvGetAlignmentPatternPositions(lVersion, baAlignPatPos)
    For lIdx = 0 To lNumAlign - 1
        For lJdx = 0 To lNumAlign - 1
            If (lIdx = 0 And lJdx = 0) Or (lIdx = 0 And lJdx = lNumAlign - 1) Or (lIdx = lNumAlign - 1 And lJdx = 0) Then
                '--- Don't draw on the three finder corners
            Else
                For lDy = -1 To 1
                    For lDx = -1 To 1
                        bIsDark = (lDx = 0 And lDy = 0)
                        pvSetModuleBounded baQrCode, baAlignPatPos(lIdx) + lDx, baAlignPatPos(lJdx) + lDy, bIsDark
                    Next
                Next
            End If
        Next
    Next
    '--- Draw version blocks
    If lVersion >= 7 Then
        '--- Calculate error correction code and pack bits
        lRem = lVersion
        For lIdx = 0 To 11
            lRem = (lRem * 2) Xor ((lRem \ LNG_POW2(11)) * &H1F25)
        Next
        lBits = lVersion * LNG_POW2(12) Or lRem
        Debug.Assert lBits < LNG_POW2(18)
        '--- Draw two copies
        For lIdx = 0 To 5
            For lJdx = 0 To 2
                lKdx = lQrSize - 11 + lJdx
                bIsDark = ((lBits And 1) <> 0)
                pvSetModuleBounded baQrCode, lKdx, lIdx, bIsDark
                pvSetModuleBounded baQrCode, lIdx, lKdx, bIsDark
                lBits = lBits \ 2
            Next
        Next
    End If
End Sub

Private Sub pvDrawFormatBits(ByVal eEcl As QRCodegenEcc, ByVal eMask As QRCodegenMask, baQrCode() As Byte)
    Dim lData           As Long
    Dim lRem            As Long
    Dim lBits           As Long
    Dim lIdx            As Long
    Dim lQrSize         As Long
    
    '--- Calculate error correction code and pack bits
    lData = Choose(eEcl + 1, 1, 0, 3, 2) * 8 Or eMask
    lRem = lData
    For lIdx = 0 To 9
        lRem = (lRem * 2) Xor ((lRem \ LNG_POW2(9)) * &H537)
    Next
    lBits = (lData * LNG_POW2(10) Or lRem) Xor &H5412
    '--- Draw first copy
    For lIdx = 0 To 5
        pvSetModuleBounded baQrCode, 8, lIdx, pvGetBit(lBits, lIdx)
    Next
    pvSetModuleBounded baQrCode, 8, 7, pvGetBit(lBits, 6)
    pvSetModuleBounded baQrCode, 8, 8, pvGetBit(lBits, 7)
    pvSetModuleBounded baQrCode, 7, 8, pvGetBit(lBits, 8)
    For lIdx = 9 To 14
        pvSetModuleBounded baQrCode, 14 - lIdx, 8, pvGetBit(lBits, lIdx)
    Next
    '--- Draw second copy
    lQrSize = baQrCode(0)
    For lIdx = 0 To 7
        pvSetModuleBounded baQrCode, lQrSize - 1 - lIdx, 8, pvGetBit(lBits, lIdx)
    Next
    For lIdx = 8 To 14
        pvSetModuleBounded baQrCode, 8, lQrSize - 15 + lIdx, pvGetBit(lBits, lIdx)
    Next
    pvSetModuleBounded baQrCode, 8, lQrSize - 8, True
End Sub

Private Function pvGetAlignmentPatternPositions(ByVal lVersion As Long, baResult() As Byte) As Long
    Dim lNumAlign       As Long
    Dim lStep           As Long
    Dim lIdx            As Long
    Dim lPos            As Long
    
    If lVersion > 1 Then
        lNumAlign = lVersion \ 7 + 2
        lStep = IIf(lVersion = 32, 26, ((lVersion * 4 + lNumAlign * 2 + 1) \ (lNumAlign * 2 - 2)) * 2)
        lPos = lVersion * 4 + 10
        For lIdx = lNumAlign - 1 To 1 Step -1
            baResult(lIdx) = lPos
            lPos = lPos - lStep
        Next
        baResult(0) = 6
        pvGetAlignmentPatternPositions = lNumAlign
    End If
End Function

Private Sub pvFillRectangle(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, baQrCode() As Byte)
    Dim lDy             As Long
    Dim lDx             As Long
    
    For lDy = 0 To lHeight - 1
        For lDx = 0 To lWidth - 1
            pvSetModuleBounded baQrCode, lLeft + lDx, lTop + lDy, True
        Next
    Next
End Sub

Private Sub pvDrawCodewords(baData() As Byte, baQrCode() As Byte)
    Dim lQrSize         As Long
    Dim lBitLen         As Long
    Dim lIdx            As Long
    Dim lRight          As Long
    Dim lVert           As Long
    Dim lJdx            As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim bIsDark         As Boolean
    
    lQrSize = baQrCode(0)
    lBitLen = (UBound(baData) + 1) * 8
    For lRight = lQrSize - 1 To 1 Step -2
        If lRight = 6 Then
            lRight = 5
        End If
        For lVert = 0 To lQrSize - 1
            For lJdx = 0 To 1
                lX = lRight - lJdx
                If ((lRight + 1) And 2) = 0 Then
                    lY = lQrSize - 1 - lVert
                Else
                    lY = lVert
                End If
                If Not pvGetModuleBounded(baQrCode, lX, lY) And lIdx < lBitLen Then
                    bIsDark = pvGetBit(baData(lIdx \ 8), 7 - (lIdx And 7))
                    pvSetModuleBounded baQrCode, lX, lY, bIsDark
                    lIdx = lIdx + 1
                End If
            Next
        Next
    Next
    Debug.Assert lIdx = lBitLen
End Sub

Private Sub pvApplyMask(baFunctionModules() As Byte, baQrCode() As Byte, ByVal eMask As QRCodegenMask)
    Dim lQrSize         As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim bInvert         As Boolean
    Dim bVal            As Boolean
    
    Debug.Assert QRCodegenMask_0 <= eMask And eMask <= QRCodegenMask_7
    lQrSize = baQrCode(0)
    For lY = 0 To lQrSize - 1
        For lX = 0 To lQrSize - 1
            If Not pvGetModuleBounded(baFunctionModules, lX, lY) Then
                Select Case eMask
                Case QRCodegenMask_0
                    bInvert = (lX + lY) Mod 2 = 0
                Case QRCodegenMask_1
                    bInvert = lY Mod 2 = 0
                Case QRCodegenMask_2
                    bInvert = lX Mod 3 = 0
                Case QRCodegenMask_3
                    bInvert = (lX + lY) Mod 3 = 0
                Case QRCodegenMask_4
                    bInvert = (lX \ 3 + lY \ 2) Mod 2 = 0
                Case QRCodegenMask_5
                    bInvert = (lX * lY Mod 2 + lX * lY Mod 3) = 0
                Case QRCodegenMask_6
                    bInvert = (lX * lY Mod 2 + lX * lY Mod 3) Mod 2 = 0
                Case QRCodegenMask_7
                    bInvert = ((lX + lY) Mod 2 + lX * lY Mod 3) Mod 2 = 0
                End Select
                bVal = pvGetModuleBounded(baQrCode, lX, lY)
                pvSetModuleBounded baQrCode, lX, lY, (bVal Xor bInvert)
            End If
        Next
    Next
End Sub

Private Function pvGetPenaltyScore(baQrCode() As Byte) As Long
    Dim lQrSize         As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim bRunColor       As Boolean
    Dim lRunX           As Long
    Dim lRunY           As Long
    Dim aRunHistory()   As Long
    Dim lDark           As Long
    Dim lTotal          As Long
    Dim lKdx            As Long
    
    lQrSize = baQrCode(0)
    '--- Adjacent modules in row having same color, and finder-like patterns
    For lY = 0 To lQrSize - 1
        bRunColor = False
        lRunX = 0
        ReDim aRunHistory(0 To 6) As Long
        For lX = 0 To lQrSize - 1
            If pvGetModuleBounded(baQrCode, lX, lY) = bRunColor Then
                lRunX = lRunX + 1
                If lRunX = 5 Then
                    pvGetPenaltyScore = pvGetPenaltyScore + PENALTY_N1
                ElseIf lRunX > 5 Then
                    pvGetPenaltyScore = pvGetPenaltyScore + 1
                End If
            Else
                pvFinderPenaltyAddHistory lRunX, aRunHistory, lQrSize
                If Not bRunColor Then
                    pvGetPenaltyScore = pvGetPenaltyScore + pvFinderPenaltyCountPatterns(aRunHistory, lQrSize) * PENALTY_N3
                End If
                bRunColor = pvGetModuleBounded(baQrCode, lX, lY)
                lRunX = 1
            End If
        Next
        pvGetPenaltyScore = pvGetPenaltyScore + pvFinderPenaltyTerminateAndCount(bRunColor, lRunX, aRunHistory, lQrSize) * PENALTY_N3
    Next
    '--- Adjacent modules in column having same color, and finder-like patterns
    For lX = 0 To lQrSize - 1
        bRunColor = False
        lRunY = 0
        ReDim aRunHistory(0 To 6) As Long
        For lY = 0 To lQrSize - 1
            If pvGetModuleBounded(baQrCode, lX, lY) = bRunColor Then
                lRunY = lRunY + 1
                If lRunY = 5 Then
                    pvGetPenaltyScore = pvGetPenaltyScore + PENALTY_N1
                ElseIf lRunY > 5 Then
                    pvGetPenaltyScore = pvGetPenaltyScore + 1
                End If
            Else
                pvFinderPenaltyAddHistory lRunY, aRunHistory, lQrSize
                If Not bRunColor Then
                    pvGetPenaltyScore = pvGetPenaltyScore + pvFinderPenaltyCountPatterns(aRunHistory, lQrSize) * PENALTY_N3
                End If
                bRunColor = pvGetModuleBounded(baQrCode, lX, lY)
                lRunY = 1
            End If
        Next
        pvGetPenaltyScore = pvGetPenaltyScore + pvFinderPenaltyTerminateAndCount(bRunColor, lRunY, aRunHistory, lQrSize) * PENALTY_N3
    Next
    '--- 2*2 blocks of modules having same color
    For lY = 0 To lQrSize - 2
        For lX = 0 To lQrSize - 2
            bRunColor = pvGetModuleBounded(baQrCode, lX, lY)
            If bRunColor = pvGetModuleBounded(baQrCode, lX + 1, lY) And _
                    bRunColor = pvGetModuleBounded(baQrCode, lX, lY + 1) And _
                    bRunColor = pvGetModuleBounded(baQrCode, lX + 1, lY + 1) Then
                pvGetPenaltyScore = pvGetPenaltyScore + PENALTY_N2
            End If
        Next
    Next
    '--- Balance of dark and light modules
    For lY = 0 To lQrSize - 1
        For lX = 0 To lQrSize - 1
            If pvGetModuleBounded(baQrCode, lX, lY) Then
               lDark = lDark + 1
            End If
        Next
    Next
    lTotal = lQrSize * lQrSize
    '--- Compute the smallest integer k >= 0 such that (45-5k)% <= dark/total <= (55+5k)%
    lKdx = ((Abs(lDark * 20 - lTotal * 10) + lTotal - 1) \ lTotal) - 1
    Debug.Assert 0 <= lKdx And lKdx <= 9
    pvGetPenaltyScore = pvGetPenaltyScore + lKdx * PENALTY_N4
    Debug.Assert 0 <= pvGetPenaltyScore And pvGetPenaltyScore <= 2568888
End Function

Private Function pvFinderPenaltyCountPatterns(aRunHistory() As Long, ByVal lQrSize As Long) As Long
    Dim lN              As Long
    Dim bCore           As Boolean
    
    lN = aRunHistory(1)
    Debug.Assert lN <= lQrSize * 3
    bCore = (lN > 0 And aRunHistory(2) = lN And aRunHistory(3) = lN * 3 And aRunHistory(4) = lN And aRunHistory(5) = lN)
    '-- The maximum QR Code size is 177, hence the dark run length n <= 177.
    pvFinderPenaltyCountPatterns = IIf(bCore And aRunHistory(0) >= lN * 4 And aRunHistory(6) >= lN, 1, 0) _
                                 + IIf(bCore And aRunHistory(6) >= lN * 4 And aRunHistory(0) >= lN, 1, 0)
End Function

Private Function pvFinderPenaltyTerminateAndCount(ByVal bCurrentRunColor As Boolean, ByVal lCurrentRunLength As Long, aRunHistory() As Long, ByVal lQrSize As Long) As Long
    If bCurrentRunColor Then
        pvFinderPenaltyAddHistory lCurrentRunLength, aRunHistory, lQrSize
        lCurrentRunLength = 0
    End If
    lCurrentRunLength = lCurrentRunLength + lQrSize
    pvFinderPenaltyAddHistory lCurrentRunLength, aRunHistory, lQrSize
    pvFinderPenaltyTerminateAndCount = pvFinderPenaltyCountPatterns(aRunHistory, lQrSize)
End Function

Private Function pvFinderPenaltyAddHistory(ByVal lCurrentRunLength As Long, aRunHistory() As Long, ByVal lQrSize As Long) As Long
    If aRunHistory(0) = 0 Then
        lCurrentRunLength = lCurrentRunLength + lQrSize
    End If
    Debug.Assert UBound(aRunHistory) + 1 = 7
    aRunHistory(6) = aRunHistory(5)
    aRunHistory(5) = aRunHistory(4)
    aRunHistory(4) = aRunHistory(3)
    aRunHistory(3) = aRunHistory(2)
    aRunHistory(2) = aRunHistory(1)
    aRunHistory(1) = aRunHistory(0)
    aRunHistory(0) = lCurrentRunLength
End Function

Private Function pvGetModuleBounded(baQrCode() As Byte, ByVal lX As Long, ByVal lY As Long) As Boolean
    Dim lQrSize         As Long
    Dim lIndex          As Long
    
    lQrSize = baQrCode(0)
    Debug.Assert 21 <= lQrSize And lQrSize <= 177 And 0 <= lX And lX < lQrSize And 0 <= lY And lY < lQrSize
    lIndex = lY * lQrSize + lX
    pvGetModuleBounded = pvGetBit(baQrCode(lIndex \ 8 + 1), lIndex And 7)
End Function

Private Function pvSetModuleBounded(baQrCode() As Byte, ByVal lX As Long, ByVal lY As Long, ByVal bIsDark As Boolean) As Long
    Dim lQrSize         As Long
    Dim lIndex          As Long
    Dim lByteIndex      As Long
    
    lQrSize = baQrCode(0)
    Debug.Assert 21 <= lQrSize And lQrSize <= 177 And 0 <= lX And lX < lQrSize And 0 <= lY And lY < lQrSize
    lIndex = lY * lQrSize + lX
    lByteIndex = lIndex \ 8 + 1
    If bIsDark Then
        baQrCode(lByteIndex) = baQrCode(lByteIndex) Or LNG_POW2(lIndex And 7)
    Else
        baQrCode(lByteIndex) = baQrCode(lByteIndex) And Not LNG_POW2(lIndex And 7)
    End If
End Function

Private Function pvSetModuleUnbounded(baQrCode() As Byte, ByVal lX As Long, ByVal lY As Long, ByVal bIsDark As Boolean) As Long
    Dim lQrSize         As Long
    
    lQrSize = baQrCode(0)
    If 0 <= lX And lX < lQrSize And 0 <= lY And lY < lQrSize Then
        pvSetModuleBounded baQrCode, lX, lY, bIsDark
    End If
End Function

Private Function pvGetBit(ByVal lX As Long, ByVal lIdx As Long) As Boolean
    pvGetBit = (lX And LNG_POW2(lIdx)) <> 0
End Function

Private Function pvCalcSegmentBitLength(ByVal eMode As QRCodegenMode, ByVal lNumChars As Long) As Long
    If lNumChars > INT16_MAX Then
        pvCalcSegmentBitLength = -1
    Else
        pvCalcSegmentBitLength = lNumChars
        Select Case eMode
        Case QRCodegenMode_NUMERIC
            pvCalcSegmentBitLength = (pvCalcSegmentBitLength * 10 + 2) \ 3
        Case QRCodegenMode_ALPHANUMERIC
            pvCalcSegmentBitLength = (pvCalcSegmentBitLength * 11 + 1) \ 2
        Case QRCodegenMode_BYTE
            pvCalcSegmentBitLength = pvCalcSegmentBitLength * 8
        Case QRCodegenMode_KANJI
            pvCalcSegmentBitLength = pvCalcSegmentBitLength * 13
        Case QRCodegenMode_ECI
            Debug.Assert lNumChars = 0
            pvCalcSegmentBitLength = 3 * 8
        End Select
        If pvCalcSegmentBitLength > INT16_MAX Then
            pvCalcSegmentBitLength = -1
        End If
    End If
End Function

Private Function pvGetBufferLenForVersion(ByVal lVersion As Long) As Long
    pvGetBufferLenForVersion = (((lVersion * 4 + 17) * (lVersion * 4 + 17) + 7) \ 8 + 1)
End Function

Private Sub pvAppendBitsToBuffer(ByVal lVal As Long, ByVal lNumBits As Long, baBuffer() As Byte, lBitLen As Long)
    Dim lIdx            As Long
    
    Debug.Assert 0 <= lNumBits And lNumBits <= 16
    For lIdx = lNumBits - 1 To 0 Step -1
        If (lVal And LNG_POW2(lIdx)) <> 0 Then
            baBuffer(lBitLen \ 8) = baBuffer(lBitLen \ 8) Or LNG_POW2(7 - (lBitLen And 7))
        End If
        lBitLen = lBitLen + 1
    Next
End Sub

Private Function pvToUtf8Array(sText As String) As Byte()
    Const CP_UTF8       As Long = 65001
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    pvToUtf8Array = baRetVal
End Function

Private Function DispCallByVtbl(pUnk As stdole.IUnknown, ByVal lIndex As Long, ParamArray a() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
#If Win64 Then
    Const PTR_SIZE      As Long = 8
#Else
    Const PTR_SIZE      As Long = 4
#End If
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As LongPtr
    Dim hResult         As Long
    
    vParam = a
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * PTR_SIZE, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function

Private Sub pvConstructVectors(baQrCode() As Byte, ByVal SquareModules As Boolean, uVectors() As RECT)
    Dim lQrSize         As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim lHasLeft        As Long
    Dim lHasUp          As Long
    Dim lHasRight       As Long
    Dim lHasDown        As Long
    Dim bHasDiag        As Boolean
    
    lQrSize = QRCodegenGetSize(baQrCode)
    '--- oversize array by one element in all 4 directions
    ReDim uVectors(-1 To lQrSize + 1, -1 To lQrSize + 1) As RECT
    For lY = 0 To lQrSize
        For lX = 0 To lQrSize
            If QRCodegenGetModule(baQrCode, lX, lY) And Not QRCodegenGetModule(baQrCode, lX, lY - 1) Then
                uVectors(lX, lY).Right = LINE_TO
            End If
            If QRCodegenGetModule(baQrCode, lX - 1, lY) And Not QRCodegenGetModule(baQrCode, lX, lY) Then
                uVectors(lX, lY).Bottom = LINE_TO
            End If
            If QRCodegenGetModule(baQrCode, lX - 1, lY - 1) And Not QRCodegenGetModule(baQrCode, lX - 1, lY) Then
                uVectors(lX, lY).Left = LINE_TO
            End If
            If QRCodegenGetModule(baQrCode, lX, lY - 1) And Not QRCodegenGetModule(baQrCode, lX - 1, lY - 1) Then
                uVectors(lX, lY).Top = LINE_TO
            End If
        Next
    Next
    If SquareModules Then
        Exit Sub
    End If
    '--- mark curving vectors
    For lY = 0 To lQrSize - 1
        For lX = 0 To lQrSize - 1
            lHasLeft = -QRCodegenGetModule(baQrCode, lX - 1, lY)
            lHasUp = -QRCodegenGetModule(baQrCode, lX, lY - 1)
            lHasRight = -QRCodegenGetModule(baQrCode, lX + 1, lY)
            lHasDown = -QRCodegenGetModule(baQrCode, lX, lY + 1)
            If lHasLeft + lHasUp + lHasRight + lHasDown < 3 Then
                If QRCodegenGetModule(baQrCode, lX, lY) Then
                    If lHasLeft = 0 And lHasUp = 0 Then
                        uVectors(lX, lY + 1).Top = TURN_RIGHT
                    End If
                    If lHasRight = 0 And lHasUp = 0 Then
                        uVectors(lX, lY).Right = TURN_RIGHT
                    End If
                    If lHasRight = 0 And lHasDown = 0 Then
                        uVectors(lX + 1, lY).Bottom = TURN_RIGHT
                    End If
                    If lHasLeft = 0 And lHasDown = 0 Then
                        uVectors(lX + 1, lY + 1).Left = TURN_RIGHT
                    End If
                Else
                    If QRCodegenGetModule(baQrCode, lX - 1, lY - 1) And lHasLeft > 0 And lHasUp > 0 Then
                        uVectors(lX + 1, lY).Left = TURN_LEFT
                    End If
                    If QRCodegenGetModule(baQrCode, lX + 1, lY - 1) And lHasRight > 0 And lHasUp > 0 Then
                        uVectors(lX + 1, lY + 1).Top = TURN_LEFT
                    End If
                    If QRCodegenGetModule(baQrCode, lX - 1, lY + 1) And lHasLeft > 0 And lHasDown > 0 Then
                        uVectors(lX, lY).Bottom = TURN_LEFT
                    End If
                    If QRCodegenGetModule(baQrCode, lX + 1, lY + 1) And lHasRight > 0 And lHasDown > 0 Then
                        uVectors(lX, lY + 1).Right = TURN_LEFT
                    End If
                End If
            End If
        Next
    Next
    '--- mark wide and super-wide curves
    For lY = 0 To lQrSize - 1
        For lX = 0 To lQrSize - 1
            If QRCodegenGetModule(baQrCode, lX, lY) Then
                lHasLeft = -QRCodegenGetModule(baQrCode, lX - 1, lY)
                If lHasLeft <> 0 Then
                    lHasLeft = lHasLeft + -QRCodegenGetModule(baQrCode, lX - 2, lY)
                End If
                lHasUp = -QRCodegenGetModule(baQrCode, lX, lY - 1)
                If lHasUp <> 0 Then
                    lHasUp = lHasUp + -QRCodegenGetModule(baQrCode, lX, lY - 2)
                End If
                lHasRight = -QRCodegenGetModule(baQrCode, lX + 1, lY)
                If lHasRight <> 0 Then
                    lHasRight = lHasRight + -QRCodegenGetModule(baQrCode, lX + 2, lY)
                End If
                lHasDown = -QRCodegenGetModule(baQrCode, lX, lY + 1)
                If lHasDown <> 0 Then
                    lHasDown = lHasDown + -QRCodegenGetModule(baQrCode, lX, lY + 2)
                End If
                bHasDiag = QRCodegenGetModule(baQrCode, lX - 1, lY - 1) Or QRCodegenGetModule(baQrCode, lX - 1, lY + 1) _
                        Or QRCodegenGetModule(baQrCode, lX + 1, lY - 1) Or QRCodegenGetModule(baQrCode, lX + 1, lY + 1)
                If Not bHasDiag And lHasLeft = 0 And lHasUp = 0 And lHasRight > 0 And lHasDown > 0 Then
                    If uVectors(lX + 2, lY + 1).Left = TURN_LEFT And lHasRight > 1 And lHasDown > 1 Then
                        uVectors(lX, lY + 1).Top = SKIP_TO
                        uVectors(lX, lY + 2).Top = SUPERWIDE_RIGHT
                        uVectors(lX + 2, lY + 1).Left = WIDE_LEFT
                    Else
                        uVectors(lX, lY + 1).Top = WIDE_RIGHT
                    End If
                ElseIf Not bHasDiag And lHasLeft > 0 And lHasUp = 0 And lHasRight = 0 And lHasDown > 0 Then
                    If uVectors(lX, lY + 2).Top = TURN_LEFT And lHasLeft > 1 And lHasDown > 1 Then
                        uVectors(lX, lY).Right = SKIP_TO
                        uVectors(lX - 1, lY).Right = SUPERWIDE_RIGHT
                        uVectors(lX, lY + 2).Top = WIDE_LEFT
                    Else
                        uVectors(lX, lY).Right = WIDE_RIGHT
                    End If
                ElseIf Not bHasDiag And lHasLeft > 0 And lHasUp > 0 And lHasRight = 0 And lHasDown = 0 Then
                    If uVectors(lX - 1, lY).Right = TURN_LEFT And lHasLeft > 1 And lHasUp > 1 Then
                        uVectors(lX + 1, lY).Bottom = SKIP_TO
                        uVectors(lX + 1, lY - 1).Bottom = SUPERWIDE_RIGHT
                        uVectors(lX - 1, lY).Right = WIDE_LEFT
                    Else
                        uVectors(lX + 1, lY).Bottom = WIDE_RIGHT
                    End If
                ElseIf Not bHasDiag And lHasLeft = 0 And lHasUp > 0 And lHasRight > 0 And lHasDown = 0 Then
                    If uVectors(lX + 1, lY - 1).Bottom = TURN_LEFT And lHasUp > 1 And lHasRight > 1 Then
                        uVectors(lX + 1, lY + 1).Left = SKIP_TO
                        uVectors(lX + 2, lY + 1).Left = SUPERWIDE_RIGHT
                        uVectors(lX + 1, lY - 1).Bottom = WIDE_LEFT
                    Else
                        uVectors(lX + 1, lY + 1).Left = WIDE_RIGHT
                    End If
                End If
            End If
        Next
    Next
End Sub

Private Sub pvConstructPolygons(uVectors() As RECT, ByVal ModuleSize As Long, uPoints() As POINTAPI, aSizes() As Long)
    Const TURN_STEPS     As Long = 4
    Const WIDE_STEPS     As Long = 2 * TURN_STEPS
    Const SUPERWIDE_STEPS As Long = 3 * TURN_STEPS
    Dim lX              As Long
    Dim lY              As Long
    Dim lNumPolys       As Long
    Dim lPos            As Long
    Dim lSize           As Long
    Dim lValue          As Long
    
    ReDim uPoints(0 To 8) As POINTAPI
    ReDim aSizes(0 To 8) As Long
    Do While pvFindStartVector(uVectors, lX, lY)
        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize, lY * ModuleSize
        Do
            With uVectors(lX, lY)
                If .Right <> 0 Then
                    lValue = .Right
                    .Right = 0
                    Select Case lValue
                    Case TURN_LEFT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize + ModuleSize \ 2, lY * ModuleSize
                        pvAppendLeftTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize, lY * ModuleSize - ModuleSize \ 2, TURN_STEPS
                    Case TURN_RIGHT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize + ModuleSize \ 2, lY * ModuleSize
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize, lY * ModuleSize + ModuleSize \ 2, TURN_STEPS
                    Case WIDE_LEFT
                        pvAppendLeftTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize, lY * ModuleSize - ModuleSize, WIDE_STEPS
                    Case WIDE_RIGHT
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize, lY * ModuleSize + ModuleSize, WIDE_STEPS
                    Case SUPERWIDE_RIGHT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize + ModuleSize \ 2, lY * ModuleSize
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize + 2 * ModuleSize, lY * ModuleSize + ModuleSize * 3 \ 2, SUPERWIDE_STEPS
                    End Select
                    lX = lX + 1
                ElseIf .Bottom <> 0 Then
                    lValue = .Bottom
                    .Bottom = 0
                    Select Case lValue
                    Case TURN_LEFT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize, lY * ModuleSize + ModuleSize \ 2
                        pvAppendLeftTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize \ 2, lY * ModuleSize + ModuleSize, TURN_STEPS
                    Case TURN_RIGHT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize, lY * ModuleSize + ModuleSize \ 2
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize \ 2, lY * ModuleSize + ModuleSize, TURN_STEPS
                    Case WIDE_LEFT
                        pvAppendLeftTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize, lY * ModuleSize + ModuleSize, WIDE_STEPS
                    Case WIDE_RIGHT
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize, lY * ModuleSize + ModuleSize, WIDE_STEPS
                    Case SUPERWIDE_RIGHT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize, lY * ModuleSize + ModuleSize \ 2
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize * 3 \ 2, lY * ModuleSize + 2 * ModuleSize, SUPERWIDE_STEPS
                    End Select
                    lY = lY + 1
                ElseIf .Left <> 0 Then
                    lValue = .Left
                    .Left = 0
                    Select Case lValue
                    Case TURN_LEFT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize - ModuleSize \ 2, lY * ModuleSize
                        pvAppendLeftTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize, lY * ModuleSize + ModuleSize \ 2, TURN_STEPS
                    Case TURN_RIGHT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize - ModuleSize \ 2, lY * ModuleSize
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize, lY * ModuleSize - ModuleSize \ 2, TURN_STEPS
                    Case WIDE_LEFT
                        pvAppendLeftTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize, lY * ModuleSize + ModuleSize, WIDE_STEPS
                    Case WIDE_RIGHT
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize, lY * ModuleSize - ModuleSize, WIDE_STEPS
                    Case SUPERWIDE_RIGHT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize - ModuleSize \ 2, lY * ModuleSize
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize - 2 * ModuleSize, lY * ModuleSize - ModuleSize * 3 \ 2, SUPERWIDE_STEPS
                    End Select
                    lX = lX - 1
                ElseIf .Top <> 0 Then
                    lValue = .Top
                    .Top = 0
                    Select Case lValue
                    Case TURN_LEFT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize, lY * ModuleSize - ModuleSize \ 2
                        pvAppendLeftTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize \ 2, lY * ModuleSize - ModuleSize, TURN_STEPS
                    Case TURN_RIGHT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize, lY * ModuleSize - ModuleSize \ 2
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize \ 2, lY * ModuleSize - ModuleSize, TURN_STEPS
                    Case WIDE_LEFT
                        pvAppendLeftTurn uPoints, lPos, lSize, lX * ModuleSize - ModuleSize, lY * ModuleSize - ModuleSize, WIDE_STEPS
                    Case WIDE_RIGHT
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize, lY * ModuleSize - ModuleSize, WIDE_STEPS
                    Case SUPERWIDE_RIGHT
                        pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize, lY * ModuleSize - ModuleSize \ 2
                        pvAppendRightTurn uPoints, lPos, lSize, lX * ModuleSize + ModuleSize * 3 \ 2, lY * ModuleSize - 2 * ModuleSize, SUPERWIDE_STEPS
                    End Select
                    lY = lY - 1
                Else
                    lValue = -1
                End If
            End With
            If lValue < 0 Then
                Exit Do
            ElseIf lValue = LINE_TO Then
                pvAppendLineTo uPoints, lPos, lSize, lX * ModuleSize, lY * ModuleSize
            End If
        Loop
        If UBound(aSizes) < lNumPolys Then
            ReDim Preserve aSizes(0 To 2 * UBound(aSizes)) As Long
        End If
        aSizes(lNumPolys) = lSize
        lNumPolys = lNumPolys + 1
        lPos = lPos + lSize
        lSize = 0
    Loop
    ReDim Preserve aSizes(0 To lNumPolys - 1) As Long
    ReDim Preserve uPoints(0 To lPos - 1) As POINTAPI
End Sub

Private Function pvFindStartVector(uVectors() As RECT, lX As Long, lY As Long) As Boolean
    For lY = 0 To UBound(uVectors, 2)
        For lX = 0 To UBound(uVectors, 1)
            With uVectors(lX, lY)
                If .Left <> 0 And .Left <> SKIP_TO Or .Top <> 0 And .Top <> SKIP_TO _
                        Or .Right <> 0 And .Right <> SKIP_TO Or .Bottom <> 0 And .Bottom <> SKIP_TO Then
                    pvFindStartVector = True
                    Exit Function
                End If
            End With
        Next
    Next
End Function

Private Sub pvAppendRightTurn(uPoints() As POINTAPI, ByVal lPos As Long, lSize As Long, ByVal lX As Long, ByVal lY As Long, ByVal lSteps As Long)
    Dim lIdx            As Long
    Dim lPrevX          As Long
    Dim lPrevY          As Long
    Dim lCenterX        As Long
    Dim lCenterY        As Long
    Dim dblStartAngle   As Double
    Dim lTempX          As Long
    Dim lTempY          As Long
    
    Debug.Assert lSize > 0
    lPrevX = uPoints(lPos + lSize - 1).x
    lPrevY = uPoints(lPos + lSize - 1).Y
    If lPrevX < lX And lPrevY < lY Then
        lCenterX = lPrevX
        lCenterY = lY
        dblStartAngle = M_PI_2 * 3
    ElseIf lPrevX < lX And Not lPrevY < lY Then
        lCenterX = lX
        lCenterY = lPrevY
        dblStartAngle = M_PI_2 * 2
    ElseIf Not lPrevX < lX And lPrevY < lY Then
        lCenterX = lX
        lCenterY = lPrevY
        dblStartAngle = M_PI_2 * 0
    ElseIf Not lPrevX < lX And Not lPrevY < lY Then
        lCenterX = lPrevX
        lCenterY = lY
        dblStartAngle = M_PI_2 * 1
    End If
    For lIdx = 0 To lSteps
        lTempX = Int(lCenterX + Abs(lX - lPrevX) * Cos(dblStartAngle + M_PI_2 * lIdx / lSteps) + 0.5)
        lTempY = Int(lCenterY + Abs(lY - lPrevY) * Sin(dblStartAngle + M_PI_2 * lIdx / lSteps) + 0.5)
        pvAppendLineTo uPoints, lPos, lSize, lTempX, lTempY
    Next
    Debug.Assert lTempX = lX And lTempY = lY
End Sub

Private Sub pvAppendLeftTurn(uPoints() As POINTAPI, ByVal lPos As Long, lSize As Long, ByVal lX As Long, ByVal lY As Long, ByVal lSteps As Long)
    Dim lIdx            As Long
    Dim lPrevX          As Long
    Dim lPrevY          As Long
    Dim lCenterX        As Long
    Dim lCenterY        As Long
    Dim dblStartAngle   As Double
    Dim lTempX          As Long
    Dim lTempY          As Long
    
    Debug.Assert lSize > 0
    lPrevX = uPoints(lPos + lSize - 1).x
    lPrevY = uPoints(lPos + lSize - 1).Y
    If lPrevX < lX And lPrevY < lY Then
        lCenterX = lX
        lCenterY = lPrevY
        dblStartAngle = M_PI_2 * 2
    ElseIf lPrevX < lX And Not lPrevY < lY Then
        lCenterX = lPrevX
        lCenterY = lY
        dblStartAngle = M_PI_2 * 1
    ElseIf Not lPrevX < lX And lPrevY < lY Then
        lCenterX = lPrevX
        lCenterY = lY
        dblStartAngle = M_PI_2 * 3
    ElseIf Not lPrevX < lX And Not lPrevY < lY Then
        lCenterX = lX
        lCenterY = lPrevY
        dblStartAngle = M_PI_2 * 0
    End If
    For lIdx = 0 To lSteps
        lTempX = Int(lCenterX + Abs(lX - lPrevX) * Cos(dblStartAngle - M_PI_2 * lIdx / lSteps) + 0.5)
        lTempY = Int(lCenterY + Abs(lY - lPrevY) * Sin(dblStartAngle - M_PI_2 * lIdx / lSteps) + 0.5)
        pvAppendLineTo uPoints, lPos, lSize, lTempX, lTempY
    Next
    Debug.Assert lTempX = lX And lTempY = lY
End Sub

Private Sub pvAppendLineTo(uPoints() As POINTAPI, ByVal lPos As Long, lSize As Long, ByVal lX As Long, ByVal lY As Long)
    If lSize >= 2 Then
        '--- remove redundant last point if the new one is on the same line
        If pvPointsInLine(uPoints(lPos + lSize - 2), uPoints(lPos + lSize - 1), lX, lY) Then
            lSize = lSize - 1
        End If
    End If
    If UBound(uPoints) < lPos + lSize Then
        ReDim Preserve uPoints(0 To 2 * UBound(uPoints)) As POINTAPI
    End If
    With uPoints(lPos + lSize)
        .x = lX
        .Y = lY
    End With
    lSize = lSize + 1
End Sub

Private Function pvPointsInLine(uA As POINTAPI, uB As POINTAPI, ByVal lX As Long, ByVal lY As Long) As Boolean
    If uA.x = lX Then
        pvPointsInLine = uB.x = lX
    ElseIf uA.Y = lY Then
        pvPointsInLine = uB.Y = lY
    Else
        pvPointsInLine = (uA.x - lX) * (uA.Y - lY) = (lX - uB.x) * (lY - uB.Y)
    End If
End Function
' ==============================================================
'  THANKAKA AUTO DRAWING PRO
'  Auto Drawing By Thankaka - TCI 2026
'  Version: 1.0
' ==============================================================

Option Explicit

' ==============================================================
'  KHOI TAO FORM
' ==============================================================
Private Sub UserForm_Initialize()

    On Error Resume Next
    
    ' Mo tab dau tien
    MultiPage1.value = 0
    
    ' ----- TAB 2 -----
    chkQRCode.value = True
    chkAutoSort.value = False
    chkAutoNote.value = False
    chkInsertThread.value = False
    chkPOCode.value = True
    chkRefreshPartlist.value = True
    chkAutoTotalSheets.value = True
    optCurrentSheet.value = False
    txtStartBy.Text = "1"
    
    ' ----- TAB 3 -----
    optSemiStandard.value = True
    optAll.value = True
    chkNewAbsoluteValues.value = False
    chkExcludeStandard.value = False
    
    ' ----- TAB 4 -----
    optDrawingNo.value = True
    chkPDF.value = True
    chkDXF.value = True
    chkDWG.value = False
    chkSTEP.value = True
    cboRev.Text = "R2007"
    optCurrentSheetExport.value = True
    
    On Error GoTo 0
    
End Sub

' ==============================================================
'  TAB 1 - AUTO DRAWING
' ==============================================================
Private Sub btnAutoDrawing3D_Click()
    On Error GoTo ErrorHandler
    Me.Hide
    Call RunAutoDrawing_V35
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show
End Sub

Private Sub btnCompleteDrawing2D_Click()
    On Error GoTo ErrorHandler
    Me.Hide
    Call RunCompleteDrawing_V45
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show
End Sub
Private Sub btnAutoDrawingDim_Click()
    On Error GoTo EH
    Me.Hide
    ' Phase 1 (import marked) - neu da co modAP_AutoDim
    ' Call RunAutoDrawingDim
    ' Phase 2 (dim bao, khong can 3D dim)
    Call RunAutoDrawingDimPhase2
    Me.Show
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show
End Sub
Private Sub btnCompleteDrawingDim_Click()
    On Error GoTo EH
    Me.Hide
    Call RunCompleteDrawing_V45
    'Call RunAutoDrawingDim
    Me.Show
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show
End Sub
Private Sub btnDelDrawingDim_Click()
    On Error GoTo EH
    Me.Hide
    Call DeleteDimsCurrentSheet
    Me.Show
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show
End Sub
Private Sub btnQuickTol_Click()
    Me.Hide
    On Error GoTo EH
    modAP_DeleteDangling.DeleteDanglingDims
    Me.Show vbModeless
    Exit Sub
EH:
    Me.Show vbModeless
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
End Sub
Private Sub btnDimCurrentSheet_Click()
    On Error GoTo EH
    Me.Hide
    Call modAutoDim_NoTable.RunAutoDimCurrentSheet
    Me.Show vbModeless
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show vbModeless
End Sub
Private Sub btnAutoDrawingDimNoTable_Click()
    On Error GoTo EH
    
    Me.Hide
    Call RunAutoDrawingDimPhase2_NoTable
    Me.Show vbModeless
    Exit Sub
    
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show vbModeless
End Sub
Private Sub btnDimSheetTable_Click()
    On Error GoTo EH
    Me.Hide
    Call RunAutoDimCurrentSheet_Table
    Me.Show vbModeless
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show vbModeless
End Sub
Private Sub btnDimSmart_Click()
    On Error GoTo EH
    Me.Hide
    Call RunAutoDimSmart
    Me.Show vbModeless
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show vbModeless
End Sub
Private Sub btnDimOrdinateAll_Click()
    On Error GoTo EH
    Me.Hide
    Call RunOrdinateDimAllSheets
    Me.Show vbModeless
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show vbModeless
End Sub
Private Sub btnDimOrdinateCurrent_Click()
    On Error GoTo EH
    Me.Hide
    Call RunOrdinateDimCurrentSheet
    Me.Show vbModeless
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show vbModeless
End Sub
Private Sub btnAutoDrawingPart_Click()
    Call modAutoDrawingPart.AutoDrawingPart
End Sub
Private Sub btn3DSelect_Click()
    Me.Hide
    Call TCI_AutoDrawingSelect
    Me.Show
End Sub

' ==============================================================
'  CAC NUT EXIT
' ==============================================================
Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnExit2_Click()
    Unload Me
End Sub

Private Sub btnExit3_Click()
    Unload Me
End Sub

Private Sub btnExit4_Click()
    Unload Me
End Sub

Private Sub btnExit5_Click()
    Unload Me
End Sub

Private Sub btnExit6_Click()
    Unload Me
End Sub

Private Sub btnExitFooter_Click()
    Unload Me
End Sub

' ==============================================================
'  TAB 2 - AUTO PARTLIST 2D (tam thoi)
' ==============================================================
Private Sub btnAutoPartlist2D_Click()
    On Error GoTo ErrorHandler
    Me.Hide
    Call RunAutoPartlist
    Me.Show
    Exit Sub
ErrorHandler:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show
End Sub

' ==============================================================
'  TAB 3 - QUANTITY (tam thoi)
' ==============================================================
Private Sub btnAutoQty_Click()
    On Error GoTo ErrorHandler
    Me.Hide
    Call RunAutoQuantity
    Me.Show
    Exit Sub
ErrorHandler:
    MsgBox "Loi AUTO QTY: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show
End Sub

Private Sub btnExcel_Click()
    ' Tam thoi chi ho tro ALL
    If optAll.value = True Then
        Call ExportTBTC_ALL
    Else
        MsgBox "Hien tai chi ho tro che do ALL (TBD + TBC + TBTC).", vbInformation, "Thankaka"
    End If
End Sub

Private Sub btnCompare_Click()
    Call CompareQuantity
End Sub

' ==============================================================
'  TAB 4 - EXPORT FILES (tam thoi)
' ==============================================================
Private Sub btnSaveAsFile_Click()
    On Error GoTo ErrorHandler
    Me.Hide
    Call RunExportFiles
    Me.Show
    Exit Sub
ErrorHandler:
    MsgBox "Loi Export: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show
End Sub

' ==============================================================
'  TAB 5 - MANUFACTURING (tam thoi)
' ==============================================================
   Private Sub btnCheckExcessDrawings_Click()
       On Error GoTo ErrorHandler
       Me.Hide
       frmExcessDrawings.Show
       Me.Show
       Exit Sub
ErrorHandler:
       MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
       Me.Show
   End Sub
      Private Sub btnCheckMissingDrawings_Click()
       On Error GoTo ErrorHandler
       Me.Hide
       frmMissingDrawings.Show
       Me.Show
       Exit Sub
ErrorHandler:
       MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
       Me.Show
   End Sub
      Private Sub btnCheckProperties_Click()
       On Error GoTo ErrorHandler
       Me.Hide
       frmCheckProperties.Show
       Me.Show
       Exit Sub
ErrorHandler:
       MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
       Me.Show
   End Sub

Private Sub btnSetSheetScale_Click()
    MsgBox "Chuc nang SET SHEET SCALE se duoc phat trien sau." & vbCrLf & vbCrLf & _
           "Auto Drawing By Thankaka - TCI 2026", vbInformation, "Thankaka"
End Sub

Private Sub btnViewMass_Click()
    MsgBox "Chuc nang VIEW MASS se duoc phat trien sau." & vbCrLf & vbCrLf & _
           "Auto Drawing By Thankaka - TCI 2026", vbInformation, "Thankaka"
End Sub
Private Sub btnAutoBolt_Click()
    Me.Hide
    On Error GoTo EH
    modAutoBolt.RunAutoBolt
    Me.Show vbModeless
    Exit Sub
EH:
    Me.Show vbModeless
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
End Sub
Private Sub btnAdmin_Click()
    ThankakaLicense.ShowAdminLicense
End Sub

' ==============================================================
'  TAB 6 - SMART ASM & ADV (tam thoi)
' ==============================================================
Private Sub btnFrameJoint_Click()
    On Error GoTo EH
    Me.Hide
    Call RunFrameJoint
    Me.Show vbModeless
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description, vbCritical, "Thankaka"
    Me.Show vbModeless
End Sub

Private Sub btnInsertPin_Click()
    MsgBox "Chuc nang INSERT PIN se duoc phat trien sau." & vbCrLf & vbCrLf & _
           "Auto Drawing By Thankaka - TCI 2026", vbInformation, "Thankaka"
End Sub







### 2. SOLIDWORKS API Code Examples

| API Category | Reference |
|---|---|
| Assembly Traversal | https://www.codestack.net/solidworks-api/document/assembly/components/traversing-tree/ |
| Component Management | https://www.codestack.net/solidworks-api/document/assembly/components/ |
| Drawing Automation | https://www.codestack.net/solidworks-api/document/drawing/ |
| Dimension Traversal | https://www.codestack.net/solidworks-api/document/dimensions/traverse-all/ |
| Drawing View Scaling | https://www.codestack.net/solidworks-api/document/drawing/views-size-based-scale/ |
| Custom Properties | https://www.codestack.net/solidworks-api/data-storage/custom-properties/ |
| Document Management | https://www.codestack.net/solidworks-api/application/documents/ |

### 3. AI & MCP Automation

| Project | Main Features | Repository |
|---|---|---|
| SW_MCP | Python COM automation and AI tool integration | https://github.com/d-veloping/SW_MCP |
| SolidWorks MCP | Natural language CAD automation | https://github.com/alisamsam/solidworks-mcp |
| SolidWorks Automation Skill | AI Skill, MCP, verification and execution tracing | https://github.com/wzyn20051216/solidworks-automation-skill |

### 4. .NET Development Frameworks

| Project | Main Features | Repository |
|---|---|---|
| xCAD.NET | SOLIDWORKS Add-ins, Property Manager Pages and Macro Features | https://github.com/xarial/xcad |
| xCAD Examples | SOLIDWORKS API implementation examples | https://github.com/xarial/xcad-examples |

### 5. Compatibility Policy

- Primary development target: SOLIDWORKS 2019.
- Prefer native VBA macros and documented SOLIDWORKS API methods.
- Newer API examples must be reviewed and adapted before integration.
- C#, VB.NET and Python projects are reference implementations, not directly executable VBA macros.
- Every integrated module must pass compilation and functional testing in SOLIDWORKS 2019.
- Preserve existing modules and avoid deleting existing drawing dimensions or model data.

**Important:** Repository compatibility claims do not replace actual testing in SOLIDWORKS 2019.

