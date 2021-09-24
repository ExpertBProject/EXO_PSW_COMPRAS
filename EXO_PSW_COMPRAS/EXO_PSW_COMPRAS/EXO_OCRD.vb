Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OCRD
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
    End Sub

    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_VALIDATE_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sCardCode As String = "" : Dim sCardName As String = "" : Dim sCardCodeAct As String = ""
        Dim sMensaje As String = ""
        EventHandler_VALIDATE_Before = False
        Try

            If pVal.ItemUID = "43" Or pVal.ItemUID = "45" Or pVal.ItemUID = "51" Then 'Teléfonos pestaña general
                Dim sPais As String = Left(CType(oForm.Items.Item("41").Specific, SAPbouiCOM.EditText).Value.ToString, 2)
                Dim sTelefono As String = CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).Value.ToString
                sCardCodeAct = CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value.ToString
                If sTelefono.Trim <> "" Then
                    Select Case sPais
                        Case "ES"
                            If Not IsNumeric(sTelefono) Then
                                sMensaje = "El campo de teléfono debe ser numérico y sin espacios."
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox(sMensaje)
                                Exit Function
                            ElseIf Left(sTelefono, 2) = "00" Or Left(sTelefono, 2) = "34" Then
                                sMensaje = "El campo de teléfono no puede comenzar por ""00"" o ""34""."
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox(sMensaje)
                                Exit Function
                            End If
                        Case Else
                            If Not IsNumeric(sTelefono) Then
                                sMensaje = "El campo de teléfono debe ser numérico y sin espacios."
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox(sMensaje)
                                Exit Function
                            ElseIf Left(sTelefono, 2) <> "00" Then
                                sMensaje = "El campo de teléfono debe comenzar por ""00"" seguido del código del país."
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox(sMensaje)
                                Exit Function
                            End If
                    End Select

                    'Ahora buscamos si ya existe el teléfono en otro IC.
                    sCardCode = objGlobal.refDi.SQL.sqlStringB1("SELECT ""CardCode"" FROM ""OCRD"" WHERE ""Phone1""='" & sTelefono & "' and ""CardCode""<>'" & sCardCodeAct & "' ")
                    If sCardCode.Trim = "" Then
                        sCardCode = objGlobal.refDi.SQL.sqlStringB1("SELECT ""CardCode"" FROM ""OCRD"" WHERE ""Phone2""='" & sTelefono & "' and ""CardCode""<>'" & sCardCodeAct & "' ")
                        If sCardCode.Trim = "" Then
                            sCardCode = objGlobal.refDi.SQL.sqlStringB1("SELECT ""CardCode"" FROM ""OCRD"" WHERE ""Cellular""='" & sTelefono & "' and ""CardCode""<>'" & sCardCodeAct & "' ")
                        End If
                    End If

                    If sCardCode.Trim <> "" Then
                        sCardName = objGlobal.refDi.SQL.sqlStringB1("SELECT ""CardName"" FROM ""OCRD"" WHERE ""CardCode""='" & sCardCode & "' ")
                        sMensaje = "El teléfono ya existe en el IC: " & sCardCode & " " & sCardName & ". No se puede repetir. "
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.SBOApp.MessageBox(sMensaje)
                        Exit Function
                    End If


                End If
            End If

            EventHandler_VALIDATE_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
End Class
