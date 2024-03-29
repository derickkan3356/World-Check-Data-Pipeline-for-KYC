Sub Button_onAction(control As IRibbonControl)
    Select Case control.ID
		Case "button_ImportWCdata"
			Call ImportWCdata
		Case "button_ExportWCtemplate"
			Call ExportWCtemplate
		Case "button_ExportWCdata"
			Call ExportWCdata
    End Select
End Sub