Sub exibeOculta()
	'@author=jonathan_conzatti
	'Macro que verifica se deve exibir ou ocultar as guias e cabeçalhos
	'Exemplo: Exibir/ocultar guias e cabeçalhos ocultos
	'Conceitos importantes: Somente seleção
	Dim planilhaAtiva As String
	Dim status As Boolean
	
	Application.ScreenUpdating = False
	planilhaAtiva = ActiveSheet.Name
	status = Not ActiveWindow.DisplayHeadings
	
	For Each Nm In Worksheets
			Nm.Activate
				ActiveWindow.DisplayHeadings = status
				ActiveWindow.DisplayWorkbookTabs = status
				ActiveWindow.DisplayWorkbookTabs = status
	Next
	
	Application.ScreenUpdating = True
	Sheets(planilhaAtiva).Select
End Sub