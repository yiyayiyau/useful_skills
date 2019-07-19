sFolder = "C:\Users\hengw\Downloads\BildgAutom\books\"
Set fileObject = CreateObject("Scripting.FileSystemObject")
For Each file In fileObject.GetFolder(sFolder).Files
	If LCase(fileObject.GetExtensionName(file.Name)) = "pdf" Then
		CreateObject("WScript.Shell").Run(file)
	End if
Next