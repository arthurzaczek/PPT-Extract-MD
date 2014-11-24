Extracts power point slides to Markdown using powershell

**ITS UNDER DEVELOPMENT!**

	C:\PS> Get-Help .\PPT-Extract-MD.ps1 -detailed

	NAME
		PPT-Extract-MD.ps1

	SYNOPSIS
		Extracts power point slides to Markdown js


	SYNTAX
		PPT-Extract-MD.ps1 [[-file] <String>] [-SourceCode] [-Open] [<CommonParameters>]


	DESCRIPTION
		Supports syntax highlighing


	PARAMETERS
		-file <String>

		-SourceCode [<SwitchParameter>]

		-Open [<SwitchParameter>]

		-------------------------- EXAMPLE 1 --------------------------

		C:\PS> .\PPT-Extract-MD.ps1 'MyPresentation.pptx'


		-------------------------- EXAMPLE 2 --------------------------

		C:\PS> ls *.pptx | % { .\PPT-Extract-MD.ps1 $_ }