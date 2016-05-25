

*/--------------------------------------------------------------------------------------------------------
*/ Description..: O FoxcodePlus não substituí o IntelliSense do VFP, ele interage em pontos que o
*/				  IntelliSense padrão do VFP não ajuda adequadamente ou não faz absolutamente nada.
*/				  A ideia do FoxcodePlus é trazer um pouco da funcionalidade do IntelliSense do Visual Studio
*/				  para o VFP, ou seja, dar mais agilidade e evitar erros ao escrever os programas.
*/ Updated By: mInternauta - Marcelo Andrade Jr. - Campinas SP - Brazil (Clic Sistemas Ltda)
*/ By Author.......: Rodrigo Duarte Bruscain - São Paulo SP - Brazil | Kitchener ON - Canada
*/ Date start...: May 01, 2010
*/--------------------------------------------------------------------------------------------------------
*/
*/ BETA 3.15 - 10/03, 2016
*/--------------------------------------------
* NEW: Mudado de LOCFILE para GETFILE, evitando que um erro seja disparado pelo IntelliSense
*
*/ BETA 3.15 - 09/03, 2016
*/--------------------------------------------
* NEW: Adicionado a opção para perguntar o caminho do arquivo caso o IntelliSense não achar
* NEW: Novos Icones para o IntelliSense
* NEW: Patch de Icones para o FoxPro
*
*/ BETA 3.14 - ??? ??, 2013
*/--------------------------------------------
* NEW: IntelliSense para propriedades em classes PRG checando o metodo atual e o INIT se tem propriedade recebendo objeto
* NEW: IntelliSense para propriedades em FORMS checando o metodo atual e o INIT se tem propriedade recebendo objeto
* NEW: Assinatura dos user metodos em run-time. (consigo somente qdo a assinatura esta na mesma linha, mudar this.GetMembers())
* NEW: Seleção de captionlization para campos e fields
* FIX: Tratar erro qdo esta conectado a outro database diferente de SQL. (verificar se tem como usar outros database)
*
*

*/
*/ BETA 3.13 - May 4, 2013
*/--------------------------------------------
* NEW: SQL Intellisense	(valid only inside TEXT...ENDTEXT)
*		OK - Tables with incremental search in current database connected
*		OK - Table and Alias with incremental searching in currente SQL instruction with the database disconnected
*		OK - Fields with incremental searching in current database connected considering the tables used in currente SQL instruction
*		OK - Fields list from tables and Alias when pressed "." dot
*		OK - IntelliSense for tables/alias pressing SPACE after the clauses FROM, JOIN and INTO

* NEW: Intellisense to the command "INDEX ON"
* NEW: Now, objects instatied at run-time show more informations in the tooltip.
* NEW: Intellisense for object created by "FOR EACH" at run-time and designer-time
* NEW: Intellisense for Collection object at run-time and designer-time				(ex: thisform.grid1.columns[1].)
* NEW: Referencing any object through of a variable at run-time and designer-time	(ex: loObj = _screen  /  loObj = thisform.grid1.column1)
* NEW: FoxCode table update through the merger between native FoxCode and FoxCode provided by FoxcodePlus.

* FIX: SHIFT+ARROWS, SHIFT+END and SHIFT+HOME behavior not adequate
* FIX: "THIS" in containers object in some cases the IntelliSense hasn't load.



*/ BETA 3.12 - April 21, 2013
*/--------------------------------------------
* NEW: Now, Create Table <tablename> and Create DBF <tablename> shown tables and fields in write-time.
* NEW: Possibility to select an item in IntelliSense pressing "*" or "/"
* NEW: IntelliSense to the comand REPLACE
* FIX: _MemberData Capitalization for objects at run-time.
* FIX: Thisform in form designer occur an error when the object has a property "Caption" with data type # of char.
* FIX: "Error List Window" doesn't work when activated directly from "View" menu
* FIX: In select sql-command the clause "into ..." now is dismembered
* FIX: Sometimes an error happens at run-time when the IntelliSense try to open the list
* FIX: CTRL+S conflict
* FIX: OleControl with anonymous object


*/ BETA 3.11 - April 03, 2013
*/--------------------------------------------
* NEW: Support for Windows 8
* NEW: Like in Visual Studio, the native or custom Code Snippets can be shown in the IntelliSense.
* NEW: Now, VFP can read custom Code Snippets for functions (native IntelliSense is only for commands)
* NEW: Several Code Snippets were included for commands and functions
* NEW: In the IntelliSense Manager there is an option to increment the IntelliSense only with keywords started by typed
* NEW: In the IntelliSense Manager there is an option to include the Code Snippets in IntelliSense
* NEW: In the IntelliSense Manager there is an option to replace the native IntelliSense in form and class desinger
* NEW: Now, when an object in a form and class designer is selected in incremental IntelliSense, the hierarchy is included
* NEW: IntelliSense to the command "COPY TO <MyFileName> TYPE"
* NEW: IntelliSense to the command "SET PROCEDURE TO <listfiles>"
* NEW: Now, IntelliSense can read classes and functions invoked by "SET PROCEDURE TO"
* NEW: IntelliSense to the command "SET CLASSLIB TO <filename>"
* NEW: Now, IntelliSense can read classes invoked by "SET CLASSLIB TO"
* NEW: IntelliSense to the command "DO FORM <filename>"
* NEW: IntelliSense to the command "REPORT FORM <filename>"
* NEW: IntelliSense to the command "USE <filename>"
* NEW: Now, IntelliSense is showing in the tooltip more information about the controls in forms and class designer
* NEW: Now, tooltip with signature from custom procedures and custom methods are shown. In addiction, number of parameters are controlled and summary is supported.
* NEW: Now, API functions at write-time show the signature.
* NEW: Possibility to select an item in IntelliSense pressing "," or "#"
* FIX: Typing "." in properties documentation using "&&&"
* FIX: Invalid subscript error to obtain the summary
* FIX: VFP freeze when exist _MemberData property in _Screen object
* FIX: Clauses in command "USE" in the native IntelliSense VFP have don't work correctly when table or alias name contain the name one of clauses
* FIX: SET PATH at run-time is not considered when a file is referenced at write-time
* FIX: IntelliSense doesn't work with properly with commands like SELECT SQL with ";" used to break lines
* FIX: Incremental IntelliSense doesn't work in "#Preprocessor Directive"
* FIX: "Local Array laArrayName[10]" in IntelliSense shows "Array laArrayName" and should be "laArrayName"


*/ BETA 3.XX
*/----------------------------------------
* FIX: Clause "IN" in several command doesn't open IntelliSense for tables.
* FIX: Added tables from DataEnviroment when used clause "IN" in several commands.
* FIX: Typing "." outside of the editor in some places.
* FIX: Typing "=" outside of the editor in some places.
* FIX: Sometimes an error happend to obtain constants from file .H
* FIX: API Exception ... more modifications to avoid this problem.
* FIX: FoxcodePlus has not considered the "Error Tip" when unchecked in IntelliSense Manager.
* XXX FIX: Code snippets inserted in wrong places

* NEW: IntelliSense for classes in PRG file when an object is instantiated with CreateObject() and NewObject() in the same PRG.
* NEW: Shows the FoxcodePlus version in IntelliSense Manager and in error messagebox.
* NEW: More informations if error happens in FoxcodePlus and now a file called foxcodeplus.err is generated.
* FIX: "m." wasn't work.
* FIX: Sometimes some tables in DataEnvironment weren't shown.
* FIX: Automatic selection erases contents when typing "." or "=" after close with ")" or "]"
* FIX: Restore default VFP fonts if uncheck option "Visual Studio font style" in IntelliSense Mananger.
* FIX: Sometimes when select an item from a class member or table field, they are inserted in wrong place.
* FIX: When included the path of PRG or VCX file when used Createobject() or NewObject(), sometimes the IntelliSense wasn't work.
* FIX: Possibility to select an item in IntelliSense typing with "<" or ">" or "+" or "-"

* FIX: Possibility to select an item in IntelliSense closing the parenthesis ")".
* FIX: When a procedure was incluided in a classe in PRG file sometimes a bug happend.
* FIX: Some changes in IntelliSense list to avoid the "API Call" error.
* FIX: In some cases tooltip for members class doesn't appear when the IntelliSense was opened.
* FIX: Duplicated items in IntelliSense for Sql Command.

* NEW: Included more commands to SQL Command IntelliSense
* FIX: Constants files .H with long directory + long file name .H
* FIX: Use more than one constants files .H in the same PRG, Method or Function
* FIX: Adjusted error message to possibly copy and paste the error
* FIX: DataEnviroment with relation object
* FIX: Capitalization / Expansion now can use Foxcode Default option
* FIX: Tables in DataEnvironment using CursorAdapter

* FIX: SET EXACT has been turned "ON" even when turned "OFF"
* NEW: Error list has included in default menu VIEW with HotKey to activate "Error List" window.
* NEW: Incremental IntelliSense for tables defined in DataEnvironment Forms and Reports.
* NEW: IntelliSense for fields in DataEnvironment Forms and Reports.
* NEW: IntelliSense for command "REPORT FORM"
* FIX: In class in prg file, methods inherited from class has repeted if the same method has included by developer.
* FIX: ? or ?? has not sent the message to screen when necessary.
* FIX: Canceled selection to procedural functions or commands when typing "."
* FIX: In Command window with "Error List" activated when execute the command "Clear", "Error list" was cleared.
* FIX: Setting breakpoints on prg editor triggers 'invalid function argument type or count' when the debugger is auto-opened for the first time.
* FIX: When an item was positioned using key up, down, pgup or pgdw, the item wans't selected.

*/ BETA 2 - 2012.11.10
*/------------------------------------
*/ FIX: Fast typing
*/ FIX: Picos de consumo do processador
*/ FIX: When another FLL library was called without ADDITIVE clause FoxcodePlus crashed.
*/ FIX: Objetos selecionados na lista do foxcodeplus teclando "." agora são selecionados. (OBS: In Default VFP IntelliSense doesn't work yet.)
*/ FIX: Quando estou usando o Report, nas propriedades do objeto, Aba "Print When", Propriedade "Print only when expression is true:" pressionar "="
*/ FIX: Ao criar um prg pela 1a. vez o IntelliSense nao abre, se salvar e abri-lo funciona.
*/ FIX: Quando pressiono "(" o item da lista do IntelliSense nao é completado.
*/ FIX: Error list quando dockado em outro janela ocorre erro de dimensionamento na coluna 3
*/ FIX: Visual Studio Colors Style background yellow string was removed.
*/ NEW: Error list para forms e classes
*/ NEW: Dokagem com redimensionamento do error list.
*/ NEW: _memberdata para objetos instanciados em runtime (OBS -> Sem suporte para _memberdata protected or hidden para objetos instanciados). As propriedade
*/ NEW: Ler conteudo de arquivo .H em forms e PRGs
*/ NEW: #DEFINE e #INCLUDE em forms nao aparece
*/ NEW: Em controles visuais o tooltip apresenta o caption de controles que tem a propriedade caption.
*/ NEW: IntelliSense para o comando alter table
*/ NEW: _Tally now is present in incremental IntelliSense
*/ NEW: Tooltip de erro de programação em write-time com parametro no foxcode.app

*/ BETA 1 - 2012.10.10
*/------------------------------------
*/ FIX: ESC cancelava quando estava dentro da lista do IntelliSense.
*/ FIX: Correção de apresentacao incorreta no IntelliSense qdo uma propriedade em write-time é tipada
*/ FIX: Correção do conflito com o Debugger. Agora o IntelliSense nao é apresentado enquanto o Debugger estiver aberto.
*/ FIX: Corrigido comportamento para desenvolvedores que usam o VFP em background. Estava apresentando informações no _screen indevidamente.
*/ NEW: IntelliSense agora identifica no tooltip as PEMs que são Read Only.
*/ FIX: Corrigido o metodo this.GetMembers() que não trazia algumas users pmes.
*/ FIX: Corrigido o metodo this.GetDot() que estava considerando linha comentada
*/ FIX: Tratamento de erros do FoxcodePlus (ninguem faz tudo 100% rsrsrsrs)
*/ NEW: IntelliSense de variaveis valorizadas por outras variaveis que contem createobjet(), newobject(), createobjectex()
*/ NEW: Disponibilizado IntelliSense para createobjet(), newobject(), createobjectex() em write-time
*/ NEW: IntelliSense abria dentro de textos o que atrapalhava a digitacao livre.
*/ NEW: IntelliSense abria dentro do Text...EndText, agora é abre somente dentro dos delimiters.
*/--------------------------------------------------------------------------------------------------------


*/ WISHLIST
*/--------------------------------------------------------------------------------------------------------
*- ERROR: funcoes que abrem o IntelliSense do vfp estão em conflito com o foxcodeplus ... set(), adir(), etc.
*- ERROR: Erro no "code zoom" do relatorio em "propriedades" (editor do vfp aberto em outra janela)
*- MELHORIA: Salvar dockagem da "Error List"



*/----------------------------------------------------------------------------------------------------
*/ Starting FoxcodePlus
*/----------------------------------------------------------------------------------------------------
#Define WM_KEYUP	0x0101
#Define WM_DESTROY	0x0002
#Define CRLF 		Chr(13) + Chr(10)

External Array plaCodePrg
External Array plaItemsVars

Set Console Off

If Not File(Home(1)+"foxtools.fll") Or Not Val(Substr(Version(4),1,2)) >= 9
	Return .F.
Endif

If Not File(Home(1)+"foxcodeplus.ini")
	Return .F.
Endif

Use (_Foxcode) Alias ___chkFoxcodeVersion In 0 Shared Again
Select ___chkFoxcodeVersion

Locate For "foxcodeplus" $ Lower(Data)
If Found()
	Use In ___chkFoxcodeVersion
Else
	Messagebox("FoxcodePlus has not been installed correctly."+Chr(13)+"Go to the IntelliSense Manager to update your FoxCode table.",16,"FoxcodePlus")
	Use In ___chkFoxcodeVersion
	Return .F.
Endif

Set Message To "Loading FoxcodePlus..."

With _Screen
	.AddProperty("FoxcodePlus",.Null.)
	.FoxCodePlus = Createobject("FoxcodePlusMain")
Endwith

Set Message To
Set Console On

*- Definicao do menu "Error List" no menu "View" do VFP
Define Bar 2 Of _MVIEW Prompt "\<Error List" Message "Show Error List at write-time" Key Alt+K,"Alt+K" Picture Home(1)+"foxcodeplus\images\error_list_menu.bmp" Skip For Type("_screen.foxcodeplus.EditorSource")<>"N"
On Selection Bar 2 Of _MVIEW _Screen.FoxCodePlus.showerrorlist()


** Funções para o On Key Label
Function fcp_OnDot
	If Type("_Screen.FoxCodePlus") == "O" .And. Not Isnull(_Screen.FoxCodePlus) Then
		_Screen.FoxCodePlus.GetDot()
	Else
		Keyboard "." Plain
	Endif
Endfunc

Function fcp_OnEqual
	If Type("_Screen.FoxCodePlus") == "O" .And. Not Isnull(_Screen.FoxCodePlus) Then
		_Screen.FoxCodePlus.GetEqual()
	Else
		Keyboard "=" Plain
	Endif
Endfunc





*/----------------------------------------------------------------------------------------------------
*/ FoxcodePlus Class
*/----------------------------------------------------------------------------------------------------
Define Class FoxCodePlusMain As Custom
	IntelliSense = .Null.							&&& objeto com a lista de itens do IntelliSense do foxcodeplus
	*FoxTools = .null.								&&& classe para maipulacao do editor da IDE do VFP.
	ToolbarProcInfo = .Null.						&&& objeto com a toolbar com informações do programa aberto
	FormErrorList = .Null.							&&& objeto com o form que exibe os erros de compilação em write-time
	ToolTip = .Null.								&&& objeto com o tooltip
	LastTopToolTip = 0								&&& somente para auxiliar na apresentação do tooltip
	LastLeftToolTip = 0								&&& somente para auxiliar na apresentação do tooltip
	TextLine = ""									&&& texto da linha corrente
	TextLine2 = ""									&&& texto da linha corrente (anterior a modificacao)
	WordCount = 0									&&& total de palavras na linha corrente
	LastWord = ""									&&& ultima palavra da linha atual (da posicao do ponteiro)
	Lastkey = 0										&&& ultima tecla pressionada
	CursorPos = 0									&&& posicao corrente dentro do texto
	CursorLine = 0									&&& linha atual onde o cursor esta posicionado
	MaxWidth = 0									&&& width do item mais largo incluido no IntelliSense
	*dimension NoIntelliSense[1,3] 					&&& comandos e funcoes que não devem ter IntelliSense do foxcodeplus
	NoIntelliSense = ""								&&& comandos e funcoes que não devem ter IntelliSense do foxcodeplus
	FoxcodeCore = .F.								&&& .T. indica que o IntelliSense do core do vfp esta aberto
	HasDebugger = .F.								&&& .T. indica que o debug do vfp esta aberto
	HasSelectedItem = .F.							&&& controla inserçao do code snippet para comandos e funcoes
	LoadScriptBoolean = .F.							&&& indica que o script "Boolean" do foxcode.dbf e foxcode.app deve ser executado
	EditorSource = -1								&&& editor onde o IntelliSense esta aberto
	EditorFileName = ""								&&& nome do arquivo aberto no editor
	EditorFontName = ""								&&& nome da fonte usada no editor
	EditorFontSize = 0								&&& tamanho da fonte usada no editor
	EditorHwnd = 0									&&& Handle da tela corrente do editor
	EditorToolTip = .Null.							&&& objeto com o tooltip do editor
	Tmpfile = Sys(2023)+"\tmp"+Sys(2015)+".tmp"		&&& nome do arquivo temporario usado no mecanismo do IntelliSense
	ProcClass = ""									&&& nome da classe corrente de um prg em write-time
	ProcBaseClass = ""								&&& nome do baseclass de classe corrente de um prg em write-time
	ControlClassName = ""							&&& nome da classname de um um objeto de controle inserido em um define class de um prg com ADD OBJECTS
	ControlOleClass = ""							&&& nome do Oleclass de um um objeto de controle inserido em um define class de um prg com ADD OBJECTS
	WithReference = ""								&&& armazena a ultima referencia do with/endwith antes de selecionar um item no IntelliSense
	IncrementalResult = .T.							&&& indica que o IntelliSense retorna somente oq foi encontrado oq contem na palavra digitada
	CommandCase = ""								&&& upper, lower or proper to the vfp commands
	FunctionCase = ""								&&& upper, lower or proper to the vfp functions
	HasDot = .F.									&&& indica se tem ou não "." na linha capturada
	IsComment = .F.									&&&	indica que a linha capturada é um comentario
	IsTextEndText = .F.								&&& .T. indica que esta dentro um bloco TEXT...ENDTEXT
	TextEndBlock = ""								&&& bloco de todo o texto do TEXT...ENDTEXT posicionado
	IsSqlIntelliSense = .F.							&&& indica que é uma instrucao SQL dentro de um TEXT...ENDTEXT e por isso abertura do intellisense SQL
	Dimension Items[1,4] 							&&& array with properties, methods, events, procedures, vars, cursos, tables and dlls
	Dimension ItemsTables[1,2]						&&& tabelas encontradas em write-time e o codigo de programa de criação da mesma.
	Dimension ItemsObjects[1,3]						&&& objetos adicionados pelo "define class ... add object" e suas respectivas classes.
	Dimension ItemsAuxVars[1,4]						&&& usado para auxiliar nas variaveis valorizadas por outra variavel (Var = AnotherVar)
	Dimension Environment[9]						&&& array que controla o ambiente da IDE do VFP para o Foxcodeplus
	Dimension ItemsCodeSnippets[1,2]				&&& Itens definidos no foxcode.dbf
	Dimension FoxcodeFunctions[1]					&&& Funcoes contidas no foxcode.dbf

	*- used in foxcodeplus.ini
	chkFC = "1"
	chkTF = "1"
	chkControl = "1"
	chkObj = "1"
	chkVar = "1"
	chkAPI = "1"
	chkFont = "1"
	chkColors = "1"
	chkErrorList = "1"
	chkErrorListDockPos = 3
	chkErrorToolTip = "1"
	chkCodeSnippet = "1"
	chkAutoCloseQuotes = "1"
	cboSearch = "1"
	chkMngDesignTime = "1"
	cboDisplayCount = 10
	chkTFsql = "1"									&&& SQL Server and others
	chkIncrTablesSql = "1"							&&& SQL Server and others
	chkIncrFieldsSql = "1"							&&& SQL Server and others
	chkChkSetPCls =  "1"


	*/------------------------------------------------------------------------------------------------
	*/ inicio o foxcodeplus
	*/------------------------------------------------------------------------------------------------
	Protected Procedure Init
		Set Console Off

		*- carrego as configurações do foxcodeplus
		Local lcSets, lnDisplayCount, lcAlias, lcDefaultCase
		lcSets = Iif(File(Home(1)+"foxcodeplus.ini"), Filetostr(Home(1)+"foxcodeplus.ini"), "")

		lnDisplayCount = Int( Val(Strextract(lcSets,"<cboDisplayCount>","</cboDisplayCount>")) )
		This.cboDisplayCount = Iif(Between(lnDisplayCount,10,15), lnDisplayCount, 10)

		This.chkFC = Strextract(lcSets,"<chkFC>","</chkFC>")
		This.chkTF = Strextract(lcSets,"<chkTF>","</chkTF>")
		This.chkControl = Strextract(lcSets,"<chkControl>","</chkControl>")
		This.chkObj = Strextract(lcSets,"<chkObj>","</chkObj>")
		This.chkVar = Strextract(lcSets,"<chkVar>","</chkVar>")
		This.chkAPI = Strextract(lcSets,"<chkAPI>","</chkAPI>")
		This.chkFont = Strextract(lcSets,"<chkFont>","</chkFont>")
		This.chkColors = Strextract(lcSets,"<chkColors>","</chkColors>")
		This.chkCodeSnippet = Strextract(lcSets,"<chkCodeSnippet>","</chkCodeSnippet>")
		This.chkErrorList = Strextract(lcSets,"<chkErrorList>","</chkErrorList>")
		This.chkErrorToolTip = Strextract(lcSets,"<chkErrorToolTip>","</chkErrorToolTip>")
		This.chkErrorListDockPos = Int( Val( Strextract(lcSets,"<chkErrorListDockPos>","</chkErrorListDockPos>") ) )
		This.chkAutoCloseQuotes = Strextract(lcSets,"<chkAutoCloseQuotes>","</chkAutoCloseQuotes>")
		This.cboSearch = Strextract(lcSets,"<cboSearch>","</cboSearch>")
		This.chkTFsql = Strextract(lcSets,"<chkTFsql>","</chkTFsql>")
		This.chkIncrTablesSql = Strextract(lcSets,"<chkIncrTablesSql>","</chkIncrTablesSql>")
		This.chkIncrFieldsSql = Strextract(lcSets,"<chkIncrFieldsSql>","</chkIncrFieldsSql>")
		This.chkChkSetPCls = Strextract(lcSets,"<chkChkSetPCls>","</chkChkSetPCls>")

		*- objeto que apresenta o IntelliSense
		This.IntelliSense = Newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")

		*- objeto para manipulacao do editor
		*this.FoxTools = newobject("FoxTools","FoxcodeTools.fxp")

		*- objeto para visualização do tooltip da lista de itens do IntelliSense
		This.ToolTip = Newobject("ToolTip","FoxCodeToolTip.fxp")

		*- objeto para visualização do tooltip do editor
		This.EditorToolTip = Newobject("ToolTip","FoxCodeToolTip.fxp")

		*- apresenta os erros em write-time
		If This.chkErrorList = "1"

			This.showerrorlist()
		Endif

		*- se trabalho com a toolbar do Modify Command
		*this.ToolBarProcInfo = newobject("ToolbarProcInfo","FoxCodePlusIntelliSense.vcx")
		*this.ToolBarProcInfo.show()

		*- configuro as teclas abaixo para interagir com o IntelliSense
		On Key Label "." fcp_OnDot()
		On Key Label "=" fcp_OnEqual()


		*- comandos sem IntelliSense plus porque usam o intelisense padrao
		This.NoIntelliSense = "<activate><add><append><build><browse><clear><close><crea><creat><create><copy><deactivate><do form>"+;
			"<define><local><delete><display><drop><hide><keyboard><list><modify><move><on>"+;
			"<prtinfo(><pop><push><release><remove><rename><restore><save><scatter><gather><set><set(><set collate to>"+;
			"<set database to><set date><set order to><set path to><set strictdate to>"+;
			"<set udfparms to><show><size><use><report><zap><protected><hidden><wait>"
		lcAlias = Alias()

		Use (_Foxcode) Again Alias __xfoxcode In 0 Shared
		Select __xfoxcode

		*- comandos e funcoes sem IntelliSense plus porque usam o intelisense padrao
		*!*			select type, iif(type="F", abbrev+"(", expanded))  ;
		*!*			from __xfoxcode ;
		*!*			where ( inlist(type,"F","C") and ;
		*!*					inlist(cmd,"{}","{funcmenu}","{funcmenu2}","{setsysmenu}","{onkeymenu}","{dbgetmenu}","{setmenu}","{onoffmenu}") or ;
		*!*					(cmd="{cmdhandler}" and not empty(data)) ) and ;
		*!*					not ",T" $ strtran(data," ","") and not deleted() ;
		*!*			into array this.NoIntelliSense


		*- foxcode.dbf functions
		Select Expanded From __xfoxcode Where Type = "F " Into Array This.FoxcodeFunctions


		*- Upper, lower or proper to the vfp commands e functions

		Locate For __xfoxcode.Type = "V"							&&- Foxcode Default
		lcDefaultCase = Iif(Found(), __xfoxcode.Case, "")

		Go Top
		Locate For __xfoxcode.Type = "C"							&&- Commands
		This.CommandCase = Iif(Found(), __xfoxcode.Case, "")
		If Empty(This.CommandCase)
			This.CommandCase = lcDefaultCase
		Endif

		Go Top
		Locate For __xfoxcode.Type = "F"							&&- Functions
		This.FunctionCase = Iif(Found(), __xfoxcode.Case, "")
		If Empty(This.FunctionCase)
			This.FunctionCase = lcDefaultCase
		Endif


		*- Itens para o codesnippet definidos no foxcode.dbf
		If This.chkCodeSnippet = "1"
			Select Abbrev, Expanded From __xfoxcode Where Type = "U" And Not Deleted() Into Array This.ItemsCodeSnippets
		Else
			Dimension This.ItemsCodeSnippets[1,2]
			This.ItemsCodeSnippets[1,1] = ""
		Endif


		Use In __xfoxcode

		If Used(lcAlias)
			Select (lcAlias)
		Endif


		*- Incremental IntelliSense.
		*- pesquisa pelo IntelliSense a cada tecla pressionada.
		Bindevent(0, WM_KEYUP, This, "GetKeyPressed", 4)
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ used on bindevent and to control the main method
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetKeyPressed
		Lparameters pln1 As Integer, pln2 As Integer, plnKey As Integer, pln4 As Integer, pll5 As Boolean

		Set Console Off
		Sys(2030,0)

		*- check if debug is active
		*- desabilito o IntelliSense se o debug estiver ativo para nao entrar em conflito com o foxcodeplus.
		This.ChkDebugger()
		If This.HasDebugger
			If This.IntelliSense.Showed
				This.IntelliSense.Hide()
			Endif
			Sys(2030,1)
			Return
		Endif

		*- Save VFP environment
		This.SetFoxcodePlusEnvironment(1)

		*- main function
		This.Main(pln1, pln2, plnKey, pln4, pll5)

		*- Restore VFP environment
		This.SetFoxcodePlusEnvironment(0)

		*this.EditorToolTip.NoClose = .f.
		Set Console On
		Activate Screen								&&- caso o Error list esteja aberto asseguro de que qualquer informacao "output" seja enviada ao _Scree
		Sys(2030,1)
		Return
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Verifico o que estou digitando para compor o montar o conteudo para o IntelliSense
	*/ **- MAIN FUNCTION -**
	*/------------------------------------------------------------------------------------------------
	Protected Procedure Main
		Lparameters pln1 As Integer, pln2 As Integer, plnKey As Integer, pln4 As Integer, pll5 As Boolean

		Set Console Off

		*- sempre que teclar algo, asseguro de fechar o tooltip do editor caso esteja aberto
		If Lastkey() <> 46
			This.EditorToolTip.Hide()
		Endif

		*- check for a valid editor and foxtools.fll
		If Not This.SetWontop()
			If This.IntelliSense.Showed
				This.IntelliSense.Hide()
			Endif
			Return
		Endif

		*- controlo a integridade da tela "Error List".
		*- isso pq na command window eu posso enviar um "clear" e limpar a tela.
		*- com o codigo abaixo mais alguns codigos no lostfocus da tela e no activate resolvo o problema.
		If Type("this.FormErrorList.LockScreen") = "L"
			If This.FormErrorList.LockScreen = .T.
				This.FormErrorList.LockScreen = .F.
				This.FormErrorList.Refresh()
			Endif
		Endif

		*- save lastkey pressed (used to increment)
		This.Lastkey = Lastkey()


		*- invalid combination key
		If 	(This.Lastkey = 50 And plnKey = 40)	Or  ;	&&- SHIFT + ARROW DOWN
			(This.Lastkey = 50 And plnKey = 16)	Or	;	&&- SHIFT + ARROW DOWN
			(This.Lastkey = 56 And plnKey = 38)	Or	;	&&- SHIFT + ARROW UP
			(This.Lastkey = 56 And plnKey = 16)	Or	;	&&- SHIFT + ARROW UP
			(This.Lastkey = 54 And plnKey = 39)	Or	;	&&- SHIFT + ARROW RIGHT
			(This.Lastkey = 54 And plnKey = 16)	Or  ;	&&- SHIFT + ARROW RIGHT
			(This.Lastkey = 52 And plnKey = 37)	Or	;	&&- SHIFT + ARROW LEFT
			(This.Lastkey = 52 And plnKey = 16)	Or	;	&&- SHIFT + ARROW LEFT
			(This.Lastkey = 57 And plnKey = 33)	Or	;	&&- SHIFT + ARROW PGUP
			(This.Lastkey = 57 And plnKey = 16)	Or  ;	&&- SHIFT + ARROW PGUP
			(This.Lastkey = 51 And plnKey = 34)	Or	;	&&- SHIFT + ARROW PGDN
			(This.Lastkey = 51 And plnKey = 16)	Or  ;	&&- SHIFT + ARROW PGDN
			(This.Lastkey = 49 And plnKey = 35)	Or	;	&&- SHIFT + END
			(This.Lastkey = 49 And plnKey = 16)	Or	;	&&- SHIFT + END
			(This.Lastkey = 55 And plnKey = 36)	Or	;	&&- SHIFT + HOME
			(This.Lastkey = 55 And plnKey = 16)			&&- SHIFT + HOME

			If This.IntelliSense.Showed
				This.IntelliSense.Hide()
			Endif
			Return
		Endif


		*- caso pressionei uma das teclas válidas para acionar o IntelliSense
		*- De a "A" a "Z"                 de "a" a "z" ....               de "0" a "9" ....             "." or "*" or "#" or "_" or "Backspace"
		If Between(This.Lastkey,65,90) Or Between(This.Lastkey,97,122) Or Between(This.Lastkey,48,57) Or Inlist(This.Lastkey,46,42,35,95,127)

			This.IntelliSense.ManualChoice = .F.

			*- quando um script do foxcode.dbf é acionado o IntelliSense do core do VFP é acionado
			*- neste caso se eu digitar algo fecho o IntelliSense do core do VFP para abrir o IntelliSense do Foxcodeplus.
			If This.FoxcodeCore
				This.FoxcodeCore = .F.
				Return
			Endif

			*- pego a texto da linha corrente que estou digitando até a posicao do cursor
			Local lcText, lnWordCount, lcLastFullWord, lcLastWord, lcCommand1, lcCommand2, lcCommand3, lnLines
			lcText = This.TreatLine(This.GetTextLine())
			lcText = Iif(Substr(lcText,1,1) = "#", lcText, This.TreatWords(lcText))
			lnWordCount = Getwordcount(lcText)
			lcLastFullWord = Getwordnum(lcText, lnWordCount)
			lcLastWord = Iif("."$lcLastFullWord, Substr(lcLastFullWord, Rat(".",lcLastFullWord,1)+1), lcLastFullWord)
			lcCommand1 = Getwordnum(lcText,1)
			lcCommand2 = Getwordnum(lcText,2)
			lcCommand3 = Getwordnum(lcText,3)

			This.CursorPos = _EdGetPos(This.EditorHwnd)
			This.CursorLine = This.GetLineNo()
			This.HasDot = Iif("."$lcLastFullWord,.T.,.F.)
			This.LastWord = lcLastWord


			*- se estou dentro de uma string ou dentro de um Text...EndText nao abro o IntelliSense
			If Lastkey() <> 46
				If This.IsInQuotes(lcText)
					Return
				Else
					This.IsTextEndText = This.GetTextEndText(lcText)
				Endif
			Endif

			*- especifics behaviors for some keys
			If Not This.IsTextEndText
				Do Case
						*- summary like Visual Studio
						*- If pressed three times "***"
					Case This.Lastkey = 42
						This.SetSummary()
						Return

					Case This.Lastkey = 46
						Return

						*- controlo o backspace
					Case This.Lastkey = 127
						*wait window lcText+chr(13)+this.TextLine nowait
						*- apaguei o "."
						Local llHasDelDot, llHasDelEqual
						llHasDelDot = .F.
						If Right(This.TextLine,1) = "." And Right(lcText,1) <> "." &&and "."$this.TextLine
							llHasDelDot = .T.
						Endif

						llHasDelEqual = .F.
						If Right(This.TextLine,1) = "=" And Right(lcText,1) <> "=" &&and "="$this.TextLine
							llHasDelEqual = .T.
						Endif

						*- escondo o IntelliSense se nao for possivel recompor o texto ao apagar com o backspace
						If Empty(lcText) Or Empty(This.TextLine) Or llHasDelDot Or llHasDelEqual Or (lcText==This.TextLine)
							If This.IntelliSense.Showed
								This.IntelliSense.Hide()
								This.TextLine2 = This.TextLine
								This.TextLine = lcText
								Return
							Endif
						Endif

						*- include file doesn't exist
					Case Lower(Chr(This.Lastkey))="h" And lnWordCount = 3 And Lower(Substr(lcText,1,10)) == " # include" And Lower(Right(lcText,2)) == ".h"
						If Not This.GetFilePath(@lcCommand3)
							This.ShowErrorWriteTime(1994, Upper(Justfname(lcCommand3)))
						Endif
						Return

						*- do "program" doesn't exist
					Case Inlist(Lower(Chr(This.Lastkey)),"g","r","p","e") And lnWordCount = 2 And Lower(lcCommand1) == "do" And Inlist(Lower(Right(lcCommand2,4)),".prg",".mpr",".spr",".qpr",".fxp",".app",".exe")
						If Not This.GetFilePath(@lcCommand2)
							This.ShowErrorWriteTime(1,Upper(Justfname(lcCommand2)))
						Endif
						Return
				Endcase
			Endif


			*- Nao prossigo se não mudei o que digitei ou se
			*- o comando nao tem IntelliSense plus.
			If Not This.IsTextEndText And ;
					( ;
					(This.TextLine == lcText) Or Empty(lcText) Or ;
					("<"+Lower(lcCommand1)+">" $ This.NoIntelliSense And lnWordCount>=2) Or ;
					("<"+Lower(lcCommand1+" "+lcCommand2)+">" $ This.NoIntelliSense And lnWordCount>=3) Or ;
					("<"+Lower(lcCommand1+" "+lcCommand2+" "+lcCommand3)+">" $ This.NoIntelliSense And lnWordCount>=4) Or ;
					(Lower(Getwordnum(lcText,lnWordCount-1)) == "in" And Lower(lcCommand1+" "+lcCommand2) <> "for each") Or ;
					(Lower(Getwordnum(lcText,lnWordCount-1)) == "set(" ) Or ;
					(Substr(Lower(lcCommand1),1,4) == "sele" And lnWordCount = 2 And Right(lcCommand2,1)<>".") Or ;
					(Lower(Getwordnum(lcText,lnWordCount-1)) == "as" And lnWordCount >= 2) Or ;
					(Lower(lcCommand1) == "index" And Lower(lcCommand2) == "on" And lnWordCount >= 4) Or ;
					(Lower(Substr(lcCommand1,1,4)) == "crea" And lnWordCount <= 4) ;
					)

				If This.TextLine <> lcText
					This.IntelliSense.Find(lcLastWord)
				Endif

				Return
			Else
				*- estou dentro de um text..endtext, é uma instrucao SQL e pressionei space ao lado das clausulas abaixo
				*- neste caso nao abro o intellisense incremental pois ira abrir o intellisense do foxcode.app
				If This.IsTextEndText And This.IsSqlIntelliSense And ;
						( ;
						inlist(Lower(Getwordnum(lcText,lnWordCount-1)), "from", "join", "into", "update") Or ;
						inlist(Lower(Getwordnum(lcText,lnWordCount-2)), "from", "join", "into", "update") ;
						)

					If This.TextLine <> lcText
						This.IntelliSense.Find(lcLastWord)
					Endif

					If This.IntelliSense.Showed
						This.IntelliSense.Hide()
					Endif

					Return

					*- prossigo com a checagem para abertura do intellisense incremental
				Else
					*- Preencho as propriedades auxiliares para preenchimento do IntelliSense
					This.TextLine2 = This.TextLine
					This.TextLine = lcText
					This.WordCount = lnWordCount
				Endif
			Endif


			*- sempre escondo o IntelliSense antes de reabri-lo.
			*- faço isso para limpar a lista e executar outros comandos que estão dentro no method hide.
			If Not This.HasDot
				This.IntelliSense.Hide()

				*- 1 to 9 ... prevendo erros de sintax quando palavras iniciadas por numero
				*- assim ignoro IntelliSense para isso.
				If Isdigit(This.LastWord)   &&between(asc(substr(this.LastWord,1,1)), 48, 57)
					Return
				Endif

				*- se o caracter atual for " " espaço escondo o IntelliSense ao pressionar backspace
				If This.Lastkey=127
					If _EdGetChar(This.EditorHwnd, This.CursorPos-1) = " "
						If This.IntelliSense.Showed
							This.IntelliSense.Hide()
						Endif
						Return
					Endif
				Endif

				*- busco os itens para a lista do IntelliSense
				This.IntelliSense = Newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")
				lnLines = 0
				If Not This.IsTextEndText
					*--- vfp intellisense ---*
					lnLines = lnLines + This.GetFCs(This.LastWord)						&&- funcoes e comandos
					lnLines = lnLines + This.GetCodeSnippets(This.LastWord)				&&- CodeSnippet
					lnLines = lnLines + This.GetTablesUsed(This.LastWord)			 	&&- tabelas abertas em run-time
					lnLines = lnLines + This.GetTablesDataEnvironment(This.LastWord)	&&- tabelas abertas em run-time
					lnLines = lnLines + This.GetAPIs(This.LastWord)						&&- APIs em run-time
					lnLines = lnLines + This.GetControls(This.LastWord)					&&- Objetos contidos em forms, classes e toolbar em write-time.
					lnLines = lnLines + This.GetObjectsRunTime(This.LastWord)			&&- Objetos em memória em run-time
					lnLines = lnLines + This.GetSetProcInfoPrgRunTime()					&&- Verifica os PRGs invocados pelo SET PROCEDURE TO em run-time.
					lnLines = lnLines + This.GetProcInfo(0,1,.T.)						&&- funcoes, methodes, events, variables, cursors, tables, DLLs function and #defines em write-time
				Else
					*--- sql intellisense ---*
					If This.chkTFsql = "1" And This.IsSqlIntelliSense
						*- tabelas do SQL no modo incremental são verificadas somente se for uma instrucao "SELECT" ou qualquer outra com "WHERE"
						*-
						If Getwordnum(Lower(This.TextEndBlock),1) == "select" Or " where " $ Lower(This.TextEndBlock)
							Local Array laCnx[1]
							If This.chkIncrTablesSql = "1"
								lnLines = lnLines + This.GetSqlTables(This.LastWord, .T., .T.)				&&- tabelas no SQL
								lnLines = lnLines + This.GetSqlTablesInCmd(This.LastWord, 2, .T., .F.)		&&- alias no instruncao SQL
							Else
								*- tabelas e alias existentes na select-sql
								lnLines = lnLines + This.GetSqlTablesInCmd(This.LastWord, 0, .T., .F.)		&&- tabelas e alias no instruncao SQL
							Endif
						Endif

						*- Todos os campos das tabelas e alias incluidas no instruncao SQL no modo incremental
						If This.chkIncrFieldsSql = "1"
							lnLines = lnLines + This.GetSqlFieldsInAllTablesCmd()
						Endif
					Endif
				Endif
			Else
				*- a funcao this.GetDot() ja preencheu a lista do IntelliSense pois a mesma é acionada pelo commando On key label "."
				*- com isso, por padrao posiciona no 1o. item da lista.
				If This.IntelliSense.Rows >= 1
					lnLines = 1
				Else
					lnLines = 0
				Endif
			Endif


			*- Apresento ou escondo o IntelliSense
			If lnLines > 0
				With This.IntelliSense
					*- se o IntelliSense ja esta com a tela aberta fecho-o... porem o parametro ".T."
					*- significa que o seu conteudo permarecerá o mesmo.

					If .Showed
						.Hide(.T.)
					Endif

					*- posiciono na lista conforme o que foi digitado
					.LastFind = ""
					.Find(lcLastWord)

					*- se o tooltip padrao da IDE do VFP estiver aberto a funcao abaixo fecha-o
					*- tambem asseguro que vou apresentar o IntelliSense no Handle correto.
					_wSelect(This.EditorHwnd)

					*- apresento o IntelliSense
					.Show()
				Endwith

			Else
				This.IntelliSense.Hide()
			Endif

			Return


			*- nenhuma das teclas validas para chamada da lista de itens,
			*- então trato outras funcionalidades
		Else

			*- inicio tratamentos de navegação do IntelliSense sem o foco no mesmo
			Do Case
					*- se abrir aspas, aspas simples ou colchetes, é fechado automaticamente.
				Case Inlist(This.Lastkey,34,39,91) And This.chkAutoCloseQuotes = "1"
					Do Case
						Case This.Lastkey=34		&&- "
							Keyboard '"'

						Case This.Lastkey=39		&&- '
							Keyboard "'"

						Case This.Lastkey=91		&&- [
							Keyboard ']'
					Endcase
					Keyboard '{leftarrow}'


					*- up arrow, down arrow, page up, page down, ctrl+home, ctrl+end or enter
				Case Inlist(This.Lastkey,5,24,3,18,29,23,13)
					*- se troquei de linha faço a compilação em write-time para obter os erros
					If Not This.IntelliSense.Showed
						If This.GetLineNo() <> This.CursorLine
							This.CursorLine = This.GetLineNo()
							This.GetErrorList()
						Endif
					Endif


					*- se pressionei "(" or "," or LeftArrow or RightArrow
				Case Inlist(This.Lastkey,40,44,19,4) And Not Inlist(plnKey,17,83) 		&&-plnKey 17 and 83 is CTRL+S
					If This.IntelliSense.Showed
						This.IntelliSense.Hide()
					Endif

					*- verifico se tem assinatura de metodo ou funcao para ser apresentanda no tooltip
					If Not This.IsTextEndText
						This.GetSignature()
					Endif


					*- limpo pq fechei o IntelliSense e posso recomeçar a pesquisa incremental
				Case This.Lastkey = 27
					This.HasSelectedItem = .F.
					If This.IntelliSense.Showed
						This.TextLine = ""
					Endif


					*- ATTENTION: SPACE IS USED ESCLUSIVELY FOR DEFAULT VFP CORE IntelliSense.
					*- NEVER USE THIS GAP TO INCLUDE SOMETHING. FOR THAT WE HAVE TO CUSTOM
					*- FOXCODE.APP PROJECT. FOXCODEPLUS.APP WORKS TOGETHER FOXCODE.APP
				Case This.Lastkey = 32
					* dont use this place.


					*- escondo o IntelliSense
				Otherwise
					If This.Lastkey <> 46
						If This.IntelliSense.Showed
							This.IntelliSense.Hide()
						Endif
					Endif
			Endcase

		Endif

		Return
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ Verifico se a tela ativa da IDE do VFP é uma tela valida para o IntelliSense
	*/ e se for capturo o handle do tela
	*/------------------------------------------------------------------------------------------------
	Procedure SetWontop
		Lparameters plcWindowsNumbers
		Local lnOldEditorHwnd

		Set Console Off

		If Not "foxtools.fll" $ Lower(Set("Library"))
			Set Library To Home(1)+"foxtools.fll" Additive
		Endif

		*- handle do editor do VFP ativo
		lnOldEditorHwnd = This.EditorHwnd
		This.EditorHwnd = _wontop()

		*- se troquei de editor e o IntelliSense esta aberto deve fecha-lo
		If Not Isnull(This.IntelliSense) Then
			If This.EditorHwnd <> lnOldEditorHwnd
				If This.IntelliSense.Showed
					This.IntelliSense.Hide()
				Endif
			Endif
		Endif

		*- Não prossigo se nao foi possivel obter o Hwnd do editor corrente ou se existir uma tela filha do editor aberta com foco
		*- ex: tela "Find" or "Go to line" and so on.
		If This.EditorHwnd <= 0 Or (Sys(2325,This.EditorHwnd) = This.EditorHwnd) Or This.GetLineNo()=-1
			Return .F.
		Endif


		*- -1 -> the window is not an edit window
		*-  0 -> command window
		*-  1 -> modify command window
		*-  2 -> modify file window
		*-  3 -> memo window
		*-  6 -> Query
		*-  7 -> screen
		*-  8 -> menu designer code window
		*-  9 -> view
		*- 10 -> method edit window in class or form designer
		*- 11 -> Text
		*- 12 -> modify procedure window
		*- 13 -> Project Text

		*- se não for uma tela de IDE valida para abertura do IntelliSense não prossigo
		Local Array laEditorSets[25]
		Try
			_EdGetEnv(This.EditorHwnd, @laEditorSets)
		Catch
			laEditorSets[25] = -1
			This.EditorFileName = ""
		Endtry

		This.EditorFileName = Alltrim(laEditorSets[1])
		This.EditorSource = laEditorSets[25]
		This.EditorFontName = laEditorSets[22]
		This.EditorFontSize = laEditorSets[23]

		*- configuro o editor com a font do visual studio
		If This.chkFont = "1"
			If Lower(laEditorSets[22]) <> "consolas"
				laEditorSets[22] = "Consolas"
				_EdSetEnv(This.EditorHwnd, @laEditorSets[25])
				_wSelect(This.EditorHwnd)
			Endif
		Else
			If Lower(laEditorSets[22]) = "consolas"
				laEditorSets[22] = "courier new"
				_EdSetEnv(This.EditorHwnd, @laEditorSets[25])
				_wSelect(This.EditorHwnd)
			Endif
		Endif

		*- se for um editor válido
		If Empty(plcWindowsNumbers)
			If Inlist(This.EditorSource, 1, 8, 10, 12)
				Return .T.
			Else
				Return .F.
			Endi
		Else
			If Transform(This.EditorSource,"@L 99") $ plcWindowsNumbers
				Return .T.
			Else
				Return .F.
			Endif
		Endif
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ se a janela do debug estiver ativada pauso o foxcodeplus ate que o debug seja fechado.
	*/ faço isso para evitar conflitos durante a depuraçao de programas.
	*/------------------------------------------------------------------------------------------------
	Protected Procedure ChkDebugger
		If Wvisible("visual foxpro debugger") And Not This.HasDebugger
			Local lnDebugHwnd
			*- capturo o hwnd da tela do debug
			lnDebugHwnd = This.GetDebuggerHwnd()

			*- pauso o foxcodeplus
			This.HasDebugger = .T.
			*unbindevents(0, WM_KEYUP)

			*- quando fechar o debug o foxcodeplus sera restartado (despausado)
			Bindevent( lnDebugHwnd, WM_DESTROY, This, "ReStartIntelliSense", 4 )
			Return .F.
		Endif
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ usado exclusivamente no bindevent() da this.ChkDebugger() para reativar o foxcodeplus.
	*/ quando fechar o Debugger o FoxcodePlus volta a funcionar
	*/------------------------------------------------------------------------------------------------
	Protected Procedure ReStartIntelliSense(HWnd As Integer, msg As Integer, wparam As Integer, Lparam As Integer)
		Unbindevents(HWnd, msg)

		Sys(2030,0)
		This.HasDebugger = .F.
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ usado exclusivamente no bindevent() da this.ChkDebugger() para obter o Hwnd da tela do Debugger do VFP
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetDebuggerHwnd
		Set Console Off

		Declare Integer GetActiveWindow In win32api As xxfcpWinAPI_GetActiveWindow
		Declare Integer GetWindow In Win32API As xxfcpWinAPI_GetWindow Integer HWnd, Integer nType
		Declare Integer GetWindowText In Win32API As xxfcpWinAPI_GetWindowText Integer HWnd, String @cText, Integer nType

		Local lnNext, lcText
		lnNext = xxfcpWinAPI_GetActiveWindow()

		*- iterate through the open windows
		Do While lnNext<>0
			*- get window title
			lcText = Replicate(Chr(0),80)
			xxfcpWinAPI_GetWindowText(lnNext,@lcText,80)
			If "visual foxpro debugger" $ Lower(lcText)
				Return lnNext
			Endif
			lnNext = xxfcpWinAPI_GetWindow(lnNext,2)
		Enddo

		Clear Dlls "xxfcpWinAPI_GetActiveWindow","xxfcpWinAPI_GetWindow","xxfcpWinAPI_GetWindowText"

	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Retorna o numero da linha a qual estou posicionado no codigo
	*/ caso ocorra erro retorna -1
	*/------------------------------------------------------------------------------------------------
	Procedure GetLineNo
		Set Console Off

		Local lnCursorPos, lnLineNo

		Try
			lnCursorPos = _EdGetPos(This.EditorHwnd)
			If lnCursorPos > 0
				lnLineNo = _EdGetLNum(This.EditorHwnd, lnCursorPos) + 1
			Else
				lnLineNo = 0
			Endif
		Catch
			lnLineNo = -1
		Endtry

		Return lnLineNo
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Captura o texto do editor do VFP conforme as posição linha inicial e linha final
	*/------------------------------------------------------------------------------------------------
	Procedure GetText
		Lparameters plnStart, plnEnd

		Set Console Off

		Local lnStartPos, lnEndPos, lcString

		Try
			lnStartPos = _EdGetLPos(This.EditorHwnd, plnStart)
			lnEndPos = _EdGetLPos(This.EditorHwnd, plnEnd + 1 ) - 1
			lcString = _EdGetStr(This.EditorHwnd, lnStartPos, lnEndPos)
		Catch
			lcString = ""
		Endtry

		Return lcString
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Retorna o texto conforme numero da linha especificada, ou o texto da linha atual.
	*/ O retorno pode ser a linha inteira ou a linha até onde o cursor esta posicionado.
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetTextLine
		Lparameters plnLineNo, pllFullLine
		Local lnStartPos, lnLineNo, lnEndPos, lcString, lnChkLineNo

		Set Console Off

		plnLineNo = Evl(plnLineNo,0)

		*- linha onde esta posicionado o cursor
		If plnLineNo = 0
			lnLineNo = This.GetLineNo()
			If lnLineNo = -1
				lnLineNo = This.GetLineNo()
				If lnLineNo = -1
					Return ""
				Endif
			Endif
			lnLineNo = lnLineNo - 1
		Else
			lnLineNo = plnLineNo - 1
		Endif

		*- Verifico na linha anterior se existe quebra de linha ";"
		*- caso exista procuro a linha onde foi iniciado o comando.
		Do While .T.
			lnChkLineNo = lnLineNo - 1
			If lnChkLineNo > 0
				lcString = Strtran(Strtran(_EdGetStr(This.EditorHwnd, _EdGetLPos(This.EditorHwnd, lnChkLineNo), _EdGetLPos(This.EditorHwnd, lnChkLineNo+1)-1), Chr(13), ""), Chr(10), "")
				If Right(Alltrim(lcString),1)=";"
					lnLineNo = lnChkLineNo
					Loop
				Else
					Exit
				Endif
			Else
				Exit
			Endif
		Enddo

		*- inicio da linha
		lnStartPos = _EdGetLPos(This.EditorHwnd, lnLineNo)

		*- fim da linha
		If pllFullLine
			*- texto da linha inteira
			lnEndPos = _EdGetLPos(This.EditorHwnd, lnLineNo+1)-1
		Else
			*- texto da linha até onde o cursor esta posicionado
			lnEndPos = _EdGetPos(This.EditorHwnd)
		Endif

		*- retorno a texto entre as posicoes indicadas
		If lnStartPos == lnEndPos
			lcString = ""
		Else
			lnEndPos = lnEndPos - 1
			lcString = Strtran(_EdGetStr(This.EditorHwnd, lnStartPos, lnEndPos), ";", "")
		Endif

		Return lcString
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ substituo a palavra que estou digitando pelo item que selecionei
	*/------------------------------------------------------------------------------------------------
	Procedure SelectItem
		Lparameters plnKeyAscii
		Local lcNewWord, lnStartPos, lnEndPos, lnNewPos, lnHandle, lcThisWhat, lcWith, lcTextLine, lcText, lcLastFullWord, lcLastWord

		Set Console Off

		This.HasSelectedItem = .T.
		lcText = This.TreatLine(This.GetTextLine())		&&- capturo o texto ate a posicao onde  estou posicionado
		lcText = Iif(Substr(lcText,1,1) = "#", lcText, This.TreatWords(lcText))

		*- dentro de IntelliSense com conteudo do SQL só fecho e não seleciono se pressionei space
		*- se eu navegar nas opcoes com as setas UP e DOWN e pressionar space, neste caso a opcao pode ser selecionada com space.
		If This.IsTextEndText And This.IsSqlIntelliSense And plnKeyAscii = 32 And Not This.IntelliSense.ManualChoice
			This.IntelliSense.Hide()
			This.IsSqlIntelliSense = .F.
			Keyboard Chr(32)
			Return
		Endif

		*- item selecionado no IntelliSense
		lcNewWord = This.IntelliSense.ActiveItem

		Do Case
				*- se selecionei uma classe coloco entre aspas caso esteja posicionado nas funcoes abaixo
			Case This.IntelliSense.ActiveImage = 1 And ( "createobject(" $ Lower(lcText) Or "createobjectex(" $ Lower(lcText) Or "newobject(" $ Lower(lcText) )
				lcNewWord = ["]+lcNewWord+["]

				*- large table name
			Case This.IntelliSense.ActiveImage = 16 And Getwordcount(lcNewWord) >= 2
				lcNewWord = "["+lcNewWord+"]"

				*- field types
			Case This.IntelliSense.ActiveImage = 12 And "," $ lcNewWord
				lcNewWord = Substr(lcNewWord,1,1)
		Endcase


		*- aqui estou prevendo a digitação rapida, ou seja, para casos que a digitação
		*- é mais rapida que o posicionamento do item no IntelliSense, assim, no momento da escolha
		*- verifico se o conteudo do editor corresponde ao selecionado no IntelliSense
		lcLastFullWord = Strtran(Strtran(Strtran(Strtran(Strtran(Alltrim(lcText),"  "," "), "[ ","["), " ]", "]"), "( ","("), " )", ")")
		lcLastFullWord = Getwordnum(lcLastFullWord, Getwordcount(lcLastFullWord))
		If Substr(lcText,1,1) <> "#"
			lcLastWord = This.TreatWords(lcLastFullWord)
			lcLastWord = Iif("."$lcLastWord, Substr(lcLastWord, Rat(".",lcLastWord,1)+1), lcLastWord)
		Else
			lcLastWord = lcText
		Endif
		This.LastWord = Getwordnum(lcLastWord,Getwordcount(lcLastWord))
		If Alltrim(Lower(This.LastWord)) == Alltrim(Lower(lcNewWord))		&&- nao prossigo se o item digitado é exatamente igual ao selecionado.
			This.IntelliSense.Hide()
			If plnKeyAscii <> 9
				Keyboard Chr(plnKeyAscii) Plain
			Endif
			Return
		Endif


		*- Forms controls and visual classes
		If Not Empty(lcNewWord) And This.IntelliSense.Found

			If This.IntelliSense.ActiveImage = 13 And This.EditorSource = 10

				lcWith = Lower(This.GetWithHierarchy())
				lcTextLine = This.TreatLine(This.GetTextLine())
				lcThisWhat = ""

				*- se a palavra conter "this." ou "thisform." ignoro.
				*- e se estou fora de um bloco with..endwith tambem ignoro... assim,
				*- so entro no "if" abaixo se estou dentro de um with...endwith sem incluir o "this." ou "thisform." na linha

				&&-not inlist(lower(substr(lcLastFullWord, 1, at(".",lcLastFullWord))), "this.","thisform.","thisformset.")
				If Empty(Lower(Substr(lcLastFullWord, 1, At(".",lcLastFullWord)))) And ;
						(Empty(lcWith) Or (Not Empty(lcWith) And Substr(lcTextLine,1,1) <> "." And Substr(Getwordnum(lcTextLine,2),1,1) <> "."))

					Local lcCurrentControl
					Local Array laControl[1]
					laControl[1] = ""

					lcCurrentControl = _wtitle(This.EditorHwnd)
					lcCurrentControl = Substr(lcCurrentControl, 1, At(".",lcCurrentControl)-1)
					Aselobj(laControl, 1 )

					If Vartype(laControl[1]) = "O"
						Do Case
								*- objetos na raiz do form
							Case laControl[1].BaseClass = "Form"
								If Alltrim(Lower(lcCurrentControl)) <> Alltrim(Lower(lcNewWord))
									lcThisWhat = "Thisform."
								Else
									lcThisWhat = ""
									lcNewWord = "This"
								Endif

								*- objetos em objetos do tipo container
							Case Alltrim(Lower(lcCurrentControl)) == Alltrim(Lower(laControl[1].Name))
								lcThisWhat = "This."

								*- se selecionei o mesmo objeto ao qual estou posicionado, neste caso,
								*- troco o nome do objeto pelo "This", assim faço a referencia correta.
							Case Alltrim(Lower(lcCurrentControl)) == Alltrim(Lower(lcNewWord))
								lcThisWhat = ""
								lcNewWord = "This"

								*- selecionei um objeto o qual devo informar o caminho completo do objeto
							Otherwise
								lcThisWhat = "This.Parent."
						Endcase

					Endif
				Endif

				lcNewWord = lcThisWhat + lcNewWord
			Endif


			*- upper, lower or proper vfp commands
			If Inlist(This.IntelliSense.ActiveImage, 2, 18)
				Do Case
					Case Upper(This.CommandCase)="U" Or Substr(lcNewWord,1,1) = "#"
						lcNewWord = Upper(lcNewWord)
					Case Upper(This.CommandCase)="L"
						lcNewWord = Lower(lcNewWord)
					Case Upper(This.CommandCase)="P"
						lcNewWord = Proper(lcNewWord)
					Case Upper(This.CommandCase)="M"
						lcNewWord = lcNewWord
					Case Upper(This.CommandCase)="X"	&&- No Expand keyword
						lcAlias = Alias()

						Use (_Foxcode) Again Alias __xfoxcode In 0 Shared

						Locate For __xfoxcode.Type = "C" And __xfoxcode.Expanded = Padr(lcNewWord,Len(__xfoxcode.Expanded))
						If Found()
							lcNewWord = Subs(lcNewWord, 1, Len(Alltrim(__xfoxcode.Abbrev)))
						Endif

						Use In __xfoxcode
						If Used(lcAlias)
							Select(lcAlias)
						Endif

					Otherwise
						Return ""
				Endcase
			Endif


			*- upper, lower or proper vfp functions
			If This.IntelliSense.ActiveImage = 19
				Do Case
					Case Upper(This.FunctionCase)="U"
						lcNewWord = Upper(lcNewWord)
					Case Upper(This.FunctionCase)="L"
						lcNewWord = Lower(lcNewWord)
					Case Upper(This.FunctionCase)="P"
						lcNewWord = Proper(lcNewWord)
					Case Upper(This.FunctionCase)="M"
						lcNewWord = lcNewWord
					Case Upper(This.FunctionCase)="X"		&&- No Expand keyword
						lcAlias = Alias()

						Use (_Foxcode) Again Alias __xfoxcode In 0 Shared
						Select __xfoxcode

						Locate For __xfoxcode.Type = "F" And __xfoxcode.Expanded = Padr(lcNewWord,Len(__xfoxcode.Expanded))
						If Found()
							lcNewWord = Subs(lcNewWord, 1, Len(Alltrim(__xfoxcode.Abbrev)))
						Endif

						Use In __xfoxcode
						If Used(lcAlias)
							Select(lcAlias)
						Endif

					Otherwise
						Return ""
				Endcase
			Endif


			lnStartPos = _EdGetPos(This.EditorHwnd) - Len(This.LastWord)
			lnEndPos = This.CursorPos

			If lnStartPos > lnEndPos
				lnEndPos = lnStartPos
			Endif

			lnNewPos = lnStartPos + Len(lcNewWord)


			*- inserindo o texto
			This.IntelliSense.Hide()
			_EdSelect(This.EditorHwnd, lnStartPos, lnEndPos)
			_EdDelete(This.EditorHwnd)
			_EdInsert(This.EditorHwnd, lcNewWord, Len(lcNewWord))
			_EdSetPos(This.EditorHwnd, lnNewPos)
			This.CursorPos = lnNewPos
		Endif

		Do Case
				*- nenhum item na lista do IntelliSense foi selecionado
				*- porem o IntelliSense foi aberto
			Case Empty(lcNewWord)
				Keyboard Chr(plnKeyAscii) Plain

				*- se pressionei "(" para escolher o item
			Case plnKeyAscii = 40
				If Right(lcNewWord,2) <> "()"
					Keyboard Chr(plnKeyAscii) Plain
				Endif

				*- se pressionei ")" e for uma variavel
			Case plnKeyAscii = 41
				If Right(lcNewWord,1) <> ")"
					Keyboard Chr(plnKeyAscii) Plain
				Endif

				*- se pressionei "ENTER" or " ",  #   *   +   ,   -   /   <   >
			Case Inlist(plnKeyAscii, 13, 32, 35, 42, 43, 44, 45, 47, 60, 62)
				Keyboard Chr(plnKeyAscii) Plain

				*- se pressionie "." or "="
			Case Inlist(plnKeyAscii, 46, 61)
				*- see this.GetDot() and this.GetEqual()

				*-	nao faço nada apos inserir o texto
			Otherwise

		Endcase

		This.IntelliSense.Hide()
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busca todos os cursores e tabelas abertos em run-time
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetTablesUsed
		Lparameters plcWord

		Set Console Off

		Local Array laTables[1]
		laTables[1] = ""
		Local lnLines, lnItemsFound, lnRows, lnItems, lnx, lcItem, lcToolTip

		If This.chkTF <> "1"		&& foxcodeplus.ini
			Return 0
		Endif

		lnLines = Aused(laTables)
		If lnLines = 0
			Return 0
		Endif

		*- conto quantas tabelas/cursores estao abertos conforme oque estou digitando
		If Not Empty(plcWord)
			lnItemsFound = 0
			For lnx = 1 To lnLines
				If This.ChkIncremental(plcWord,laTables[lnx,1])
					lnItemsFound = lnItemsFound + 1
				Endif
			Endfor
		Else
			lnItemsFound = lnLines
		Endif

		*- insiro na lista
		If lnItemsFound > 0
			For lnx = 1 To lnLines
				lcItem = Proper(laTables[lnx,1])
				If This.ChkIncremental(plcWord,lcItem)
					lcToolTip = "Table "+Proper(Juststem(Dbf(lcItem)))+" alias "+lcItem + Chr(13) + "Table opened outside "+_wtitle(This.EditorHwnd)
					This.AddItem(lcItem, 16, lcToolTip)
				Endif
			Endfor
		Endif

		Return lnItemsFound
	Endpro



	*/------------------------------------------------------------------------------------------------
	*/ Busca todas as tabelas abertas no DataEnvironment (forms and reports)
	*/------------------------------------------------------------------------------------------------
	Procedure GetTablesDataEnvironment
		Lparameters plcWord, plnAdd
		Local lnx, lnItemsFound, loObj
		Local Array laControl[1]

		Set Console Off

		If This.chkTF <> "1"				&&- foxcodeplus.ini
			Return 0
		Endif

		If This.EditorSource <> 10
			Return 0
		Endif

		Declare This.ItemsTables[1,2]		&&- limpo o array.
		This.ItemsTables[1,1] = ""

		lnx = 0
		lnItemsFound = 0
		loObj = .Null.
		plnAdd = Iif(Parameters()=1, .T., plnAdd)


		*- REPORT designer aberto, entao procuro pelo DataEnvironment do FRX.
		If Pemstatus(_Screen,"reportbuilderdata",5)
			laControl[1] = ""
			For Each loObj In _vfp.Objects
				If Pemstatus(loObj,"baseclass",5) And Lower(loObj.BaseClass) = "dataenvironment"
					laControl[1] = loObj
				Endif
			Endfor

			*- FORM designer aberto, entao procuro pelo DataEnviroment do SCX.
		Else
			Aselobj(laControl, 2 )
		Endif


		*- checagem das tabelas no DataEnvironment (caso exista tabelas)
		If Vartype(laControl[1]) = "O"
			If Lower(laControl[1].BaseClass) = "dataenvironment"

				*- quantidade de tabelas abertas no dataenviromente conforme o que eu digitei
				For Each loObj In laControl[1].Objects
					lnx = lnx + 1
					If Pemstatus(loObj,"alias",5)
						If Iif(This.IncrementalResult, This.ChkIncremental(plcWord,loObj.Alias), .T.)
							lnItemsFound = lnItemsFound + 1
							Dimension This.ItemsTables[lnItemsFound,2]
							This.ItemsTables[lnItemsFound,1] = loObj.Alias

							*- free tables
							If loObj.Class = "Cursor" And Pemstatus(loObj,"cursorsource",5)
								This.ItemsTables[lnItemsFound,2] = "Alias " + This.ItemsTables[lnItemsFound,1] + " As Object " + loObj.Name + Chr(13) + ;
									"CursorSource " + loObj.CursorSource + Chr(13) + ;
									"Dataenvironment " + laControl[1].Name
								*- cursor adapter
							Else
								This.ItemsTables[lnItemsFound,2] = "Alias " + This.ItemsTables[lnItemsFound,1] + " As Object " + loObj.Name + Chr(13) + ;
									"CursorAdapter " + Chr(13) + ;
									"Dataenvironment " + laControl[1].Name
							Endif
						Endif
					Endif
				Endfor

				*- insiro na lista
				If lnItemsFound > 0 And plnAdd
					For lnx = 1 To lnItemsFound
						This.AddItem(This.ItemsTables[lnx,1], 16, This.ItemsTables[lnx,2])
					Endfor
				Endif

			Endif
		Endif

		Return lnItemsFound
	Endpro


	*/------------------------------------------------------------------------------------------------
	*/ Verifica se é um nome de tabela ou campo valido na tabela DBF ou SQL
	*/------------------------------------------------------------------------------------------------
	Protected Procedure ChkValidTableOrFieldName
		Lparameters plcTableOrField
		Local lnx, lnAsc

		If Empty(plcTableOrField) Or Not Isalpha(plcTableOrField)
			Return .F.
		Endif

		For lnx = 1 To Len(plcTableOrField)
			lnAsc = Asc(Substr(plcTableOrField,lnx,1))
			If !Between(lnAsc,65,90) And !Between(lnAsc,97,122) And !Between(lnAsc,48,57) And !lnAsc = 95
				Return .F.
			Endif
		Endfor

		Return .T.
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busca todos os campos de uma tabela/cursor
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetFields
		Lparameters plcTable

		Set Console Off

		Local Array laFields[1]
		laFields[1] = ""
		Local lnLines, lnRows, lnItems, lnx

		If This.chkTF <> "1"		&& foxcodeplus.ini
			Return 0
		Endif

		lnLines = Afields(laFields, plcTable)
		If lnLines = 0
			Return 0
		Endif

		*- insiro na lista
		If lnLines > 0
			For lnx = 1 To lnLines

				lcToolTip = "Field "+Proper(laFields[lnx,1])+ " type "

				Do Case
					Case laFields[lnx,2] = "C"
						lcToolTip = lcToolTip + "Character C("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "Y"
						lcToolTip = lcToolTip + "Currency Y("+Alltrim(Str(laFields[lnx,3]))+","+Alltrim(Str(laFields[lnx,4]))+")"

					Case laFields[lnx,2] = "D"
						lcToolTip = lcToolTip + "Date D("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "T"
						lcToolTip = lcToolTip + "DateTime T("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "B"
						lcToolTip = lcToolTip + "Double B("+Alltrim(Str(laFields[lnx,3]))+","+Alltrim(Str(laFields[lnx,4]))+")"

					Case laFields[lnx,2] = "F"
						lcToolTip = lcToolTip + "Float F("+Alltrim(Str(laFields[lnx,3]))+","+Alltrim(Str(laFields[lnx,4]))+")"

					Case laFields[lnx,2] = "G"
						lcToolTip = lcToolTip + "General G("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "I"
						lcToolTip = lcToolTip + "Integer I("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "L"
						lcToolTip = lcToolTip + "Logical or Boolean L("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "M"
						lcToolTip = lcToolTip + "Memo M("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "N"
						lcToolTip = lcToolTip + "Numeric N("+Alltrim(Str(laFields[lnx,3]))+","+Alltrim(Str(laFields[lnx,4]))+")"

					Case laFields[lnx,2] = "Q"
						lcToolTip = lcToolTip + "Varbinary Q("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "V"
						lcToolTip = lcToolTip + "Varchar or Varchar Binary V("+Alltrim(Str(laFields[lnx,3]))+")"

					Case laFields[lnx,2] = "W"
						lcToolTip = lcToolTip + "Blob W("+Alltrim(Str(laFields[lnx,3]))+")"

					Otherwise
						lcToolTip = lcToolTip + "(Invalid)"
				Endcase

				This.AddItem(Proper(laFields[lnx,1]), 17, lcToolTip)
			Endfor
		Endif

		Return lnLines
	Endpro


	*/------------------------------------------------------------------------------------------------
	*/ Busca as tabelas do database SGBD - SQL
	*/------------------------------------------------------------------------------------------------
	Procedure GetSqlTables
		Lparameters plcWord, pllAdd, pllClearArray
		Local lnItemsFound, lcAlias, lcToolTip, lcSqlTables, lnItemsCnt
		Local Array laCnx[1]

		Set Console Off

		If This.chkTFsql <> "1"
			Return 0
		Endif

		If pllClearArray
			Declare This.ItemsTables[1,2]
			This.ItemsTables[1,1] = ""
			This.ItemsTables[1,2] = ""
		Endif

		lnItemsFound = 0
		lnItemsCnt = Iif(Empty(This.ItemsTables[1,1]), 0, Alen(This.ItemsTables,1))
		lcSqlTables = "tmp"+Sys(2015)

		*- busco as tabelas no database
		If Asqlhandles(laCnx) > 0 And SQLTables(1,"TABLE",lcSqlTables) = 1
			lcAlias = Alias()

			Select (lcSqlTables)
			Scan
				If This.ChkIncremental(plcWord, table_name)
					lnItemsFound = lnItemsFound + 1


					lcToolTip = "Table " + Alltrim(table_cat)+"."+Alltrim(Table_schem)+"."+Alltrim(table_name)




					If Not Empty(Nvl(Remarks,""))
						lcToolTip = lcToolTip + Chr(13) + Alltrim(Remarks)
					Endif

					Dimension This.ItemsTables[lnItemsCnt+lnItemsFound,2]
					This.ItemsTables[lnItemsCnt+lnItemsFound,1] = Alltrim(table_name)
					This.ItemsTables[lnItemsCnt+lnItemsFound,2] = Alltrim(lcToolTip)

					If pllAdd
						This.AddItem(Alltrim(table_name), 16, lcToolTip)
					Endif
				Endif
			Endscan

			Use In &lcSqlTables
			If Used(lcAlias)
				Select (lcAlias)
			Endif
		Endif

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Capturo todas as tabelas e alias contidas na select
	*/ plnMode = 0 	-> Tabelas e Alias
	*/ plnMode = 1  -> Tabelas
	*/ plnMode = 2  -> Alias
	*/------------------------------------------------------------------------------------------------
	Procedure GetSqlTablesInCmd
		Lparameters plcWord, plnMode, pllAdd, pllClearArray
		Local lcWord, lcSelect, lnx, lcSqlAlias, lcSqlTable, lnItemsFound, lnItemsCnt

		Set Console Off

		If This.chkTFsql <> "1"
			Return 0
		Endif

		If pllClearArray
			Declare This.ItemsTables[1,2]
			This.ItemsTables[1,1] = ""
			This.ItemsTables[1,2] = ""
		Endif

		lnItemsFound = 0
		lnItemsCnt = Iif(Empty(This.ItemsTables[1,1]), 0, Alen(This.ItemsTables,1))

		*- Se nao estou conectado forço o intellisense a trabalhar com Tables and Alias
		plnMode= Iif(plnMode=2 And Asqlhandles(laCnx) <= 0, 0, plnMode)


		*- valido somente para select, insert, update and delete
		If	( Getwordnum(Lower(This.TextEndBlock),1) == "select" And (" from " $ This.TextEndBlock Or " join " $ This.TextEndBlock) ) Or;
				( Getwordnum(Lower(This.TextEndBlock),1) == "insert" And Getwordnum(Lower(This.TextEndBlock),2) == "into" ) Or;
				( Getwordnum(Lower(This.TextEndBlock),1) == "delete" And Getwordnum(Lower(This.TextEndBlock),2) == "from" ) Or;
				( Getwordnum(Lower(This.TextEndBlock),1) == "update" )

			*- preparo a select para leitura
			lcSelect = Strtran(Strtran(This.TreatWords(Strtran(This.TextEndBlock, " as ", " ",1,-1,1)), "[ ", "["), " ]", "]")
			lcSelect = Strtran(Strtran(lcSelect, ", ", ","), " ,", ",")
			lcSelect = Strtran(lcSelect, "group by", "group_by",1,-1,1)
			lcSelect = Strtran(lcSelect, "order by", "order_by",1,-1,1)
			lcSelect = Strtran(lcSelect, "with (", "(",1,-1,1)
			lcSelect = Strtran(lcSelect, "( nolock )", "",1,-1,1)
			lcSelect = Strtran(lcSelect, "( updlock )", "",1,-1,1)
			lcSelect = Strtran(lcSelect, "( holdlock )", "",1,-1,1)
			lcSelect = Strtran(lcSelect, "( updlock", "",1,-1,1)
			lcSelect = Strtran(lcSelect, "( holdlock", "",1,-1,1)

			*- faco a leitura
			For lnx = 1 To Getwordcount(lcSelect)

				lcWord = Getwordnum(lcSelect, lnx)

				*- considero tabelas depois do comando FROM ou JOIN
				If Inlist(Lower(Alltrim(lcWord)), "from", "join", "into", "update")

					lcSqlTable = Getwordnum(lcSelect, lnx+1)
					If Substr(lcSqlTable,1,1) = "["		&&- large table name
						lnx = lnx + 1
						lcSqlTable = lcSqlTable + " " + Getwordnum(lcSelect, lnx+1)
					Endif

					*- entedo que se o nome da tabela iniciar com "(" trata-se de uma subselect
					If Substr(lcSqlTable,1,1) = "("
						lcSqlAlias = "**subselect**"		&&- devo descobrir onde fecha o parenteses antes do alias
						Loop								&&
					Else
						lcSqlAlias = This.ClearQuotes(Getwordnum(lcSelect, lnx+2))
					Endif

					*- Alias
					If plnMode = 0 Or plnMode = 2
						If This.ChkValidTableOrFieldName(lcSqlAlias)
							If Not Inlist(Lower(lcSqlAlias),"(",")", "where", "inner", "join", "left", "right", "cross", "order_by", "group_by", "full", "union", "on", "set", "values")
								If This.ChkIncremental(plcWord, lcSqlAlias)
									lnItemsFound = lnItemsFound + 1
									lcToolTip = "Table " + lcSqlTable + " As " + Alltrim(lcSqlAlias)

									Dimension This.ItemsTables[lnItemsCnt+lnItemsFound,2]
									This.ItemsTables[lnItemsCnt+lnItemsFound,1] = Alltrim(lcSqlAlias)
									This.ItemsTables[lnItemsCnt+lnItemsFound,2] = Alltrim(lcToolTip)

									If pllAdd
										This.AddItem(lcSqlAlias, 16, lcToolTip)
									Endif
									Loop 	&&- se adicionei o ALIAS entao nao adiciono a tabela relativo ao ALIAS
								Endif
							Endif
						Endif
					Endif

					*- Tables
					If plnMode = 0 Or plnMode = 1
						If This.ChkValidTableOrFieldName(lcSqlTable)
							If Not Inlist(Lower(lcSqlTable),"(",")", "where", "inner", "join", "left", "right", "cross", "order_by", "group_by", "full", "union", "on", "set", "values")
								If This.ChkIncremental(plcWord, lcSqlTable)
									lnItemsFound = lnItemsFound + 1
									lcToolTip = "Table " + lcSqlTable

									Dimension This.ItemsTables[lnItemsCnt+lnItemsFound,2]
									This.ItemsTables[lnItemsCnt+lnItemsFound,1] = Alltrim(lcSqlTable)
									This.ItemsTables[lnItemsCnt+lnItemsFound,2] = Alltrim(lcToolTip)

									If pllAdd
										This.AddItem(lcSqlTable, 16, lcToolTip)
									Endif
								Endif
							Endif
						Endif
					Endif

				Endif

			Endfor

		Endif

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busca os campos de uma tabela do database SGBD - SQL
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetSqlFields
		Lparameters plcTable
		Local lcSqlFields, lnItemsFound, lcAlias, lcToolTip
		Local Array laCnx[1]

		Set Console Off

		If This.chkTFsql <> "1"
			Return 0
		Endif

		lnItemsFound = 0
		lcSqlFields = "tmp"+Sys(2015)

		Try
			If Asqlhandles(laCnx) > 0 And SQLColumns(1, plcTable, "NATIVE", lcSqlFields) = 1
				lcAlias = Alias()

				Select (lcSqlFields)
				Scan
					If This.ChkIncremental(This.LastWord, column_name)
						lnItemsFound = lnItemsFound + 1

						lcToolTip = "Column " + Alltrim(column_name) + ", " + Alltrim(type_name)
						lcToolTip = lcToolTip + Iif(Isnull(sql_datetime_sub), "(" + Alltrim(Str(column_size)) + Iif(Empty(Nvl(decimal_digits,0)), "", ","+Alltrim(Str(decimal_digits))) + ")"  ,"") +  ", "
						lcToolTip = lcToolTip + Iif(Lower(Alltrim(is_nullable))="no","not","") + " null"

						If Not Empty(Nvl(column_def,""))
							lcToolTip = lcToolTip + Chr(13) + "Defaul value: " + Alltrim(Strtran(column_def, Chr(0), " "))
						Endif

						If Not Empty(Nvl(Remarks,""))
							lcToolTip = lcToolTip + Chr(13) + Alltrim(Remarks)
						Endif

						lcToolTip = lcToolTip + Chr(13) + "Table " + Alltrim(table_cat)+"."+Alltrim(Table_schem)+"."+Alltrim(table_name)

						This.AddItem(Alltrim(column_name), 17, lcToolTip)
					Endif
				Endscan

				Use In &lcSqlFields
				If Used(lcAlias)
					Select (lcAlias)
				Endif
			Endif
		Catch
		Endtry

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busca os campos de todas as tabelas contidas na select-sql
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetSqlFieldsInAllTablesCmd
		Local lnLines, lnx, lnItemsFound

		Set Console Off

		* obtenho todas as tabelas contidas na select
		This.IncrementalResult =  .F.
		lnLines = This.GetSqlTablesInCmd(This.LastWord, 1, .F., .T.)		&&- tabelas e alias existentes na select-sql
		This.IncrementalResult =  .T.

		If lnLines = 0
			Return 0
		Endif

		* obtenho os campos de todas as tabelas contidas na select
		lnItemsFound = 0
		For lnx = 1 To lnLines
			lnItemsFound = lnItemsFound + This.GetSqlFields(This.ItemsTables[lnx,1])
		Endfor

		Return lnItemsFound
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ Busca o WITH...ENDWITH de origem de chamada do objeto
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetWithHierarchy
		Local lcSafety, lcThisCode, lnCurrentLine, lnTotClasses, lnx, lnz, lnCurrentLine,;
			lcClassName, lcBaseClass, lnClassLine, lcTextWord1, lcTextWord2, lcWith, lnEndWith, lcTextLine,;
			lcLastWord, lcObjReference, lcObjName, lnRowFound

		Set Console Off

		This.WithReference = ""

		Local Array laProcClass[1,4]
		laProcClass[1,1] = ""

		Local Array laCodePrg[1]

		*- se for classe verifico tambem os metodos e propriedades que inclui na classe no programa em write-time.
		*- Copio todo o codigo e salvo num arquivo temporario
		lcSafety = Set("Safety")
		Set Safety Off
		lcThisCode = This.GetText(0,130000)			&&- capturo o texto da janela atual
		Strtofile(lcThisCode, This.Tmpfile)
		Set Safety &lcSafety

		*- coloco todo o programa em um array
		Alines(laCodePrg, lcThisCode)
		lcThisCode = ""

		*- no. da linha atual
		lnCurrentLine = This.GetLineNo()
		If lnCurrentLine = -1
			lnCurrentLine = 0
		Else
			lnCurrentLine = lnCurrentLine - 1
		Endif

		*- Se nao estou no editor do form ou class designer
		If This.EditorSource <> 10
			*- verifico em qual classe estou posicionado
			Aprocinfo(laProcClass, This.Tmpfile, 1)		&&- Classes do arquivo
			lnTotClasses = Alen(laProcClass,1)
			If lnTotClasses	= 1 And Empty(laProcClass[1,1])
				lnTotClasses = 0
			Endif

			This.ProcBaseClass = ""
			This.ProcClass = ""

			If lnTotClasses > 0
				For lnx = 1 To lnTotClasses
					lnz = (lnTotClasses+1)-lnx
					If lnCurrentLine > laProcClass[lnz,2]
						lcClassName = laProcClass[lnz,1]
						lcBaseClass = laProcClass[lnz,3]
						lnClassLine = laProcClass[lnz,2]
						This.ProcBaseClass = lcBaseClass
						This.ProcClass = lcClassName
						Exit
					Endif
				Endfor
			Endif

			*- Form or class editor
		Else
			Local loControl
			Local Array laControl[1]
			laControl[1] = .Null.
			Aselobj(laControl, 1)
			This.ProcBaseClass = ""
			This.ProcClass = ""

			If Vartype(laControl[1]) = "O"
				*- capturo os textos de alguns metodos de um form ou toolbar
				loControl = laControl[1]
				lcClassName = loControl.Name
				lcBaseClass = Iif(Lower(loControl.BaseClass)="olecontrol", loControl.OleClass, loControl.BaseClass)
				lnClassLine = 0
				This.ProcBaseClass = lcBaseClass
				This.ProcClass = lcClassName
			Endif
		Endif


		*- busca a ultima palavra antes do "."
		lcTextLine = This.TreatLine(This.GetTextLine())
		lcLastWord = Getwordnum(lcTextLine,Getwordcount(lcTextLine))
		lcLastWord = Alltrim( Iif("."$lcLastWord, Substr(lcLastWord,1,Len(lcLastWord)-1), lcLastWord) )

		If Alltrim(lcLastWord) == Alltrim(lcTextLine)
			lcLastWord = ""
		Endif

		*- inicio a composição do with... endwith
		lcWith = ""
		lnEndWith = 0
		For lnx = 1 To lnCurrentLine
			lnz = (lnCurrentLine+1)-lnx

			*- capturo o conteudo da linha
			lcTextLine = Lower(This.TreatLine(laCodePrg[lnz]))
			If Empty(lcTextLine)
				Loop
			Endif

			lcTextWord1 = Getwordnum(lcTextLine,1)
			lcTextWord2 = Getwordnum(lcTextLine,2)

			If Substr(lcTextWord1,1,4) == "endw" And Empty(lcTextWord2)			&&- conto os aninhamentos
				lnEndWith = lnEndWith + 1
			Endif

			*- with dentro de outro with
			If lcTextWord1 = "with" And Substr(lcTextWord2,1,1) = "."
				If lnEndWith = 0
					lcWith = lcTextWord2 + lcWith
				Else
					lnEndWith = lnEndWith - 1
					lnEndWith = Iif(lnEndWith<0,0,lnEndWith)
				Endif
			Endif

			*- se cheguei no fim da classe ou procedure
			*- e nao achei o with, entao saio.
			If This.EditorSource <> 10
				If 	Getwordcount(lcTextLine) = 1 And Inlist(Substr(lcTextWord1,1,4),"endp","endf") Or;
						getwordcount(lcTextLine) = 1 And Substr(lcTextWord1,1,5) == "endde" Or;
						this.IsProc(lcTextLine)

					lcWith = ""
					lcLastWord = ""
					Exit
				Endif
			Endif

			*- cheguei no 1o. with do objeto
			If lcTextWord1 == "with" And Substr(lcTextWord2,1,1) <> "."
				If lnEndWith = 0					&&- o IntelliSense deve abrir dentro do with/endwith
					lcWith = lcTextWord2 + lcWith
				Else
					lcWith = ""
					lcLastWord = ""
				Endif
				Exit
			Endif
		Endfor

		lcObjReference = lcWith + lcLastWord
		This.WithReference = lcObjReference

		*- verifico se é um objeto inserido via "ADD OBJECTS" no "DEFINE CLASS"
		If This.EditorSource <> 10
			If Substr(lcObjReference,1,5) == "this." And Occurs(".",lcObjReference) = 1
				This.GetProcInfo(4,0,.F.)
				lcObjName = Substr(lcObjReference,At(".",lcObjReference)+1)
				lnRowFound = Ascan(This.ItemsObjects, lcObjName, -1,-1, 1, 15)
				If lnRowFound > 0
					This.ControlClassName = Alltrim(This.ItemsObjects[lnRowFound,2])
					This.ControlOleClass = Alltrim(This.ItemsObjects[lnRowFound,3])
				Else
					This.ControlClassName = ""
					This.ControlOleClass = ""
				Endif
			Endif
		Endif

		Return lcObjReference
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busco as PMEs de um objeto em memoria (run-time) ou de uma classe nativa do vfp.
	*/ pluObjClass..: é o nome da classe ou objeto da classe em memoria
	*/ pllAdd.......: .t. - indica que os itens encontrados serão incluidos no IntelliSense e no array this.items
	*/			      .f. - serão incluídos somente no array this.items
	*/ pllClearArray: .t. - limpo o array this.items
	*/				  .f. - mantenho this.items com as informações encontradas por outros metodos
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetMembers
		Lparameters pluObjClass, pllAdd, pllClearArray
		Local lcAlias, llVFPnativeClass, lnItemsFound, lnRowFound

		Set Console Off

		pllAdd = Iif(Parameters()<=1, .T., pllAdd)
		lcAlias = Alias()

		If Not Used("foxcodeplus2")
			Use foxcodeplus2 Alias foxcodeplus2 In 0
		Endif

		If pllClearArray
			Declare This.Items[1,4]
			This.Items[1,1] = ""
		Endif

		Select foxcodeplus2
		Set Order To typecode

		*- se consigo obter "C" conforme abaixo é pq é um classe nativa do VFP mesmo que derivada.
		*- se nao consigo é pq pode ser um OLE Control ou as propriedades abaixo estao protected or hidden.
		llVFPnativeClass = (Type("pluObjClass.baseclass")="C" And Type("pluObjClass.class")="C") Or Type("pluObjClass")="C"

		*- capturo os membros da classe de todas as formas possiveis
		If llVFPnativeClass
			This.TreatMembers(pluObjClass, 1, "G")		&&- Public
			This.TreatMembers(pluObjClass, 1, "N")		&&- Native
			This.TreatMembers(pluObjClass, 1, "U")		&&- User Define
			This.TreatMembers(pluObjClass, 1, "C")		&&- Changed
			This.TreatMembers(pluObjClass, 1, "I")		&&- Inherited
			This.TreatMembers(pluObjClass, 1, "B")		&&- Base
			This.TreatMembers(pluObjClass, 1, "P") 		&&- Protected*
			This.TreatMembers(pluObjClass, 1, "H")		&&- Hidden*
			This.TreatMembers(pluObjClass, 1, "R")		&&- Read Only*
		Else
			This.TreatMembers(pluObjClass, 3, "")		&&- Only for OLE
		Endif

		lnItemsFound = Iif(!Empty(This.Items[1,1]), Alen(This.Items,1), 0)

		Use In foxcodeplus2
		If Used(lcAlias)
			Select (lcAlias)
		Endif

		*- verifico se a propriedade _MemberData existe e
		*- caso exista faço o tratamento dos captions
		If Vartype(pluObjClass) = "O"
			lnRowFound = Ascan(This.Items, "_MemberData", -1,-1, 1, 15)
			If lnRowFound > 0 And This.Items[lnRowFound,2] = 7				&&- OBS: Para objetos em run-time é impossivel obter o _memberdata se o mesmo estiver hidden ou protected
				This.TreatMemberData(pluObjClass._MemberData)
			Endif
		Endif

		*- adiciono todos os items encontrados ao IntelliSense
		If pllAdd
			If lnItemsFound > 0
				For lnx = 1 To lnItemsFound
					This.AddItem(This.Items[lnx,1], This.Items[lnx,2], This.Items[lnx,3])
				Endfor
			Endif
		Endif

		Return lnItemsFound


		*/------------------------------------------------------------------------------------------------
		*/ used only within method this.GetMembers
		*/------------------------------------------------------------------------------------------------
	Protected Procedure TreatMembers
		Lparameters pluObjClass, plnMode, plcFlag
		Local lcPMEtype, lnNewRow, lnx, lcItem, lnImage, lcToolTip, lcAuxItem, lcType, lnNewPos, lcCaption, lcOleClass, loControl
		Local Array laMembersX[1,4]
		Store "" To laMembersX[1,4]

		Set Console Off

		Do Case
			Case plcFlag = "P"	&&- Protected*
				lcPMEtype = "Protected"

			Case plcFlag = "H"	&&- Hidden*
				lcPMEtype = "Hidden"

			Case plcFlag = "R"	&&- Read Only*
				lcPMEtype = "Read Only"

			Otherwise
				lcPMEtype = " "
		Endcase

		If Alen(This.Items,1) = 1 And Empty(This.Items[1,1])
			lnNewRow = 0
		Else
			lnNewRow = Alen(This.Items,1)
		Endif

		Try
			If Not Empty(plcFlag)
				Amembers(laMembersX, pluObjClass, plnMode, plcFlag)
			Else
				Amembers(laMembersX, pluObjClass, plnMode)
			Endif
		Catch
		Endtry


		*- inicio a analise das pmes encontradas
		If Not Empty(laMembersX[1,1])
			For lnx = 1 To Alen(laMembersX,1)
				lcItem = ""
				lnImage = 0
				lcToolTip = ""

				lcAuxItem = laMembersX[lnx,1]
				lcType = Iif(Substr(Lower(laMembersX[lnx,2]),1,8)="property", "property", Lower(laMembersX[lnx,2]))
				lcType = Iif(lcType="object","control",lcType)

				*- se for classes nativas do VFP faço o tratamento dos captions e tips
				If Inlist(plnMode, 0, 1, 2)
					*- se o membro do objeto for algum control obtenho mais informacoes para o tooltip
					*- tambem faço o tratamento do caption de acordo como foi formatado a propriedade name.
					If lcType == "control"
						If Type("pluObjClass") = "O"
							loControl = Evaluate("pluObjClass."+lcAuxItem)
							lcItem = Iif(Pemstatus(loControl,"name",5) And Vartype(loControl.Name) = "C", loControl.Name, Proper(lcAuxItem))
							lcOleClass = Iif(Pemstatus(loControl,"oleclass",5) And Vartype(loControl.OleClass) = "C", loControl.OleClass, "")

							If Pemstatus(loControl,"name",5) And Pemstatus(loControl,"parent",5) And Pemstatus(loControl,"baseclass",5)
								lcToolTip = "Control "+Lower(loControl.Parent.Name)+"."+lcItem+" Class "+loControl.Class + Iif(!Empty(lcOleClass), " OleClass "+lcOleClass,"") + Chr(13) +;
									"BaseClass: "+loControl.BaseClass + Chr(13) +;
									"ClassLibrary: "+Iif(Empty(loControl.ClassLibrary), "(None)", loControl.ClassLibrary)
							Else
								lcToolTip = "Object " + lcItem + " BaseClass Empty"
							Endif

							If Pemstatus(loControl,"caption",5) And Vartype(loControl.Caption) = "C"
								lcCaption = Alltrim(loControl.Caption)
								lcToolTip = lcToolTip + Chr(13) + "Caption: "+Iif(Len(lcCaption)>40, Substr(lcCaption,1,40)+"...", lcCaption)
							Endif
						Else
							lcItem = Proper(laMembersX[lnx,1])
							lcToolTip = "Unknown control"
						Endif


						*- diferente de controles, formato o tooltip com base no foxcodeplus2.dbf
						*- tambem faço o tratamento do caption
					Else
						If Seek(Upper(Substr(lcType,1,1)) + Lower(Padr(lcAuxItem,30)), "foxcodeplus2")
							lcItem = Alltrim(foxcodeplus2.Code)
							lcToolTip = lcPMEtype + " " + Proper(lcType) + " " + lcItem + ;
								iif(Upper(Substr(lcType,1,1)) $ "M//E", "(" + Iif(!Empty(foxcodeplus2.Param), Alltrim(foxcodeplus2.Param), " ") + ")" , "") + ;
								chr(13) + Alltrim(foxcodeplus2.tip)
						Else
							lcItem = Proper(lcAuxItem)
							lcToolTip = lcPMEtype + " " + Proper(lcType) + " " + lcItem + ;
								iif(Upper(Substr(lcType,1,1)) $ "M//E", "( )" , "")
						Endif
					Endif

					*- plnMode = 3
				Else
					lcItem = Iif(Isupper(Right(lcAuxItem,1)), Proper(lcAuxItem), lcAuxItem)
					lcToolTip = Iif(!Empty(lcPMEtype), lcPMEtype + " ", "") + Proper(lcType) + " " + lcItem + ;
						iif(Upper(Substr(lcType,1,1)) $ "M//E", "(" + Iif(!Empty(laMembersX[lnx,3]), laMembersX[lnx,3], " ") + ")" , "") + ;
						iif(!Empty(laMembersX[lnx,4]), Chr(13) + laMembersX[lnx,4], "")
				Endif


				*- imagem
				Do Case
						*- class
					Case lcType == "class"
						lnImage = 1

						*- control
					Case lcType == "control"
						lnImage = 13

						*- normal property
					Case lcType = "property" And Not plcFlag $ "H//P"
						lnImage = 7

						*- hidden property
					Case lcType == "property" And plcFlag = "H"
						lnImage = 8

						*- protected property
					Case lcType == "property" And plcFlag = "P"
						lnImage = 9

						*- methode
					Case lcType == "method" And Not plcFlag $ "H//P"
						lnImage = 4

						*- Hidden methode
					Case lcType == "method" And plcFlag = "H"
						lnImage = 5

						*- Protected methode
					Case lcType == "method" And plcFlag = "P"
						lnImage = 6

						*- event
					Case lcType == "event" And Not plcFlag $ "H//P"
						lnImage = 3

						*- Hidden event
					Case lcType == "event" And plcFlag = "H"
						lnImage = 14

						*- Protected event
					Case lcType == "event" And plcFlag = "P"
						lnImage = 15


						*- tipo nao previsto (sem imagem)
					Otherwise
						lnImage = 0
				Endcase

				lcToolTip = Strtran(Strtran(lcToolTip, "((","("), "))", ")")

				*- inclusao das PMEs encontradas
				lnRowFound = This.SeekItem(lcItem, lnImage)
				If lnRowFound = 0
					lnNewRow = lnNewRow + 1
					Dimension This.Items[lnNewRow,4]
					This.Items[lnNewRow,1] = lcItem
					This.Items[lnNewRow,2] = lnImage
					This.Items[lnNewRow,3] = lcToolTip
					This.Items[lnNewRow,4] = ""
				Else
					If Inlist(plcFlag, "P", "H", "R")
						This.Items[lnRowFound,2] = lnImage
						This.Items[lnRowFound,3] = lcToolTip
					Endif
				Endif

			Endfor
		Endif

		Return .T.
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ Aplica o caption dos itens do IntelliSense conforme definido no XML da propriedade _MemberData
	*/------------------------------------------------------------------------------------------------
	Protected Procedure TreatMemberData
		Lparameters plcMemberDataXML
		Local lnRowFound, lnx, lcMemberName, lcMemberDisplay, lcMemberType
		Local Array laMemberData[1]

		Set Console Off

		plcMemberDataXML = Nvl(plcMemberDataXML,"")
		plcMemberDataXML = Evl(plcMemberDataXML,"")
		plcMemberDataXML = Strtran(plcMemberDataXML, ">", ">"+Chr(13))

		lnRowFound = Ascan(This.Items, "_MemberData", -1,-1, 1, 15)
		If lnRowFound > 0
			This.Items[lnRowFound,3] = "XML Metadata for customizable properties, methods and events."
		Endif

		If Alines(laMemberData, plcMemberDataXML) > 0
			For lnx = 1 To Alen(laMemberData,1)

				If "<memberdata" $ laMemberData[lnx]
					laMemberData[lnx] = Strtran( Strtran( Strtran( Strtran( Strtran(laMemberData[lnx], "=", " = "), ['], ["]), "[", '"'), "]", '"'), "  ", " ")
					laMemberData[lnx] = Strtran(laMemberData[lnx], 'name = "', 'Name = "', -1, -1, 1)
					laMemberData[lnx] = Strtran(laMemberData[lnx], 'display = "', 'Display = "', -1, -1, 1)
					laMemberData[lnx] = Strtran(laMemberData[lnx], 'type = "', 'Type = "', -1, -1, 1)

					lcMemberName = Strextract(laMemberData[lnx],[Name = "],["])					&&- pme name
					lcMemberDisplay = Strextract(laMemberData[lnx],[Display = "],["]) 			&&- display caption to the IntelliSense
					lcMemberType = Lower(Strextract(laMemberData[lnx],[Type = "],["]))			&&- check if is property, event or method

					*- procuro pela pme e atribuo o caption que foi definido na propriedade _MemberData
					lnRowFound = 1
					Do While .T.

						lnRowFound = Ascan(This.Items, lcMemberName, lnRowFound,-1, 1, 15)
						If lnRowFound > 0
							Do Case
								Case lcMemberType = "property"
									If Not Inlist(This.Items[lnRowFound,2],7,8,9,10)
										lnRowFound = lnRowFound + 1
										If lnRowFound > Alen(This.Items,1)
											Exit
										Endif
										Loop
									Endif

								Case lcMemberType = "method"
									If Not Inlist(This.Items[lnRowFound,2],4,5,6)
										lnRowFound = lnRowFound + 1
										If lnRowFound > Alen(This.Items,1)
											Exit
										Endif
										Loop
									Endif

								Case lcMemberType = "event"
									If Not Inlist(This.Items[lnRowFound,2],3,14,15)
										lnRowFound = lnRowFound + 1
										If lnRowFound > Alen(This.Items,1)
											Exit
										Endif
										Loop
									Endif

								Otherwise
									Exit
							Endcase

							If Lower(This.Items[lnRowFound,1]) == Lower(lcMemberDisplay)
								This.Items[lnRowFound,1] = lcMemberDisplay
								If lcMemberType = "property"
									This.Items[lnRowFound,3] = Strtran(This.Items[lnRowFound,3], "roperty " + lcMemberDisplay, "roperty " + lcMemberDisplay, -1, -1, 1)
								Else
									This.Items[lnRowFound,3] = Strtran(This.Items[lnRowFound,3], lcMemberDisplay + "(", lcMemberDisplay + "(", -1, -1, 1)
								Endif
								This.Items[lnRowFound,3] = This.Items[lnRowFound,3] + Chr(13) + "(_MemberData capitalization)"
							Endif
						Endif

						Exit
					Enddo
				Endif

			Endfor

		Endif
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Analisa todos os programas invocados pelo comando SET PROCEDURE TO que esta em memoria (run-time)
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetSetProcInfoPrgRunTime
		Return This.GetSetProcInfoPrg("SET PROCEDURE TO "+Set("Procedure"), .T.)
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busca em um PRG invacado pelo comando SET PROCEDURE TO especifico as classes e funcoes procedurais
	*/ plcSetProcLine -> linha que contem o comando SET PROCEDURE TO MyProgram1.PRG, MyProgram2.PRG ADDITIVE
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetSetProcInfoPrg
		Lparameters plcSetProcLine, pllAdd
		Local lcPrgFile, lnRow, lnx, lnTotProc, lnBackLine, lnItemsFound
		Local Array laProc[1,4]
		Local Array laCodePrg[1]
		Local Array laDir[1]

		Set Console Off

		If Getwordcount(plcSetProcLine)<=3
			Return 0
		Endif

		plcSetProcLine = Lower(Substr(plcSetProcLine, At(" ",plcSetProcLine,3)+1))
		plcSetProcLine = "<FILE>"+Strtran(plcSetProcLine,",","</FILE><FILE>")+"</FILE>"
		lnRow = Iif(Empty(This.Items[1,1]), 0, Alen(This.Items,1))
		lnItemsFound = 0

		*- verifico todos os PRGs declarados no "set procedure to"
		Do While .T.
			lcPrgFile = Strextract(plcSetProcLine,"<FILE>","</FILE>")
			plcSetProcLine = Strtran(plcSetProcLine, "<FILE>"+lcPrgFile+"</FILE>", "")
			lcPrgFile = Alltrim(lcPrgFile)

			If Empty(lcPrgFile) Or Substr(lcPrgFile,1,4) = "addi"
				Exit
			Endif

			lcPrgFile = Forceext(lcPrgFile,"PRG")

			If Not This.GetFilePath(@lcPrgFile)
				Loop
			Endif

			Alines(laCodePrg, Filetostr(lcPrgFile))
			If Empty(laCodePrg[1]) And Alen(laCodePrg,1) <= 1
				Loop
			Endif

			lnTotProc = Aprocinfo(laProc, lcPrgFile, 0)

			*- se achou algo no PRG, verifico oque pode ser incluido no IntelliSense
			If lnTotProc > 0
				For lnx = 1 To lnTotProc
					If laProc[lnx,3] = "Procedure" And (Not "." $ laProc[lnx,1]) And This.chkFC = "1"

						If Not This.ChkIncremental(This.LastWord, laProc[lnx,1])
							Loop
						Endif

						*- corrijo resultado da procinfo() qdo a funcao é declarada com quebra de linha
						For lnBackLine = 1 To 99
							If laProc[lnx,2] > 0 And Not This.IsProc(laCodePrg[laProc[lnx,2]])
								laProc[lnx,2] = laProc[lnx,2] - 1
							Else
								Exit
							Endif
						Endfor

						*- incluo a funcao no array do IntelliSense
						lnItemsFound = lnItemsFound + 1
						lnRow = lnRow + 1
						Dimension This.Items[lnRow,4]
						This.Items[lnRow,1] = laProc[lnx,1]
						This.Items[lnRow,2] = 19
						This.Items[lnRow,3] = laProc[lnx,1] + "("+This.GetParameters(@laCodePrg, laProc[lnx,2])+")" + This.GetSummary(@laCodePrg, laProc[lnx,2]) + Chr(13) + "File " + Upper(lcPrgFile)
						This.Items[lnRow,4] = "SET PROCEDURE TO <FILE>"+Upper(lcPrgFile)+"</FILE>"

						If pllAdd
							This.AddItem(This.Items[lnRow,1], This.Items[lnRow,2], This.Items[lnRow,3])
						Endif

						Loop
					Endif

					If laProc[lnx,3] = "Class" And This.chkControl = "1"

						If Not This.ChkIncremental(This.LastWord, laProc[lnx,1])
							Loop
						Endif

						*- incluo a classe no array do IntelliSense
						lnItemsFound = lnItemsFound + 1
						lnRow = lnRow + 1
						Dimension This.Items[lnRow,4]
						This.Items[lnRow,1] = Getwordnum(laProc[lnx,1],1)
						This.Items[lnRow,2] = 1
						This.Items[lnRow,3] = "Class " + This.Items[lnRow,1] + " as baseclass " + Getwordnum(laProc[lnx,1],3) + Chr(13) + "File " + Upper(lcPrgFile)
						This.Items[lnRow,4] = "SET PROCEDURE TO <FILE>"+Upper(lcPrgFile)+"</FILE>"

						If pllAdd
							This.AddItem(This.Items[lnRow,1], This.Items[lnRow,2], This.Items[lnRow,3])
						Endif

						Loop
					Endif
				Endfor
			Endif
		Enddo

		Return lnItemsFound
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ Busca em um VCX invacado pelo comando SET CLASSLIB TO as classes contidas no mesmo
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetSetClassInfoVcx
		Lparameters plcFileVcx
		Local lnRow, lnx, lnTotClasses, lnItemsFound
		Local Array laClasses[1,11]

		Set Console Off

		plcFileVcx = Forceext(plcFileVcx,"VCX")
		If Not This.GetFilePath(@plcFileVcx)
			Return 0
		Endif

		lnRow = Iif(Empty(This.Items[1,1]), 0, Alen(This.Items,1))
		lnItemsFound = 0
		lnTotClasses = Avcxclasses(laClasses, plcFileVcx)

		*- incluo a classe no array do IntelliSense
		If lnTotClasses > 0
			For lnx = 1 To lnTotClasses
				If This.chkControl = "1"

					*if this.IncrementalResult
					If Not This.ChkIncremental(This.LastWord, laClasses[lnx,1])
						Loop
					Endif
					*endif

					lnItemsFound = lnItemsFound + 1
					lnRow = lnRow + 1
					Dimension This.Items[lnRow,4]
					This.Items[lnRow,1] = laClasses[lnx,1]
					This.Items[lnRow,2] = 1
					This.Items[lnRow,3] = "Class " + This.Items[lnRow,1] + " as baseclass " + laClasses[lnx,3] + Chr(13) + ;
						"File " + Upper(plcFileVcx) + ;
						iif(!Empty(laClasses[lnx,4]), Chr(13) + "Baseclass file " + laClasses[lnx,4], "")  + ;
						iif(!Empty(laClasses[lnx,8]), Chr(13) + Chr(13) + laClasses[lnx,8], "")

					This.Items[lnRow,4] = "SET CLASSLIB TO <FILE>"+Upper(plcFileVcx)+"</FILE>"
				Endif
			Endfor
		Endif

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busca em um PRG especifico as pmes de um classe e as pmes de sua baseclass
	*/ pllAdd.......: .t. - indica que os itens encontrados serão incluidos no IntelliSense e no array this.items
	*/			      .f. - serão incluídos somente no array this.items
	*/ pllClearArray: .t. - limpo o array this.items
	*/				  .f. - mantenho this.items com as informações encontradas por outros metodos
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetProcInfoPrg
		Lparameters plcClassName, plcFilePrg, pllAdd, pllClearArray

		Local lnClassLine, lcBaseClass, lnx, lnClassLine, lcTextLine, lnWordCount, lcTextWord1, lcTextWord2, lnImage,;
			lcOleClass, lcClassControl, lcControlName, lnCountPropertiesInLine, lnz, llIsAnArray, lnRowFound, llHasChangedProperty,;
			lcProperty, lcToolTip, lnItemsFound, lnItemRow, lcProcType, lcMemberDataXML, lcSetTalk, lcSetNotify, lcSetMessage

		Local Array laCodePrg[1], laProc[1,4], laDir[1,5]

		Set Console Off

		*- se existe caracters especiais no nome do arquivo ou da classe desconsidero a analise do IntelliSense
		*- nesse caso possivelmente existe programacao para composicao dos nomes
		If "&" $ plcFilePrg Or "(" $ plcFilePrg Or "+" $ plcFilePrg Or ".." $ plcFilePrg	 &&or "'" $ plcFilePrg or "[" $ plcFilePrg or '"' $ plcFilePrg
			Return 0
		Endif

		If "&" $ plcClassName Or "(" $ plcClassName Or "+" $ plcClassName Or "." $ plcClassName Or "'" $ plcClassName Or "[" $ plcClassName Or '"' $ plcClassName
			Return 0
		Endif

		*-
		plcClassName = Lower(plcClassName)
		lcBaseClass = ""
		lnItemsFound = 0
		lcMemberDataXML = ""

		If pllClearArray
			Declare This.Items[1,4]				&&- limpo o array.
			This.Items[1,1] = ""
		Endif

		If Not This.GetFilePath(@plcFilePrg)
			This.ShowErrorWriteTime(1, Upper(Justfname(plcFilePrg))) 	&&- file doesn't exist
			Return 0
		Endif

		*- verifico se a classe existe no prg e se existir capturo o baseclass
		Aprocinfo(laProc, plcFilePrg, 0)
		For lnx = 1 To Alen(laProc,1)
			If Lower(laProc[lnx,3]) == "class" And Alltrim(Substr(Lower(laProc[lnx,1]),1,At(" ",laProc[lnx,1]))) == plcClassName
				lcBaseClass = Getwordnum(laProc[lnx,1],3)
				lnClassLine = laProc[lnx,2]
				Exit
			Endif
		Endfor

		If Empty(lcBaseClass)
			This.ShowErrorWriteTime(1733, Iif(Empty(plcClassName), '"UNKNOW"', Upper(plcClassName))) 	&&- class doesn't exist
			Return 0
		Endif

		*- coloco o prg num array
		Alines(laCodePrg, Filetostr(plcFilePrg))
		If Empty(laCodePrg[1]) And Alen(laCodePrg,1) <= 1
			Return 0
		Endif

		*- adiciono as pmes do baseclass
		*- .f. indico que não adicionarei ao IntelliSense nesse momento, somente ao array this.items
		lnItemsFound = This.GetMembers(lcBaseClass,.F.,.F.)			&&-PMEs nativas da classe

		*- inicio a varredura por propriedades
		For lnx = lnClassLine+1 To Alen(laCodePrg,1)
			lcTextLine = This.TreatLine(laCodePrg[lnx], @lnx, @laCodePrg)
			lcTextLine = Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(lcTextLine,"="," = "), "(","["), ")","]"), "  "," "), ", ",","), " ,",",")
			lnWordCount = Getwordcount(lcTextLine)
			lcTextWord1 = Getwordnum(Lower(lcTextLine),1)
			lcTextWord2 = Getwordnum(Lower(lcTextLine),2)

			If Empty(lcTextLine)
				Loop
			Endif

			*- fim das propriedades
			If 	lnWordCount  = 1 And Inlist(Substr(lcTextWord1,1,4),"endd","endp","endf") Or;
					lnWordCount  = 2 And Inlist(Substr(lcTextWord1,1,4),"proc","func") Or;
					lnWordCount >= 3 And Inlist(Substr(lcTextWord1,1,4),"hidd","prot") And Inlist(Substr(lcTextWord2,1,4),"proc","func")
				Exit
			Endif

			*- capturo a propriedade
			Do Case
					*- capturo o conteudo do _MemberData
				Case lcTextWord1 == "_memberdata"
					lcMemberDataXML = Substr(lcTextLine, At("=",lcTextLine)+1)
					lcTextLine = "_MemberData"
					lnImage = 7

					*- normal property
				Case (lnWordCount = 3 And Alltrim(lcTextWord2) = "=")
					lcTextLine = Getwordnum(lcTextLine,1)
					lnImage = 7

					*- normal array property
				Case Inlist(Substr(lcTextWord1,1,4),"decl","dime") And Alltrim(lcTextWord2) <> "="
					lcTextLine = Substr(lcTextLine, At(" ",lcTextLine)+1)
					lnImage = 7

					*- hidden property
				Case Substr(lcTextWord1,1,4) == "hidd" And Alltrim(lcTextWord2) <> "="
					lcTextLine = Substr(lcTextLine, At(" ",lcTextLine)+1)
					lnImage = 8

					*- protected property
				Case Substr(lcTextWord1,1,4) == "prot" And Alltrim(lcTextWord2) <> "="
					lcTextLine = Substr(lcTextLine, At(" ",lcTextLine)+1)
					lnImage = 9

					*- controls
				Case Lower(Substr(lcTextLine,1,8)) == "add obje" And " as " $ Lower(lcTextLine)
					Local lcControlName, lcClassControl, lcOleClass
					lcOleClass = ""

					If Lower(Getwordnum(lcTextLine,3)) = "protected"
						lcClassControl = This.ClearQuotes(Getwordnum(lcTextLine,6))
						lcControlName = This.ClearQuotes(Getwordnum(lcTextLine,4))
					Else
						lcClassControl = This.ClearQuotes(Getwordnum(lcTextLine,5))
						lcControlName = This.ClearQuotes(Getwordnum(lcTextLine,3))
					Endif

					*- with oleclass = "COMCTL.ProgCtrl.1"
					If Lower(lcClassControl) == "olecontrol"
						lcOleClass = ""
						If "with oleclass" $ Lower(lcTextLine)
							lcOleClass = Substr(lcTextLine, At("with oleclass",lcTextLine))
							lcOleClass = Alltrim(Substr(lcOleClass, At("=",lcOleClass)+1))
							lcOleClass = Iif(" "$lcOleClass, Substr(lcOleClass,1,At(" ",lcOleClass)-1), lcOleClass)
							lcOleClass = This.ClearQuotes(lcOleClass)
						Endif
						lnImage = 20	&&- Adicionei um controle que é uma activex
					Else
						lnImage = 1		&&- Adicionei um controle com uma classe do VFP
					Endif

					lcTextLine = lcControlName

					*- não é uma propriedade valida (podem ser outros comandos)
				Otherwise
					Loop
			Endcase


			*- verifico quantas propriedades foram declaradas na mesma linha
			lnCountPropertiesInLine = Iif("," $ lcTextLine, Occurs(",",lcTextLine)+1, 1)

			*- capturo a propriedade
			For lnz = 1 To lnCountPropertiesInLine
				lcProperty = Getwordnum(lcTextLine,lnz,",")

				If "." $ lcProperty
					Loop
				Endif

				*- se for uma propriedade tipada
				lcPropertyType = ""
				If " as " $ lcProperty
					lcPropertyType = Alltrim(Substr(lcProperty, At(" as ",lcProperty)+4))
					lcProperty = Substr(lcProperty, 1, At(" ",lcProperty)-1)
				Endif

				*- se a propriedade for um array
				llIsAnArray = .F.
				If "[" $ lcProperty
					llIsAnArray = .T.
					lcProperty = Substr(lcProperty, 1, At("[",lcProperty)-1)
				Endif

				If Empty(lcProperty) Or "[" $ lcProperty Or "," $ lcProperty Or "]" $ lcProperty			&&- previno propriedades invalidas
					Loop
				Endif

				*- nao incluo o item se o mesmo ja existir e sendo do mesmo tipo
				*- neste caso verifico somente se o item passou a ser hidden ou protected
				llHasChangedProperty = .F.
				lnRowFound = This.SeekItem(lcProperty, lnImage)
				If lnRowFound > 0
					If This.Items[lnRowFound,2] = lnImage
						Loop
					Else
						*- propriedade já incluída no array this.Items porem, agora nessa linha, incluída como Protected ou Hidden
						*- dessa forma mudo a imagem do IntelliSense para Hidden ou Protected somente se for uma propriedade
						If "As Array" $ This.Items[lnRowFound,3] And This.Items[lnRowFound,2] = 7
							This.Items[lnRowFound,2] = lnImage
							llHasChangedProperty = .T.
							llIsAnArray = .T.
						Else
							Loop
						Endif
					Endif
				Endif


				*- tooltip para a propriedade capturada
				Do Case
						*- Ole control
					Case lnImage = 1
						lcToolTip = "Control " + lcControlName + " Class " + lcClassControl

						*- normal property
					Case lnImage = 7
						lcToolTip = "Property " + lcProperty + Iif(llIsAnArray," As Array", "") + Iif(!Empty(lcPropertyType), " As " + Proper(lcPropertyType), "")

						*- hidden property
					Case lnImage = 8
						lcToolTip = "Hidden Property " + lcProperty + Iif(llIsAnArray," As Array", "") + Iif(!Empty(lcPropertyType), " As " + Proper(lcPropertyType), "")

						*- protected property
					Case lnImage = 9
						lcToolTip = "Protected Property " + lcProperty + Iif(llIsAnArray," As Array", "") + Iif(!Empty(lcPropertyType), " As " + Proper(lcPropertyType), "")

						*- Ole control
					Case lnImage = 20
						lcToolTip = "Control " + lcControlName + " Class " + lcClassControl + Iif(!Empty(lcOleClass), " OleClass "+lcOleClass,"")
				Endcase

				*- se a propriedade é comentada com "&&&" a frente coloco o comentario no tooltip
				If "&"+"&"+"&" $ laCodePrg[lnx]
					lcToolTip = lcToolTip + Chr(13) + Alltrim(Strtran(Strtran(Substr(laCodePrg[lnx],At("&"+"&"+"&",laCodePrg[lnx])+3), "&",""), Chr(9), ""))
				Endif

				*- se não for uma atualização de propriedade incluida anteriormente, adiciono um nova.
				If Not llHasChangedProperty
					lnItemsFound = lnItemsFound + Iif(Alen(This.Items,1)=1 And Empty(This.Items[1,1]), 0, 1)
					lnItemRow = lnItemsFound
					Dimension This.Items[lnItemsFound,4]			&&- Incremento o array
				Else
					lnItemRow = lnRowFound
				Endif

				This.Items[lnItemRow,1] = lcProperty		&&- Descricao
				This.Items[lnItemRow,2] = lnImage 			&&- No. imagem
				This.Items[lnItemRow,3] = lcToolTip 		&&- Tooltip
				This.Items[lnItemRow,4] = ""
			Endfor
		Endfor


		*- METHODES AND EVENTS -*
		lcSetTalk = Set("talk")
		lcSetNotify = Set("Notify",2)
		lcSetMessage = Set("Message")
		Set Talk Off
		Set Notify Cursor Off
		Set Message To ""

		If Not Used("foxcodeplus2")
			Use foxcodeplus2 Alias foxcodeplus2 In 0
		Endif

		Select foxcodeplus2
		Set Order To typecode

		For lnx = 1 To Alen(laProc,1)

			*- capturo o conteudo da linha
			lcTextLine = This.TreatLine(laCodePrg[laProc[lnx,2]])
			If Empty(lcTextLine)
				Loop
			Endif

			Do Case
					*- considero somente os metodos e eventos da classe que estou posicionado
				Case "." $ laProc[lnx,1] And Lower(Substr(laProc[lnx,1],1,At(".",laProc[lnx,1])-1)) == Lower(plcClassName)

					*- removo o nome da classe do conteudo do array
					lcProc = Alltrim(Substr(laProc[lnx,1], At(".",laProc[lnx,1])+1))

					If "." $ lcProc 	&& nao permito procedure de objetos da classe, somente procedure da classe (ex: procedure textbox.init)
						Loop
					Endif

					Do Case
						Case Lower(Substr(lcTextLine,1,4)) == "prot"					&&- Protected methode or event
							lnImage = Iif(Seek("E"+Lower(Padr(lcProc,30)),"foxcodeplus2"), 15, 6)
							lcProcType = Iif(lnImage=15, "Protected Event", "Protected Method")

						Case Lower(Substr(lcTextLine,1,4)) == "hidd"					&&- Hidden methode or event
							lnImage = Iif(Seek("E"+Lower(Padr(lcProc,30)),"foxcodeplus2"), 14, 5)
							lcProcType = Iif(lnImage=14, "Hidden Event", "Hidden Method")

						Otherwise														&&- public methode or event
							lnImage = Iif(Seek("E"+Lower(Padr(lcProc,30)),"foxcodeplus2"),  3, 4)
							lcProcType = Iif(lnImage=3, "Event", "Method")
					Endcase

					lcToolTip = lcProcType + " " + plcClassName + "." + lcProc + "("+This.GetParameters(@laCodePrg, laProc[lnx,2])+")" +;
						this.GetSummary(@laCodePrg, laProc[lnx,2])

					*- não considero
				Otherwise
					Loop
			Endcase

			*- adiciono ou atualizo o item
			lnRowFound = This.SeekItem(lcProc,lnImage)
			If lnRowFound = 0
				lnItemsFound = lnItemsFound + 1
				Dimension This.Items[lnItemsFound,4]
				lnRowFound = lnItemsFound
			Endif

			This.Items[lnRowFound,1] = lcProc			&&- Descricao
			This.Items[lnRowFound,2] = lnImage 			&&- No. imagem
			This.Items[lnRowFound,3] = lcToolTip 		&&- Tooltip
			This.Items[lnRowFound,4] = ""
		Endfor

		Use In foxcodeplus2

		If lcSetTalk = "ON"
			Set Talk On
		Endif

		If lcSetNotify = "ON"
			Set Notify Cursor On
		Endif

		Set Message To lcSetMessage


		*- se existe a propriedade _MemberData faço o tratamento dos captions
		If Not Empty(lcMemberDataXML)
			This.TreatMemberData(lcMemberDataXML)
		Endif

		*- Adiciono tudo que encontrei ao IntelliSense
		If Not Empty(This.Items[1,1]) And pllAdd
			For lnx = 1 To Alen(This.Items,1)
				This.AddItem(This.Items[lnx,1], This.Items[lnx,2], This.Items[lnx,3])
			Endfor
		Endif

		lnItemsFound = lnItemsFound + Alen(This.Items,1)

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busca em uma VCX especifico as pmes de um classe e as pmes de suas baseclass
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetProcInfoVcx
		Lparameters plcClassName, plcFileVcx
		Local lcAlias, lcText, lcObjects, lcBreakLine, lcSafety, lcAuxClassName, lcAuxFileVcx, lcMemberDataXML, lcXML, lcPath

		Set Console Off

		*- se existe caracters especiais no nome do arquivo ou da classe desconsidero a analise do IntelliSense
		*- nesse caso possivelmente existe programacao para composicao dos nomes
		If "&" $ plcFileVcx Or "(" $ plcFileVcx Or "+" $ plcFileVcx Or ".." $ plcFileVcx 		&&or "'" $ plcFileVcx or "[" $ plcFileVcx or '"' $ plcFileVcx
			Return 0
		Endif

		If "&" $ plcClassName Or "(" $ plcClassName Or "+" $ plcClassName Or "." $ plcClassName Or "'" $ plcClassName Or "[" $ plcClassName Or '"' $ plcClassName
			Return 0
		Endif

		If Not This.GetFilePath(@plcFileVcx)
			This.ShowErrorWriteTime(1, Upper(Justfname(plcFileVcx))) 	&&- file doesn't exist
			Return 0
		Endif

		*-
		lcSetDeleted = Set("Deleted")
		Set Deleted On

		lcAlias = Alias()
		lcBreakLine = Chr(13)
		lcMemberDataXML = ""
		lcAuxClassName = plcClassName
		lcAuxFileVcx = plcFileVcx

		Declare This.Items[1,4]
		This.Items[1,1] = ""


		*- faço o "do while" enquanto existir classes herdadas. Faço isso para obter (herdar) todos as PMEs.
		*- o "do while" termina somente qdo é encontrado uma classe herdada de uma classe padrão do VFP.
		Do While .T.
			lcText = ""
			lcObjects = ""
			lcClassLoc = ""

			If Empty(lcAuxClassName)
				Exit
			Endif

			If Used("___xfcpVcx")
				Use In ___xfcpVcx
			Endif

			*- abro o arquivo vcx para iniciar a varredura pela classe e subclasses
			Try
				Use (lcAuxFileVcx) Alias ___xfcpVcx In 0 Shared Again
				llFileOk = .T.
			Catch To loException
				This.ShowErrorWriteTime(1747, Justfname(Upper(plcFileVcx)) ) 	&&- class file doesn't exist
				llFileOk = .F.
			Endtry

			If Not llFileOk					&&- Arquivo corrompido ou invalido
				Declare This.Items[1,4]
				This.Items[1,1] = ""
				Set Deleted &lcSetDeleted
				If Used(lcAlias)
					Select (lcAlias)
				Endif
				Return 0
			Endif

			*- capturos os objetos (controles) incluidos na classe
			Select ___xfcpVcx
			Scan For Lower(Alltrim(Parent)) == Lower(Alltrim(lcAuxClassName))
				If Alltrim(Lower(BaseClass))="olecontrol"
					lcObjects = lcObjects + "add object " + Alltrim(objname) + " as olecontrol with oleclass = '"+Alltrim(Class)+"'" + lcBreakLine
				Else
					lcObjects = lcObjects + "add object " + Alltrim(objname) + " as " + Alltrim(Class) + lcBreakLine
				Endif
			Endscan

			*- monto o prg baseado na vcx
			Go Top
			If Not "." $ lcAuxClassName
				lcClassName = lcAuxClassName
				Locate For Empty(Parent) And Not Empty(Class) And Not Empty(BaseClass) And Not Empty(objname) And reserved1 == "Class" And Lower(Alltrim(objname)) == Lower(Alltrim(lcClassName))
			Else
				lcClassName = Substr(lcAuxClassName, Rat(".",lcAuxClassName)+1)
				lcParent = Substr(lcAuxClassName, 1, Rat(".",lcAuxClassName)-1)
				Locate For Lower(Alltrim(Parent)) == Lower(Alltrim(lcParent)) And Lower(Alltrim(objname)) == Lower(Alltrim(lcClassName))
			Endif

			If Found()
				lcProperties = ""
				lcMethods = methods

				Declare laPMEsHP[1]
				Declare laPMEs[1]
				Alines(laPMEsHP, Protected)				&&- Todas as PMEs protected ou hidden
				Alines(laPMEs, reserved3)				&&- todos as PMEs que não foram usadas mas que pertencem a classe
				lnTotPropHP = Alen(laPMEsHP,1)

				If Not Empty(laPMEs[1])
					For lnx = 1 To Alen(laPMEs,1)

						*- METHODS -*
						If Substr(laPMEs[lnx],1,1)="*"

							lcThisMethod = Substr(Getwordnum(laPMEs[lnx],1),2)

							*- verifico se é hidden ou protected
							lcHP = ""
							If lnTotPropHP > 0
								For lnz = 1 To lnTotPropHP
									If Substr(laPMEsHP[lnz],1,Len(lcThisMethod)) == lcThisMethod
										lcHP = Iif(Right(laPMEsHP[lnz],1)="^", "HIDDEN ","PROTECTED ")
										Exit
									Endif
								Endfor
							Endif

							*- verifico se tem conteudo para o summary e caso tenha eu crio o sumary para o method
							lcSummary = ""
							If Getwordcount(laPMEs[lnx]) > 1
								lcSummary = "*** <summary>"+Chr(13)+;
									"*** "+Substr(laPMEs[lnx],At(" ",laPMEs[lnx]))+Chr(13)+;
									"*** </summary>"+Chr(13)
							Endif

							*- verifico se o metodo tem codigo de programa e se nao estiver incluo-o no prg
							lcThisMethod = "PROCEDURE " + lcThisMethod + lcBreakLine
							If Not lcThisMethod $ lcMethods
								lcMethods = lcMethods + lcBreakLine + lcSummary + lcHP + lcThisMethod + "ENDPROC" + lcBreakLine
							Else
								*- se for protected ou hidden
								If Not Empty(lcHP)
									lcMethods = Strtran(lcMethods, lcThisMethod, lcHP + lcThisMethod, 1, 1)
									lcThisMethod = lcHP + lcThisMethod
								Endif

								*- se estiver em uso coloco o summary
								If Not Empty(lcSummary)
									lcMethods = Strtran(lcMethods, lcThisMethod, lcSummary + lcThisMethod, 1, 1)
								Endif
							Endif


							*- PROPERTIES -*
						Else
							lcThisProperty = Getwordnum(laPMEs[lnx],1)
							If "[" $ lcThisProperty
								lcThisProperty = Substr(lcThisProperty,1,At("[",lcThisProperty)-1)
							Endif

							llIsArray = .F.
							If Substr(lcThisProperty,1,1) = "^"			&&- se for um array
								lcThisProperty = Substr(lcThisProperty,2)
								llIsArray = .T.
							Endif

							*- verifico se é hidden ou protected
							lcHP = ""
							If lnTotPropHP > 0
								For lnz = 1 To lnTotPropHP
									If Substr(laPMEsHP[lnz],1,Len(lcThisProperty)) == lcThisProperty
										lcHP = Iif(Right(laPMEsHP[lnz],1)="^", "HIDDEN ","PROTECTED ")
										Exit
									Endif
								Endfor
							Endif

							lcSummary = ""

							*- memberdata
							If Lower(lcThisProperty) == "_memberdata"
								lcThisProperty = "_MemberData"
								lcXML = Substr(___xfcpVcx.properties, At("_memberdata =",___xfcpVcx.properties))
								lcXML = Strextract( Strtran(Strtran(lcXML, "<VFPdata>", "<vfpdata>",1,1,1), "</VFPdata>", "</vfpdata>",1,1,1), "<vfpdata>", "</vfpdata>")
								lcMemberDataXML = lcMemberDataXML + lcXML
							Else
								*- verifico se tem conteudo para o summary e caso tenha eu crio o sumary para a propriedade
								If Getwordcount(laPMEs[lnx]) > 1
									lcSummary = Chr(9)+Chr(9)+"&"+"&"+"& "+Substr(laPMEs[lnx],At(" ",laPMEs[lnx])+1)
								Endif
							Endif

							*- verifico se a propriedade esta em uso e se nao estiver incluo no prg
							If llIsArray
								If Empty(lcHP)
									lcThisProperty = "DIMENSION " + lcThisProperty + "[1,0]"
								Else
									lcThisProperty = lcHP + lcThisProperty + "[1,0]"
								Endif
							Else
								If Empty(lcHP)
									lcThisProperty = lcThisProperty + ' = ""'
								Else
									lcThisProperty = lcHP + lcThisProperty
								Endif
							Endif

							*- propriedades encontradas
							lcProperties = lcProperties + lcBreakLine + lcThisProperty + lcSummary
						Endif
					Endfor
				Endif

				lcText = "define class "+Alltrim(objname)+" as "+Alltrim(Class) + lcBreakLine +;
					lcProperties + lcBreakLine +;
					lcObjects + lcBreakLine +;
					"" + lcBreakLine +;
					lcMethods + lcBreakLine +;
					"enddefine"
			Endif

			*- obtenho as pmes da classe atraves do prg gerado (vcx2prg)
			If Not Empty(lcText)
				lcSafety = Set("Safety")
				Set Safety Off
				Strtofile(lcText,This.Tmpfile)
				Set Safety &lcSafety

				This.GetProcInfoPrg(lcClassName, This.Tmpfile, .F., .F.)
			Endif

			*- verifico se a classe é baseada em uma classe nativa do VFP
			Local Array laVFPnativeClass[1]
			laVFPnativeClass[1] = ""
			Select Code From FoxCodePlus Where FoxCodePlus.Type = "T2" And Lower(Alltrim(FoxCodePlus.Code)) == Lower(Alltrim(___xfcpVcx.Class)) Into Array laVFPnativeClass

			Use In FoxCodePlus

			*- se nao for procuro a classe base
			If Empty(laVFPnativeClass[1])
				lcAuxFileVcx = ___xfcpVcx.classloc		&&iif(empty(justpath(___xfcpVcx.classloc)),lcPath,"") + ___xfcpVcx.classloc
				lcAuxClassName = ___xfcpVcx.Class
				Use In 	___xfcpVcx
				Loop
				*- se for a classe base nativa do vfp finalizo (as pmes ja foram capturadas em this.GetProcInfoPrg())
			Else
				Use In 	___xfcpVcx
				Exit
			Endif
		Enddo

		*- se existe a propriedade _MemberData faço o tratamento dos captions
		If Not Empty(lcMemberDataXML)
			lcMemberDataXML = "<VFPData>" + Strtran(Strtran(lcMemberDataXML,"<vfpdata>",""),"</vfpdata>","") + "</VFPData>"
			This.TreatMemberData(lcMemberDataXML)
		Endif

		*- adiciono o que foi encontrado ao IntelliSense
		If Not Empty(This.Items[1,1])
			lnItemsFound = Alen(This.Items,1)
			For lnx = 1 To lnItemsFound
				This.AddItem(This.Items[lnx,1], This.Items[lnx,2], This.Items[lnx,3])
			Endfor
		Else
			lnItemsFound = 0
		Endif

		Set Deleted &lcSetDeleted

		If Used(lcAlias)
			Select (lcAlias)
		Endif

		If lnItemsFound = 0
			This.ShowErrorWriteTime(1576, Iif(Empty(plcClassName), '"UNKNOW"', Upper(plcClassName))) 	&&- class doesn't exist
		Endif

		Return lnItemsFound
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ Busca em todo o texto do programa corrente (prg, methods and events in forms or class) as
	*/ funcoes, procedures, classes, variaveis, #defines e APIs at write-time, etc, etc.
	*/
	*/ plnMode: 0 - all
	*/			1 - classes
	*/			2 - methodes and events
	*/			3 - #defines
	*/			4 - class properties and controls
	*/			5 - variables in current procedure, methode or event
	*/			6 - cursors and tables in current procedure, methode or event in write-time
	*/			7 - DLL functions in current procedure, methode or event
	*/			8 - Procedural functions and classes in a PRG invoked through the "set procedure to" and
	*/              Classes inside in a VCX file invoked through the "set classlib to"
	*/
	*/ plnWho.: 0 - Retorna somente os membros da classe ao qual estou posicionado
	*/			1 - Retorna tudo exceto os membros da classe ao qual estou posicionado
	*/
	*/ pllAdd.: .t. - indica que os itens encontrados serão incluidos no IntelliSense e no array this.items
	*/			.f. - serão incluídos somente no array this.items
	*/
	*/ plnStartLine, plnEndLine:
	*/			.t. - indica que checarei a procedure para buscar os itens do IntelliSense entre as linhas informadas.
	*/		    .f. - checagem do inicio da procedure até a linha corrente.
	*/			(Valido somente para plnMode igual 5, 6, 7 e 8)
	*/------------------------------------------------------------------------------------------------
	Procedure GetProcInfo
		Lparameters plnMode, plnWho, pllAdd, plnStartLine, plnEndLine

		Set Console Off

		pllAdd = Iif(Parameters()<=2, .T., pllAdd)
		plnMode = Evl(plnMode,0)
		plnWho = Evl(plnWho,0)

		Declare This.Items[1,4]				&&- limpo o array.
		This.Items[1,1] = ""

		Declare This.ItemsTables[1,2]		&&- limpo o array.
		This.ItemsTables[1,1] = ""

		Declare This.ItemsObjects[1,3]		&&- limpo o array.
		This.ItemsObjects[1,1] = ""

		Local Array laProc[1,4]
		laProc[1,1] = ""

		Local Array laProcClass[1,4]
		laProcClass[1,1] = ""

		Local Array laCodePrg[1]
		Local lcThisCode, lcSafety, lnLines, lnItemsFound, lnCurrentLine, lcClassName, lcBaseClass, lnClassLine, lnxAux,;
			lnImage, lcParameters, lcToolTip, lcTextLine, lnRowFound, lnTotClasses, lnx, lcProcType, lcProc, lnItemsObjects,;
			lnTotProc, lcInitialValue, lcPropertyType

		lnItemsFound = 0
		lnItemsObjects = 0
		lcClassName = ""
		lcBaseClass = ""
		lnClassLine = 0
		lnImage = 0
		lcParameters = ""
		lcToolTip = ""
		lnTotProc = 0
		lnTotClasses = 0


		*- linha onde estou posicionado no editor
		If This.SetWontop()
			lnCurrentLine = This.GetLineNo()
			If lnCurrentLine = -1
				lnCurrentLine = 0
			Endif
		Else
			Return 0
		Endif


		*- se for classe verifico tambem os metodos e propriedades que inclui na classe no programa em write-time.
		*- Copio todo o codigo e salvo num arquivo temporario
		lcSafety = Set("Safety")
		Set Safety Off
		lcThisCode = This.GetText(0,130000)			&&- capturo o texto da janela atual
		Strtofile(lcThisCode, This.Tmpfile)
		Set Safety &lcSafety

		*- coloco todo o programa em um array
		lnLines = Alines(laCodePrg, lcThisCode)
		lcThisCode = ""
		If lnCurrentLine > lnLines
			lnCurrentLine = lnLines
		Endif

		*- PRGs
		*- obtenho as classes, methodes, procedures and #DEFINE preprocessor directive
		If This.EditorSource <> 10
			Aprocinfo(laProc, This.Tmpfile, 0)

			*- se a linha que cria a procedure conter quebra de linha, a linha informada no array
			*- é a ultima linha da criação da procedure, sendo assim devo achar a 1a. linha
			*- de criacao da procedure e corrigir o array criado pela procinfo() nativa do VFP.
			lnTotProc = Alen(laProc,1)
			If lnTotProc > 0
				For lnx = 1 To lnTotProc

					If Lower(laProc[lnx,3]) = "procedure"
						For lnBackLine = 1 To 99
							If laProc[lnx,2] > 0 And Not This.IsProc(laCodePrg[laProc[lnx,2]])
								laProc[lnx,2] = laProc[lnx,2] - 1
							Else
								Exit
							Endif
						Endfor
					Endif

				Endfor
			Endif

			*- verifico em qual classe estou posicionado
			Aprocinfo(laProcClass, This.Tmpfile, 1)		&&- Classes do arquivo
			lnTotClasses = Alen(laProcClass,1)
			If lnTotClasses	= 1 And Empty(laProcClass[1,1])
				lnTotClasses = 0
			Endif

			If lnTotClasses > 0
				For lnx = 1 To lnTotClasses
					lnxAux = (lnTotClasses+1)-lnx

					If lnCurrentLine > laProcClass[lnxAux,2]
						lcClassName = laProcClass[lnxAux,1]
						lcBaseClass = laProcClass[lnxAux,3]
						lnClassLine = laProcClass[lnxAux,2]
						This.ProcClass = lcClassName
						This.ProcBaseClass = lcBaseClass
						Exit
					Else
						This.ProcClass = ""
						This.ProcBaseClass = ""
					Endif
				Endfor
			Else
				This.ProcBaseClass = ""
				This.ProcClass = ""
			Endif


			*- Form or class editor
		Else
			Local loControl
			Local Array laControl[1]
			laControl[1] = .Null.
			Aselobj(laControl, 1)
			This.ProcBaseClass = ""
			This.ProcClass = ""

			Aprocinfo(laProc, This.Tmpfile, 3)			&&- como estou dentro de um metodo ou evento verifico somente se existem #includes e #defines

			If Vartype(laControl[1]) = "O"
				*- capturo os textos de alguns metodos de um form ou toolbar
				loControl = laControl[1]
				lcClassName = loControl.Name
				lcBaseClass = Iif(Lower(loControl.BaseClass)="olecontrol", loControl.OleClass, loControl.BaseClass)
				lnClassLine = 0
				This.ProcBaseClass = lcBaseClass
				This.ProcClass = lcClassName
			Endif
		Endif


		*- CLASSES -*
		*- considero todas as classes do programa
		If plnWho = 1 And Inlist(plnMode, 0, 1) And lnTotClasses > 0 And This.chkControl = "1"
			For lnx = 1 To lnTotClasses
				lcProc = Getwordnum(laProcClass[lnx,1],1)
				lnImage = 1
				lcToolTip = "Class "+laProcClass[lnx,1]+" as baseclass "+laProcClass[lnx,3]

				This.AddItemProcInfo(lcProc, lnImage, lcToolTip, @lnItemsFound, "")
			Endfor
		Endif


		*- PROPERTIES AND CONTROLS -*
		*- obtenho as propriedades criadas by user da classe que estou posicionado (alem da nativas da classe)
		If plnWho = 0 And Inlist(plnMode, 0, 4)

			*- inicio a varredura por propriedades
			For lnx = lnClassLine+1 To Alen(laCodePrg,1)

				*- desconsidero a linha que estou posicionado
				If lnCurrentLine = lnx
					Loop
				Endif

				*-
				lcTextLine = This.TreatLine(laCodePrg[lnx], @lnx, @laCodePrg)
				lcTextLine = Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(lcTextLine,"="," = "), "(","["), ")","]"), "  "," "), ", ",","), " ,",",")
				lnWordCount = Getwordcount(lcTextLine)
				lcTextWord1 = Getwordnum(Lower(lcTextLine),1)
				lcTextWord2 = Getwordnum(Lower(lcTextLine),2)

				If Empty(lcTextLine)
					Loop
				Endif

				*- fim das propriedades
				If 	lnWordCount  = 1 And Inlist(Substr(lcTextWord1,1,4),"endd","endp","endf") Or;
						lnWordCount  = 2 And Inlist(Substr(lcTextWord1,1,4),"proc","func") Or;
						lnWordCount >= 3 And Inlist(Substr(lcTextWord1,1,4),"hidd","prot") And Inlist(Substr(lcTextWord2,1,4),"proc","func")
					Exit
				Endif

				*- capturo a propriedade
				Do Case
						*- capturo o conteudo do _MemberData
					Case lcTextWord1 == "_memberdata"
						lcMemberDataXML = Substr(lcTextLine, At("=",lcTextLine)+1)
						lcTextLine = "_MemberData"

						*- normal property
					Case (lnWordCount = 3 And Alltrim(lcTextWord2) = "=")
						lcTextLine = Getwordnum(lcTextLine,1)
						lnImage = 7

						*- normal array property
					Case Inlist(Substr(lcTextWord1,1,4),"decl","dime") And Alltrim(lcTextWord2) <> "="
						lcTextLine = Substr(lcTextLine, At(" ",lcTextLine)+1)
						lnImage = 7

						*- hidden property
					Case Substr(lcTextWord1,1,4) == "hidd" And Alltrim(lcTextWord2) <> "="
						lcTextLine = Substr(lcTextLine, At(" ",lcTextLine)+1)
						lnImage = 8

						*- protected property
					Case Substr(lcTextWord1,1,4) == "prot" And Alltrim(lcTextWord2) <> "="
						lcTextLine = Substr(lcTextLine, At(" ",lcTextLine)+1)
						lnImage = 9

						*- controls
					Case Lower(Substr(lcTextLine,1,8)) == "add obje" And " as " $ Lower(lcTextLine)
						Local lcControlName, lcClassControl, lcOleClass
						lcOleClass = ""

						If Lower(Getwordnum(lcTextLine,3)) = "protected"
							lcClassControl = This.ClearQuotes(Getwordnum(lcTextLine,6))
							lcControlName = This.ClearQuotes(Getwordnum(lcTextLine,4))
						Else
							lcClassControl = This.ClearQuotes(Getwordnum(lcTextLine,5))
							lcControlName = This.ClearQuotes(Getwordnum(lcTextLine,3))
						Endif

						*- with oleclass = "COMCTL.ProgCtrl.1"
						If Lower(lcClassControl) == "olecontrol"
							lcOleClass = ""
							If "with oleclass" $ Lower(lcTextLine)
								lcOleClass = Substr(lcTextLine, At("with oleclass",lcTextLine))
								lcOleClass = Alltrim(Substr(lcOleClass, At("=",lcOleClass)+1))
								lcOleClass = Iif(" "$lcOleClass, Substr(lcOleClass,1,At(" ",lcOleClass)-1), lcOleClass)
								lcOleClass = This.ClearQuotes(lcOleClass)
							Endif
							lnImage = 20	&&- Adicionei um controle que é uma activex
						Else
							lnImage = 1		&&- Adicionei um controle com uma classe do VFP
						Endif

						lcTextLine = lcControlName

						lnItemsObjects = lnItemsObjects + 1
						Dimension This.ItemsObjects[lnItemsObjects,3]
						This.ItemsObjects[lnItemsObjects,1] = lcControlName		&&- nome do objeto
						This.ItemsObjects[lnItemsObjects,2] = lcClassControl	&&- nome da classe do objeto
						This.ItemsObjects[lnItemsObjects,3] = lcOleClass		&&- class name ole control


						*- não é uma propriedade valida (podem ser outros comandos)
					Otherwise
						Loop
				Endcase


				*- verifico quantas propriedades foram declaradas na mesma linha
				lnCountPropertiesInLine = Iif("," $ lcTextLine, Occurs(",",lcTextLine)+1, 1)

				*- capturo a propriedade
				For lnz = 1 To lnCountPropertiesInLine
					lcProperty = Getwordnum(lcTextLine,lnz,",")

					If "." $ lcProperty
						Loop
					Endif

					*- se for uma propriedade tipada
					lcPropertyType = ""
					If " as " $ lcProperty
						lcPropertyType = Alltrim(Substr(lcProperty, At(" as ",lcProperty)+4))
						lcProperty = Substr(lcProperty, 1, At(" ",lcProperty)-1)
					Endif

					*- se a propriedade for um array
					llIsAnArray = .F.
					If "[" $ lcProperty
						llIsAnArray = .T.
						lcProperty = Substr(lcProperty, 1, At("[",lcProperty)-1)
					Endif

					If Empty(lcProperty) Or "[" $ lcProperty Or "," $ lcProperty Or "]" $ lcProperty			&&- previno propriedades invalidas
						Loop
					Endif

					*- nao incluo o item se o mesmo ja existir
					llHasChangedProperty = .F.
					lnRowFound = This.SeekItem(lcProperty, lnImage)
					If lnRowFound > 0
						If This.Items[lnRowFound,2] = lnImage
							Loop
						Else
							*- propriedade já incluída no array this.Items porem, agora nessa linha, incluída como Protected ou Hidden
							*- dessa forma mudo a imagem do IntelliSense para Hidden ou Protected somente se for uma propriedade
							If "As Array" $ This.Items[lnRowFound,3] And This.Items[lnRowFound,2] = 7
								This.Items[lnRowFound,2] = lnImage
								llHasChangedProperty = .T.
								llIsAnArray = .T.
							Else
								Loop
							Endif
						Endif
					Endif


					*- tooltip para a propriedade capturada
					Do Case
							*- Classe
						Case lnImage = 1
							lcToolTip = "Control " + lcControlName + " Class " + lcClassControl

							*- normal property
						Case lnImage = 7
							lcToolTip = "Property " + lcProperty + Iif(llIsAnArray," As Array", "") + Iif(!Empty(lcPropertyType), " As " + Proper(lcPropertyType), "")

							*- hidden property
						Case lnImage = 8
							lcToolTip = "Hidden Property " + lcProperty + Iif(llIsAnArray," As Array", "") + Iif(!Empty(lcPropertyType), " As " + Proper(lcPropertyType), "")

							*- protected property
						Case lnImage = 9
							lcToolTip = "Protected Property " + lcProperty + Iif(llIsAnArray," As Array", "") + Iif(!Empty(lcPropertyType), " As " + Proper(lcPropertyType), "")

							*- Ole control
						Case lnImage = 20
							lcToolTip = "Control " + lcControlName + " Class " + lcClassControl + Iif(!Empty(lcOleClass), " OleClass "+lcOleClass,"")
					Endcase

					*- se a propriedade é comentada com "&&&" a frente coloco o comentario no tooltip
					If "&"+"&"+"&" $ laCodePrg[lnx]
						lcToolTip = lcToolTip + Chr(13) + Alltrim(Strtran(Strtran(Substr(laCodePrg[lnx],At("&"+"&"+"&",laCodePrg[lnx])+3), "&",""), Chr(9), ""))
					Endif

					*- se não for uma atualização de propriedade incluida anteriormente, adiciono um nova.
					If Not llHasChangedProperty
						lnItemsFound = lnItemsFound + 1
						Dimension This.Items[lnItemsFound,4]			&&- Incremento o array
						lnRowFound = lnItemsFound
					Endif

					This.Items[lnRowFound,1] = lcProperty		&&- Descricao
					This.Items[lnRowFound,2] = lnImage 			&&- No. imagem
					This.Items[lnRowFound,3] = lcToolTip 		&&- Tooltip
					This.Items[lnRowFound,4] = ""
				Endfor

			Endfor

			*- se nao encontrou objeto incluidos com "add objects" limpo o array
			If lnItemsObjects = 0
				Declare This.ItemsObjects[1,3]
			Endif
		Endif


		*- METHODES AND EVENTS -*
		*- #DEFINES, #INCLUDES AND PROCEDURAL FUNCTIONS AND PROCEDURES -*
		If Inlist(plnMode, 0, 2, 3) And Alen(laProc,1) > 0 And Not Empty(laProc[1,1])

			If Not Used("foxcodeplus2")
				Use foxcodeplus2 Alias foxcodeplus2 In 0
			Endif

			Select foxcodeplus2
			Set Order To typecode

			For lnx = 1 To Alen(laProc,1)
				*- desconsidero a linha que estou posicionado
				If lnCurrentLine = laProc[lnx,2] Or laProc[lnx,2] <= 0
					Loop
				Endif

				*- capturo o conteudo da linha
				lcTextLine = This.TreatLine(laCodePrg[laProc[lnx,2]])
				If Empty(lcTextLine)
					Loop
				Endif

				*-
				Do Case
						*- METHODES AND EVENTS -*
						*- considero somente os metodos e eventos da classe que estou posicionado
					Case plnWho = 0 And Inlist(plnMode, 0, 2) And "." $ laProc[lnx,1] And Lower(Substr(laProc[lnx,1],1,At(".",laProc[lnx,1])-1)) == Lower(lcClassName) And Not Empty(lcClassName)

						*- removo o nome da classe do conteudo do array
						lcProc = Alltrim(Substr(laProc[lnx,1], At(".",laProc[lnx,1])+1))

						If "." $ lcProc 	&& nao permito procedure de objetos da classe, somente procedure da classe (procedure textbox.init)
							Loop
						Endif


						Do Case
							Case Lower(Substr(lcTextLine,1,4)) == "prot"					&&- Protected methode or event
								lnImage = Iif(Seek("E"+Lower(Padr(lcProc,30)),"foxcodeplus2"), 15, 6)
								lcProcType = Iif(lnImage=15, "Protected Event", "Protected Method")

							Case Lower(Substr(lcTextLine,1,4)) == "hidd"					&&- Hidden methode or event
								lnImage = Iif(Seek("E"+Lower(Padr(lcProc,30)),"foxcodeplus2"), 14, 5)
								lcProcType = Iif(lnImage=14, "Hidden Event", "Hidden Method")

							Otherwise														&&- public methode or event
								lnImage = Iif(Seek("E"+Lower(Padr(lcProc,30)),"foxcodeplus2"),  3, 4)
								lcProcType = Iif(lnImage=3, "Event", "Method")
						Endcase

						lcToolTip = lcProcType + " " + Lower(This.ProcClass) + "." + lcProc + "("+This.GetParameters(@laCodePrg, laProc[lnx,2])+")" +;
							this.GetSummary(@laCodePrg, laProc[lnx,2])


						*- #DEFINES, PROCEDURAL FUNCTIONS AND PROCEDURES -*
						*- considero somente as procedures fora da classe e #defines, ou seja, somente procedures e funcoes procedurais.
					Case plnWho = 1 And Inlist(plnMode, 0, 3) And Not "." $ laProc[lnx,1] And Inlist(Lower(laProc[lnx,3]),"procedure","define")

						If This.chkFC <> "1"		&& foxcodeplus.ini
							Loop
						Endif

						lcProc = Alltrim(laProc[lnx,1])

						If Not This.ChkIncremental(This.LastWord,laProc[lnx,1])		&&- IntelliSense incremental
							Loop
						Endif

						Do Case
								*- procedures e funcoes
							Case Lower(laProc[lnx,3]) == "procedure" And plnMode = 0
								lnImage = 19
								lcToolTip = lcProc + "("+This.GetParameters(@laCodePrg, laProc[lnx,2])+")" +;
									this.GetSummary(@laCodePrg, laProc[lnx,2])
								*- #define
							Case Lower(laProc[lnx,3]) == "define"
								lnImage = 12
								lcTextLine = Strtran(lcTextLine,"  ", " ")
								lcToolTip = "Constant " + Substr(lcTextLine, At(" ",lcTextLine)+1)

								*- desconsidera qquer outra coisa
							Otherwise
								Loop
						Endcase


						*- #INCLUDE
					Case plnWho = 1 And Inlist(plnMode, 0, 3) And Lower(laProc[lnx,3]) == "directive" And Lower(Getwordnum(laProc[lnx,1],1)) == "include"

						Local lcFileH
						lcFileH = laCodePrg[laProc[lnx,2]]
						If Inlist( Substr(Getwordnum(lcFileH,2),1,1), "'",'"',"[" )
							lcFileH = Strtran(Strtran(Strtran(lcFileH,"'",'"'), "[", '"'), "]", '"')
							lcFileH = Strextract(lcFileH,'"','"')
						Else
							lcFileH = Getwordnum(lcFileH,2)
						Endif

						lnItemsFound = lnItemsFound + This.GetConstantsFromFile( lcFileH )
						Loop


						*- não considero -*
					Otherwise
						Loop
				Endcase

				*- adiciono ou atualizo o item
				lnRowFound = This.SeekItem(lcProc,lnImage)
				If lnRowFound = 0
					lnItemsFound = lnItemsFound + 1
					Dimension This.Items[lnItemsFound,4]
					lnRowFound = lnItemsFound
				Endif

				This.Items[lnRowFound,1] = lcProc			&&- Descricao
				This.Items[lnRowFound,2] = lnImage 			&&- No. imagem
				This.Items[lnRowFound,3] = lcToolTip 		&&- Tooltip
				This.Items[lnRowFound,4] = ""
			Endfor

			Use In foxcodeplus2
		Endif


		*- VARIABLES / DLLs / CURSORS / TABLES in write-time -*
		*- obtenho as variaveis que foram criadas na procedure que estou posicionado e
		*- considero somente as variaveis que criei até a linha onde estou posicionado
		If plnWho = 1 And Inlist(plnMode,0,5,6,7,8)
			Local lcWord1, lcWord2, lcWord3, lcItem, lnz, lcHasVarType, llProcFound, lnItemsTables, ;
				lnItemsVars, lcProcName, llIsPar, lnAuxAPI, lcAlias

			lnItemsTables = 0
			lnItemsVars = 0

			*- se informei a linha inicial e final
			If Vartype(plnStartLine) = "N" And plnStartLine > 0 And Vartype(plnEndLine) = "N" And plnEndLine > 0 And plnEndLine >= plnStartLine
				lnxAux = plnStartLine
				lnCurrentLine = plnEndLine

				*- posiciono na 1a. linha da procedure a qual estou posicionado e inicio a varredura das variaveis
				*- se não tem procedure... faço a varredura desde a 1a. linha de o todo PRG.
			Else
				llProcFound = .F.
				lnxAux=0
				lnTotProc = Alen(laProc,1)

				If lnTotProc > 0 And Not Empty(laProc[1,1])
					For lnx = 1 To lnTotProc
						lnxAux = (lnTotProc+1)-lnx

						*- acho em qual procedure estou dentro
						If lnCurrentLine > laProc[lnxAux,2] And Lower(laProc[lnxAux,3]) = "procedure"

							*- se estou dentro de uma classe considero somente procedures dessa classe
							*- se "." $ laProc[lnxAux,1] significa que é uma funcao procedural
							If Not Empty(lcClassName) And "." $ laProc[lnxAux,1]
								If Lower(lcClassName) <> Lower( Substr(laProc[lnxAux,1], 1, At(".",laProc[lnxAux,1])-1) )
									Loop
								Endif
							Endif

							lnxAux = laProc[lnxAux,2]
							llProcFound = .T.
							Exit
						Endif

					Endfor
				Endif


				If This.EditorSource <> 10
					*- Verifico do começo do PRG ate a linha corrente, 1a. funcao ou classe encontrada no PRG se existe alguns comandos globais
					*- faço essa verificação somente se estou dentro de uma procedure, metodo ou evento.
					*- SET PROCEDURE TO / SET CLASSLIB TO
					If llProcFound And Inlist(plnMode,0,8)
						For lnx = 1 To lnCurrentLine
							lcTextLine = This.TreatLine(laCodePrg[lnx], @lnx, @laCodePrg)
							lcWord1 = Lower(Getwordnum(lcTextLine, 1))
							lcWord2 = Lower(Getwordnum(lcTextLine, 2))
							lcWord3 = Lower(Getwordnum(lcTextLine, 3))

							If Substr(lcWord1,1,4) == "defi" And Substr(lcWord2,1,4) == "clas" Or This.IsProc(lcTextLine)
								Exit
							Else
								If lcWord1 == "set" And Substr(lcWord2,1,4) == "clas" And lcWord3 == "to"
									This.GetSetClassInfoVcx(Getwordnum(lcTextLine,4))
									Loop
								Endif

								If lcWord1 == "set" And Substr(lcWord2,1,4) == "proc" And lcWord3 == "to"
									This.GetSetProcInfoPrg(lcTextLine)
									Loop
								Endif
							Endif
						Endfor
					Endif

					*- se estou dentro de uma classe e nao estou dentro de uma procedure nao prossigo
					*- pois variaveis, cursores e dlls não podem ser criados fora de um procedure de uma classe.
					If Not Empty(lcClassName) And Not llProcFound
						Return 0
					Else
						lnxAux = Iif(lnxAux=0, 1, lnxAux)
					Endif
				Else
					lnxAux = 1
				Endif
			Endif


			*- inicio a varredura do codigo de programa
			For lnx = lnxAux To lnCurrentLine
				lcTextLine = This.TreatLine(laCodePrg[lnx], @lnx, @laCodePrg)
				lcTextLine = This.TreatWords(lcTextLine)
				lnWordCount = Getwordcount(lcTextLine )
				lcWord1 = Getwordnum(lcTextLine, 1)
				lcWord2 = Getwordnum(lcTextLine, 2)
				lcWord3 = Getwordnum(lcTextLine, 3)
				lcItem = ""
				lcToolTip = ""
				llIsPar = .F.

				*- desconsidero linha vazia ou a linha atual
				If Empty(lcTextLine) Or (lnCurrentLine = lnx And Not This.HasDot)
					Loop
				Endif

				*- asseguro que devo estar dentro de uma procedure ou metodo de um a classe
				*- caso a varredura for desde o inicio do prg, significa que estou fora de uma classe
				*- então devo ignoar tudo entre o "define class ... endclass"
				If Substr(Lower(lcWord1),1,4) == "defi" And Substr(Lower(lcWord2),1,4) == "clas"
					Do While .T.
						lnx = lnx + 1
						If Substr(Lower(laCodePrg[lnx]),1,5) == "endde" Or lnx = lnCurrentLine
							Exit
						Endif
					Enddo
				Endif

				*- desconsidero se for um bloco TEXT... ENDTEXT
				This.SkipTextEndText(lcTextLine, lnx, @laCodePrg)

				*- checagem para inclusao no IntelliSense
				Do Case
						*- variaveis com o tipo de inicialização declarada
					Case ( Inlist(Substr(Lower(lcWord1),1,4), "publ", "priv", "loca", "dime", "decl", "para", "lpar") And ;
							lcWord2 <> "=" And Lower(Getwordnum(lcTextLine,3) ) <> "in" And Lower(Getwordnum(lcTextLine,4) ) <> "in" ) Or ;
							( Substr(Lower(lcWord1),1,4) = "stor" And " to " $ lcTextLine And lcWord2 <> "=" ) Or ;
							( Inlist(Substr(Lower(lcWord1),1,4), "prot", "hidd") And Inlist(Substr(Lower(lcWord2),1,4), "proc", "func") And "(" $ lcTextLine) Or;
							( Inlist(Substr(Lower(lcWord1),1,4), "proc", "func") And lcWord2 <> "=" And "(" $ lcTextLine )

						If Not Inlist(plnMode,0,5)
							Loop
						Endif

						If This.chkVar <> "1"		&& foxcodeplus.ini
							Loop
						Endif

						*- comando nao valido para declaracao de private var
						If Substr(Lower(lcWord1),1,4)="priv" And Lower(lcWord2) = "all"
							Loop
						Endif

						Do Case
								*- variavel declarada com store
							Case Substr(Lower(lcWord1),1,4) = "stor" And " to " $ lcTextLine And lcWord2 <> "="
								lcTextLine = Alltrim(Substr(lcTextLine, At(" to ",lcTextLine)+4))

								*- parametros na mesma linha de criação da funcao
							Case ( Inlist(Substr(Lower(lcWord1),1,4), "prot", "hidd") And Inlist(Substr(Lower(lcWord2),1,4), "proc", "func") ) Or;
									( Inlist(Substr(Lower(lcWord1),1,4), "proc", "func") And lcWord2 <> "=" )

								llIsPar = .T.
								lcTextLine = Alltrim(Substr(lcTextLine, At("(",lcTextLine)+1))
								lcTextLine = Substr(lcTextLine,1,Len(lcTextLine)-1)

								*- qualquer outra declaracao tipada
							Otherwise
								lcTextLine = Strtran(lcTextLine," array "," ",1,1,1)
								lcTextLine = Alltrim(Substr(lcTextLine,At(" ",lcTextLine)))
						Endcase

						*- verifico quantas possiveis variaveis foram declaradas na mesma linha
						lnCountVarsLine = Iif("," $ lcTextLine, Occurs(",",lcTextLine)+1, 1)
						lnImage = 11		&&-Variables

						*- capturo a variavel

						For lnz = 1 To lnCountVarsLine
							lcItem = Alltrim(Getwordnum(lcTextLine,lnz,","))
							lcHasVarType = ""
							llIsAnArray = .F.												&&- Indico se a propriedade é um array ou nao

							If " as " $ Lower(lcItem)
								lcHasVarType = Alltrim(Substr(lcItem, At(" as ", Lower(lcItem))))
								lcHasVarType = " as " + This.ClearQuotes( Proper( Substr(lcHasVarType,At(" ",lcHasVarType)+1) ) )
								lcItem = Alltrim( Substr(lcItem, 1, At(" as ", Lower(lcItem))) )
							Endif

							If "[" $ lcItem Or "(" $ lcItem
								llIsAnArray = .T.
								lcItem = Alltrim( Substr(lcItem, 1, Iif("["$lcItem, At("[",lcItem), At("(",lcItem))-1) )
							Endif

							*- monto o tooltiptext
							lcToolTip = Icase(Substr(Lower(lcWord1),1,4)="publ", "Public ", Substr(Lower(lcWord1),1,4)="priv", "Private ",;
								substr(Lower(lcWord1),1,4)="loca", "Local ",  Substr(Lower(lcWord1),1,4)="lpar", "Local Parameter ",;
								substr(Lower(lcWord1),1,4)="para" Or llIsPar, "Parameter ", "")+;
								iif(llIsAnArray,"Array ","")+"Variable " + lcItem + lcHasVarType

							*- insiro no IntelliSense as variavies encontradas
							This.AddItemProcInfo(lcItem, lnImage, lcToolTip, @lnItemsFound, "")
						Endfor
						Loop


						*- variaveis sendo declaradas em sua valorização
					Case Not "." $ lcWord1 And lcWord2 = "="  And Not Empty(lcWord3) And Inlist(plnMode,0,5) And This.chkVar = "1"
						lcItem  = lcWord1
						lnImage = 11		&&-Variables

						*- se for uma variavel com um objeto instanciado em write-time
						If Getwordnum(lcTextLine,2)="=" And Inlist(Lower(Getwordnum(lcTextLine,3)),"createobject(","newobject(","createobjectex(")
							lcToolTip = "Variable " + Strtran(Strtran(Strtran(Strtran(lcTextLine,"="," as "),"  "," "),"( ","(")," )",")")

							lnItemsVars = lnItemsVars + 1
							Dimension This.ItemsAuxVars[lnItemsVars,4]
							This.ItemsAuxVars[lnItemsVars,1] = lcItem
							This.ItemsAuxVars[lnItemsVars,2] = lnImage
							This.ItemsAuxVars[lnItemsVars,3] = lcToolTip
							This.ItemsAuxVars[lnItemsVars,4] = lcTextLine

							*- qualquer outro tipo de variavel
						Else
							*- se estou valorizando uma variavel com outra variavel
							If 	(Substr(lcWord3,1,1) <> ["] And Right(lcWord3,1) <> ["]) And ;
									(Substr(lcWord3,1,1) <> ['] And Right(lcWord3,1) <> [']) And ;
									(Substr(lcWord3,1,1) <> "[" And Right(lcWord3,1) <> "]") And Not "." $ lcWord3 And Not "&" $ lcWord3

								lnRowFound = Ascan(This.ItemsAuxVars, lcWord3, -1, -1, 1, 15)
								If lnRowFound > 0
									lcTextLine = This.ItemsAuxVars[lnRowFound,4]
									If Getwordnum(lcTextLine,2) = "="
										lcTextLine = lcItem + " " + Substr(lcTextLine,At("=",lcTextLine))
									Endif
								Endif

								If Getwordnum(lcTextLine,2) = "=" And Inlist(Lower(Getwordnum(lcTextLine,3)),"createobject(","newobject(","createobjectex(")
									lcToolTip = "Variable " + Strtran(Strtran(Strtran(Strtran(lcTextLine,"="," as "),"  "," "),"( ","(")," )",")")

									lnItemsVars = lnItemsVars + 1
									Dimension This.ItemsAuxVars[lnItemsVars,4]
									This.ItemsAuxVars[lnItemsVars,1] = lcItem
									This.ItemsAuxVars[lnItemsVars,2] = lnImage
									This.ItemsAuxVars[lnItemsVars,3] = lcToolTip
									This.ItemsAuxVars[lnItemsVars,4] = lcTextLine
								Else
									lcToolTip = "Variable " + lcItem
								Endif
							Else
								lcToolTip = "Variable " + lcItem
							Endif
						Endif


						*- variaveis declaradas com "m."
					Case Substr(Lower(lcWord1),1,2) = "m." And lcWord2 = "="  And Inlist(plnMode,0,5) And This.chkVar = "1"
						lcItem  = Substr(lcWord1,3)
						lnImage = 11		&&-Variables

						*- se for uma variavel com um objeto instanciado em write-time
						If Getwordnum(lcTextLine,2)="=" And Inlist(Lower(Getwordnum(lcTextLine,3)),"createobject(","newobject(","createobjectex(")
							lcToolTip = "Variable " + Strtran(Strtran(Strtran(Strtran(lcTextLine,"="," as "),"  "," "),"( ","(")," )",")")
							*- qualquer outro tipo de variavel
						Else
							lcToolTip = "Variable " + lcItem
						Endif


						*- variaveis criadas por commandos
					Case ( Lower(lcWord1) == "text" Or ;
							lower(lcWord1) == "count" Or ;
							lower(lcWord1) == "calculate" Or ;
							lower(lcWord1) == "average" Or ;
							lower(lcWord1) == "sum" Or ;
							lower(lcWord1) == "catch" Or ;
							lower(lcWord1) == "for" ;
							) And " to " $ Lower(lcTextLine) And Inlist(plnMode,0,5) And This.chkVar = "1"

						If Lower(lcWord1) == "for"
							lcItem = Getwordnum(lcTextLine, 2)
						Else
							lcItem = Alltrim( Substr(lcTextLine, At(" to ",Lower(lcTextLine))+4) )
							lcItem = Getwordnum(lcItem, 1)
						Endif

						lnImage = 11		&&-Variables
						lcToolTip = "Variable "+lcItem


						*- for each .... endfor
					Case Lower(lcWord1) == "for" And Lower(lcWord2) == "each"
						lcItem = lcWord3
						lcTextLine = "<FOREACH>" + Alltrim(Substr(lcTextLine,At(" in ",lcTextLine)+4)) + "</FOREACH>"
						lcToolTip = "Variable " + lcItem
						lnImage = 11


						*- variaveis criadas por "select * from ... into array myVarArray"
						*- ou por "copy to array myVarArray"
					Case ( ( Substr(Lower(lcWord1),1,4) == "sele" And "into array" $ Lower(lcTextLine) ) Or ;
							( Substr(Lower(lcWord1),1,4) == "copy" And "to array" $ Lower(lcTextLine) ) ) And Inlist(plnMode,0,5) And This.chkVar = "1"

						lcItem = Alltrim( Substr(lcTextLine, At("to array",Lower(lcTextLine))+8) )
						lcItem = Chrtran(lcItem,Chr(34)+Chr(39)+Chr(91)+Chr(93)," ")
						lcItem = Getwordnum(lcItem, 1)
						lcItem = This.ClearQuotes(lcItem)
						lnImage = 11						&&-Variables
						lcToolTip = "Array variable "+lcItem


						*- cursors (create cursor)
					Case Substr(Lower(lcWord1),1,4) = "crea" And Inlist(Substr(Lower(lcWord2),1,4),"curs","tabl","dbf") And Inlist(plnMode,0,6) And This.chkTF = "1"
						lcItem = This.ClearQuotes(lcWord3)
						lnImage = 16						&&- Tables / cursors
						lcToolTip = Iif(Substr(Lower(lcWord2),1,4)=="curs", "Cursor ", "Table ") + lcItem

						If Inlist(Substr(Lower(lcWord2),1,4),"curs","tabl","dbf ")
							lcTextLine = Strtran(lcTextLine, " dbf ",   " cursor ", 1, 1, 1)
							lcTextLine = Strtran(lcTextLine, " table ", " cursor ", 1, 1, 1)
							lcTextLine = Strtran(lcTextLine, " tabl ",  " cursor ", 1, 1, 1)
						Endif

						If Not This.ChkValidTableOrFieldName(lcItem)
							Loop
						Endif

						lnItemsTables = lnItemsTables + 1
						Dimension This.ItemsTables[lnItemsTables,2]
						This.ItemsTables[lnItemsTables,1] = lcItem			&&- nome da tabela
						This.ItemsTables[lnItemsTables,2] = lcTextLine		&&- comando para abertura ou criacao da mesma


						*- cursors (select ... into cursor)
					Case Substr(Lower(lcWord1),1,4) == "sele" And  "into cursor" $ Lower(lcTextLine) And Inlist(plnMode,0,6) And This.chkTF = "1"
						lcItem = Alltrim( Substr(lcTextLine, At("into cursor",Lower(lcTextLine))+11) )
						lcItem = Getwordnum(lcItem, 1)
						lcItem = This.ClearQuotes(lcItem)
						lnImage = 16		&&- Tables / cursors
						lcToolTip = "Cursor "+Proper(lcItem)+" created by SELECT - SQL Command"

						If Not This.ChkValidTableOrFieldName(lcItem)
							Loop
						Endif


						*- tables
					Case Lower(lcWord1) = "use" And Lower(lcWord2) <> "in" And Inlist(plnMode,0,6) And This.chkTF = "1"
						lcAlias = ""
						If " alias " $ Lower(lcTextLine)
							lcAlias = Alltrim(Substr(lcTextLine,At(" alias ",Lower(lcTextLine))+7))
							lcAlias = Alltrim(Substr(lcAlias,1,Iif(" "$lcAlias, At(" ",lcAlias)-1, Len(lcAlias))))
							lcAlias = This.ClearQuotes(lcAlias)
							lcItem = lcAlias
							lcItem = This.ClearQuotes(lcItem)
							lcWord2 = Iif(Inlist(Substr(lcWord2,1,1),"(","&"), "(only at run-time)", Proper(lcWord2))
							lcToolTip = "Table "+Proper(lcWord2)+" alias "+Proper(lcAlias)
						Else
							lcItem = lcWord2
							lcItem = This.ClearQuotes(lcItem)
							lcToolTip = "Table "+Proper(lcItem)
						Endif
						lnImage = 16		&&- Tables / cursors

						If Not This.ChkValidTableOrFieldName(lcItem) And Not This.ChkValidTableOrFieldName(lcAlias)
							Loop
						Endif

						lnItemsTables = lnItemsTables + 1
						Dimension This.ItemsTables[lnItemsTables,2]
						This.ItemsTables[lnItemsTables,1] = lcItem			&&- nome da tabela
						This.ItemsTables[lnItemsTables,2] = lcTextLine		&&- comando para abertura ou criacao da mesma


						*- DLLs
					Case Lower(Substr(lcWord1,1,4)) == "decl"  And (Lower(Getwordnum(lcTextLine,3)) == "in" Or Lower(Getwordnum(lcTextLine,4)) == "in") And ;
							inlist(plnMode,0,7) And This.chkAPI = "1"

						lnAuxAPI = 0
						If Lower(Getwordnum(lcTextLine,3)) == "in"
							lnAuxAPI = 1
						Endif

						lcItem = Iif(" as " $ Lower(lcTextLine), Getwordnum(lcTextLine,7-lnAuxAPI), Getwordnum(lcTextLine,3-lnAuxAPI))
						lcItem = This.ClearQuotes(lcItem)

						lcParameters = " "
						If " as " $ Lower(lcTextLine)
							lcParameters = Alltrim(Substr(lcTextLine, At(" as ",Lower(lcTextLine))+4))
							lcParameters = Alltrim(Substr(lcParameters, At(" ", Lower(lcParameters))))
							lcParameters = Iif(Empty(lcParameters), " ", lcParameters)
							lcTextLine = Strtran(lcTextLine, lcParameters, "")
						Else
							If " in " $ Lower(lcTextLine)
								lcParameters = Alltrim(Substr(lcTextLine, At(" in ",Lower(lcTextLine))+4))
								lcParameters = Alltrim(Substr(lcParameters, At(" ", Lower(lcParameters))))
								lcParameters = Iif(Empty(lcParameters), " ", lcParameters)
								lcTextLine = Strtran(lcTextLine, lcParameters, "")
							Endif
						Endif

						lcToolTip = lcItem + "("+lcParameters+")" + Chr(13) + "DLL Function "+Alltrim(Substr(lcTextLine,At(" ",lcTextLine,2-lnAuxAPI)))
						lnImage = 4			&&- DLL


						*- funcoes e classes atraves do SET PROCEDURE TO...
					Case Lower(lcWord1) = "set" And Lower(Substr(lcWord2,1,4)) = "proc" And Lower(lcWord3) = "to" And Inlist(plnMode,0,8)
						This.GetSetProcInfoPrg(lcTextLine)
						Loop


						*- funcoes e classes atraves do SET CLASSLIB TO...
					Case Lower(lcWord1) = "set" And Lower(Substr(lcWord2,1,4)) = "clas" And Lower(lcWord3) = "to" And Inlist(plnMode,0,8)
						This.GetSetClassInfoVcx(Getwordnum(lcTextLine,4))
						Loop


						*- não há item para o IntelliSense nesta linha
					Otherwise
						Loop
				Endcase

				lcItem = This.ClearQuotes(lcItem)
				This.AddItemProcInfo(lcItem, lnImage, lcToolTip, @lnItemsFound, lcTextLine)
			Endfor

			*- se nao encontrou tabela, limpo o array
			If lnItemsTables = 0
				Declare This.ItemsTables[1,2]
			Endif
		Endif


		*- Adiciono tudo que encontrei ao IntelliSense
		lnLines = Iif(Empty(This.Items[1,1]), 0, Alen(This.Items,1))
		If lnLines > 0 And pllAdd
			For lnx = 1 To lnLines
				This.AddItem(This.Items[lnx,1], This.Items[lnx,2], This.Items[lnx,3])
			Endfor
		Else
			lnLines = Iif(!pllAdd, lnLines, 0)
		Endif
		Return lnLines
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Retorna uma string com os parametros da procedure em write-time
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetParameters
		Lparameters plaCodePrg, plnx
		Local lnx, lcParameters, lcFullLine

		Set Console Off

		lcParameters = ""
		lcFullLine = ""

		*- retiro a quebra de linha
		For lnx = plnx To Alen(plaCodePrg,1)
			lcFullLine = lcFullLine + Alltrim(plaCodePrg[lnx])
			lcFullLine = Alltrim(Strtran(Strtran(Strtran(lcFullLine, Chr(9),""), Chr(10),""), Chr(13),""))

			If Right(lcFullLine,1) = ";"
				lcFullLine = Strt(lcFullLine,";","")
			Else
				Exit
			Endif
		Endfor

		*- capturo os parametros se declarados na mesma linha da criacao da procedure
		If "(" $ lcFullLine And ")" $ lcFullLine
			lcParameters = Strextract(lcFullLine,"(",")")

			*- capturo quando declaro com lparameters and parameters
		Else
			lcParameters = ""
			lcFullLine = ""
			plnx = lnx + 1			&&- Proxima linha após linha de criação da função.

			*- a ultima linha do programa é a linha da criação da procedure (procedural)
			If plnx >= Alen(plaCodePrg,1)
				lcParameters = ""

				*- existe linha após a criação da procedure, então verifico se é a criação de parametros
			Else
				If Getwordcount(plaCodePrg[plnx]) >= 2 And Inlist(Lower(Substr(Getwordnum(plaCodePrg[plnx],1),1,4)),"lpar","para")
					For lnx = plnx To Alen(plaCodePrg,1)
						lcFullLine = lcFullLine + Alltrim(plaCodePrg[lnx])
						lcFullLine = Alltrim(Strtran(Strtran(Strtran(lcFullLine, Chr(9),""), Chr(10),""), Chr(13),""))

						If Right(lcFullLine,1) = ";"
							lcFullLine = Strt(lcFullLine,";","")
						Else
							Exit
						Endif
					Endfor

					lcParameters = Alltrim(Substr(lcFullLine, At(" ",lcFullLine)))
				Endif
			Endif
		Endif

		lcParameters = Strtran(Strtran(Strtran(lcParameters,"  ", ""), " , ", ", "), " ,", ", ")

		Return lcParameters
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Retiro as aspas de um texto
	*/------------------------------------------------------------------------------------------------
	Protected Procedure ClearQuotes
		Lparameters plcText
		Set Console Off
		Return Chrtran(plcText,Chr(34)+Chr(39)+Chr(91)+Chr(93),"")
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Adiciono um item no array this.Items fazendo antes todas as verificações necessarias
	*/------------------------------------------------------------------------------------------------
	Protected Procedure AddItemProcInfo
		Lparameters plcItem, plnImage, plcToolTip, plnItemsFound, plcTextLine
		plcItem = Alltrim(plcItem)
		Local lnRowFound

		Set Console Off

		*- IntelliSense incremental
		If Not This.ChkIncremental(This.LastWord,plcItem)
			Return .F.
		Endif

		*- previno nome de itens invalidos
		If	Empty(plcItem) Or ("[" $ plcItem) Or ("(" $ plcItem) Or ("," $ plcItem) Or (")" $ plcItem) Or ("]" $ plcItem) Or;
				("+" $ plcItem) Or ("-" $ plcItem) Or ("," $ plcItem) Or (";" $ plcItem) Or ("." $ plcItem) Or ("&" $ plcItem) Or ("@" $ plcItem)
			Return .F.
		Endif

		*- nao incluo o item se o mesmo ja existir...
		lnRowFound = This.SeekItem(plcItem, plnImage)
		If lnRowFound > 0
			If This.Items[lnRowFound,2] = plnImage
				*- se for uma variavel eu atualizo o tooltip e a linha de comando com o seu ultimo uso.
				If plnImage = 11 And Getwordnum(plcTextLine,2) = "="
					This.Items[lnRowFound,3] = plcToolTip
					This.Items[lnRowFound,4] = plcTextLine
				Endif
				Return .F.
			Endif
		Endif

		*- adiciono no array para inserir na lista
		plnItemsFound = plnItemsFound + 1
		Dimension This.Items[plnItemsFound,4]			&&- Incremento o array
		This.Items[plnItemsFound,1] = plcItem			&&- Descricao
		This.Items[plnItemsFound,2] = plnImage 			&&- No. imagem
		This.Items[plnItemsFound,3] = plcToolTip		&&- Tooltip
		This.Items[plnItemsFound,4] = plcTextLine		&&- conteudo da linha (quando passado... usado somente em alguns casos)

		Return .T.
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Get API functions in run-time
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetAPIs
		Lparameters plcWord
		Local Array laDLLs[1]
		laDLLs[1] = ""
		Local lnLines, lnItemsFound, lnRows, lnItems, lnx, lcToolTip

		Set Console Off

		Adlls(laDLLs)
		lnLines = Alen(laDLLs,1)
		If Empty(laDLLs[1])
			Return 0
		Endif

		*- conto quantas funcoes da API estao declaras conforme oque estou digitando
		If Not Empty(plcWord)
			lnItemsFound = 0
			For lnx = 1 To lnLines
				If Lower("xxfcpWinAPI_") $ Lower(Substr(laDLLs[lnx,2],1,12))		&&- Ignoro DLLs functions utilizadas pelo FoxcodePlus
					Loop
				Endif

				If This.ChkIncremental(plcWord,laDLLs[lnx,2])
					lnItemsFound = lnItemsFound + 1
				Endif
			Endfor
		Else
			lnItemsFound = lnLines
		Endif

		*- insiro na lista
		If lnItemsFound > 0
			For lnx = 1 To lnLines
				If This.ChkIncremental(plcWord,laDLLs[lnx,2])
					lcToolTip = laDLLs[lnx,2]+"( )"+Chr(13)+"DLL Function "+laDLLs[lnx,1] + Iif(laDLLs[lnx,1]<>laDLLs[lnx,2]," as "+laDLLs[lnx,2],"") + " (running)" + Chr(13) + laDLLs[lnx,3]
					This.AddItem(laDLLs[lnx,2], 4, lcToolTip)
				Endif
			Endfor
		Endif

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ busco as funcoes e comandos do vfp
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetFCs
		Lparameters plcWord
		Local lnLines, lnImage, lnx, lcTypes, lcWord1, lcWord2
		Local Array laItems[1]
		laItems[1] = ""

		Set Console Off

		lcWord1 = Lower(Getwordnum(This.TextLine,1))
		lcWord2 = Lower(Getwordnum(This.TextLine,2))
		lcTypes = Iif(This.WordCount<=1, "C //C1//F //O //O2//C3", "F //O //O2//C3")

		If This.chkFC <> "1"		&& foxcodeplus.ini
			Return 0
		Endif

		Do Case
				*- apresento as funcoes e commandos no IntelliSense conforme abaixo, somente se o comando for "Select sql"
			Case Substr(lcWord1,1,4) == "sele" And Getwordcount(This.TextLine)>=2
				Select * From FoxCodePlus ;
					where (Type = "F " And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "str ","substr ","substrc ","at ","at_c ","atc ","atcc ","int ","ceiling ","round ", "val ", "transform ")) Or ;
					(Type = "F " And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "isnull ","empty ","upper ","proper ","lower ","right ","rightc ","rat ","ratc ","alltrim ", "ltrim ")) Or ;
					(Type = "F " And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "rtrim ","min ","strtran ","strextract ","chrtran ","chrtranc ","iif ","icase ","space ","padc ", "padl ", "padr ")) Or ;
					(Type = "F " And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "avg ","cnt ","count ","max ","min ","npv ","std ","sum ","var ", "cast ")) Or ;
					(Type = "C3" And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "and ","or ")) Or ;
					(Type = "C4") And ;
					this.ChkIncremental(plcWord,FoxCodePlus.Code) ;
					into Array laItems


				*- create cursor, create table, create dbf
			Case Substr(lcWord1,1,4) == "crea" And Inlist(Substr(lcWord2,1,4),"tabl","curs","dbf ") And Getwordcount(This.TextLine)>=4
				Select * From FoxCodePlus ;
					where Inlist(Type,"C6") And ;
					this.ChkIncremental(plcWord,FoxCodePlus.Code) ;
					into Array laItems


				*- replace
			Case Substr(lcWord1,1,5) == "repla" And Getwordcount(This.TextLine)>=2
				Select * From FoxCodePlus ;
					where (Type = "F " And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "str ","substr ","substrc ","at ","at_c ","atc ","atcc ","int ","ceiling ","round ", "val ", "transform ")) Or ;
					(Type = "F " And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "isnull ","empty ","upper ","proper ","lower ","right ","rightc ","rat ","ratc ","alltrim ", "ltrim ")) Or ;
					(Type = "F " And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "rtrim ","min ","strtran ","strextract ","chrtran ","chrtranc ","iif ","icase ","space ","padc ", "padl ", "padr ")) Or ;
					(Type = "F " And Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ", "max ","min ", "cast ")) Or ;
					(Type = "C7") And ;
					this.ChkIncremental(plcWord,FoxCodePlus.Code) ;
					into Array laItems


				*- apresento os commandos no IntelliSense conforme abaixo, somente se o comando for "alter table"
			Case Substr(lcWord1,1,4) == "alte" And Substr(lcWord2,1,4) == "tabl" And Getwordcount(This.TextLine)>=4
				Select * From FoxCodePlus ;
					where Inlist(Type,"C5","C6") And ;
					this.ChkIncremental(plcWord,FoxCodePlus.Code) ;
					into Array laItems


				*- apresento as funcoes no IntelliSense conforme abaixo, somente se o comando for "calculate"
			Case Substr(lcWord1,1,4) == "calc" And Getwordcount(This.TextLine)>=2
				Select * From FoxCodePlus ;
					where Type = "F " And ;
					inlist(Lower(Alltrim(FoxCodePlus.Code))+" ","avg ","cnt ","count ","max ","min ","npv ","std ","sum ","var ") And ;
					this.ChkIncremental(plcWord,FoxCodePlus.Code) ;
					into Array laItems


				*- caso contrario apresento todas as funcoes e comandos
			Otherwise
				Select * From FoxCodePlus ;
					where Type $ lcTypes And ;
					not Inlist(Lower(Alltrim(FoxCodePlus.Code))+" ","avg ","cnt ","count ","npv ","std ","sum ","var ") And ;
					this.ChkIncremental(plcWord,FoxCodePlus.Code) ;
					into Array laItems
		Endcase


		Use In FoxCodePlus
		If Not Empty(laItems[1])
			lnLines = Alen(laItems,1)
		Else
			lnLines = 0
		Endif


		*- insiro na lista
		If lnLines > 0
			For lnx = 1 To lnLines
				Do Case
					Case Inlist(laItems[lnx,1],"C ","C3","C4","C5","C7")
						lnImage = 2
					Case laItems[lnx,1] = "C1"
						lnImage = 18
					Case laItems[lnx,1] = "F "
						lnImage = 19
					Case laItems[lnx,1] = "O "
						lnImage = 1
					Case laItems[lnx,1] = "O2"
						lnImage = 20
					Case laItems[lnx,1] = "C6"
						lnImage = 12
					Otherwise
						lnImage = 0
				Endcase

				This.AddItem(laItems[lnx,2], lnImage, laItems[lnx,3])
			Endfor
		Endif

		Return lnLines
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ busco todos os controles visuais inseridos em um form ou em uma classe visual e
	*/ todos os objetos contidos no controle são inseridos na lista conforme plcWord
	*/ Se plcWord for empty todos os controles são inseridos na lista.
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetControls
		Lparameters plcWord, ploControl
		Local lnLines, lnItemsFound, lnRows, lnItems, laControl, loControl, laObjects, lcControlName, lcClassControl, lcOleClass, lcCaption

		Set Console Off

		If This.chkControl <> "1"
			Return 0
		Endif

		If This.EditorSource <> 10		&&- Se não for dentro do form designer ou class designer
			Return 0
		Endif

		Local Array laControl[1]
		laControl[1] = .Null.

		Local Array laObjects[1]
		laObjects[1] = ""

		*- capturo o controle que estou em foco no form ou classe em write-time
		If Vartype(ploControl) <> "O"
			Aselobj(laControl, 1 )
			If Vartype(laControl[1]) <> "O"
				Return 0
			Endif
		Endif

		*- capturo os objetos incluidos dentro do controle
		loControl = Iif(Vartype(ploControl)<>"O", laControl[1], ploControl)
		Amembers(laObjects, loControl, 2)
		If Not Empty(laObjects[1])
			lnLines	= Alen(laObjects,1)
		Else
			Return 0
		Endif

		*- conto quantos objetos foram encontrados
		If Not Empty(plcWord)
			lnItemsFound = 0
			For lnx = 1 To lnLines
				If This.ChkIncremental(plcWord,laObjects[lnx])
					lnItemsFound = lnItemsFound + 1
				Endif
			Endfor
		Else
			lnItemsFound = lnLines
		Endif

		*- insiro na lista os nomes dos objetos
		If lnItemsFound > 0
			For lnx = 1 To lnLines
				If Not Empty(plcWord)
					If Not This.ChkIncremental(plcWord,laObjects[lnx])
						Loop
					Endif
				Endif

				Try
					lcControlName = Evaluate("loControl."+laObjects[lnx]+".name")
					lcClassControl = Evaluate("loControl."+laObjects[lnx]+".class")
					lcOleClass = Iif(Type('evaluate("loControl."+laObjects[lnx]+".oleclass")') = "C", Evaluate("loControl."+laObjects[lnx]+".oleclass"), "")
					lcToolTip = "Control "+lcControlName+" Class "+lcClassControl + Iif(!Empty(lcOleClass), " OleClass "+lcOleClass,"")

					If Type('evaluate("loControl."+laObjects[lnx]+".caption")') = "C"
						lcCaption = Alltrim(Evaluate("loControl."+laObjects[lnx]+".caption"))
						lcToolTip = lcToolTip + Chr(13) + "Caption: "+Iif(Len(lcCaption)>40, Substr(lcCaption,1,40)+"...", lcCaption)
					Endif
				Catch
					lcControlName = laObjects[lnx]
					lcToolTip = ""
				Endtry

				This.AddItem(lcControlName, 13, lcToolTip)

			Endfor
		Endif

		Return lnItemsFound
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ busco os objetos que estão em memória.
	*/ todos os objetos em memória são inseridos na lista conforme plcWord
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetObjectsRunTime
		Lparameters plcWord
		Local lnLines, lcTmpFile, lnLines, lnItemsFound, lnx, lnRows, lcSafety, lcVar, lcToolTip

		Set Console Off

		Local Array laObjects[1], laTmpFile[1]
		laObjects[1] = ""
		laTmpFile[1] = ""
		lnLines = 0
		lnItemsFound = 0

		If This.chkObj <> "1"			&& foxcodeplus.ini
			Return 0
		Endif

		If Empty(plcWord)
			Return 0
		Endif

		*- capturo os objetos em memoria
		lcSafety = Set("Safety")
		Set Safety Off
		List Memory To File (This.Tmpfile) Noconsole
		Set Safety &lcSafety

		lcTmpFile = Filetostr(This.Tmpfile)
		Alines(laTmpFile, lcTmpFile)

		For lnx = 1 To Alen(laTmpFile,1)
			If "variables defined," $ laTmpFile[lnx]
				Exit
			Endif

			If Empty(laTmpFile[lnx]) Or Inlist(laTmpFile[lnx],"(",")") Or Empty(Substr(laTmpFile[lnx],1,1))
				Loop
			Endif

			lcVar = laTmpFile[lnx]
			If Empty(Substr(laTmpFile[lnx+1],1,1))
				lcVar = lcVar + laTmpFile[lnx+1]
			Endif

			If "  O  " $ lcVar
				lnLines = lnLines + 1
				Dimension laObjects[lnLines]
				laObjects[lnLines] = Substr(lcVar, 1, At(" ",lcVar)-1)
			Endif
		Endfor

		*- insiro na lista
		If lnLines > 0
			For lnx = 1 To lnLines
				If Empty(laObjects[lnx]) Or Not This.ChkIncremental(plcWord,laObjects[lnx])
					Loop
				Endif

				lnItemsFound = lnItemsFound + 1

				Try
					lcToolTip = "Object " + Lower(laObjects[lnx]) + " Class " + &laObjects[lnx]..Class + Chr(13) +;
						"BaseClass: " + &laObjects[lnx]..BaseClass + Chr(13) +;
						"ClassLibrary: " + Iif(Empty(&laObjects[lnx]..ClassLibrary), "(None)", &laObjects[lnx]..ClassLibrary)
				Catch
					lcToolTip = "Object " + Lower(laObjects[lnx])
				Endtry

				lcToolTip = lcToolTip + Chr(13) + Chr(13) + "This is an object instantiated in memory (at run-time)"

				This.AddItem(Proper(laObjects[lnx]), 20, lcToolTip)
			Endfor
		Endif

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Captura todas as constantes incluidas no arquivo plcFile e adiciona ao array this.Items
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetConstantsFromFile
		Lparameters plcFile
		Local lcThisCode, lnx, lcConst, lnImage, lnRowFound, lnItemsFound, lnTotItems, lnRowFound, lcTextLine, lcToolTip, lcConst
		Local Array laProc[1,3], laCodePrg[1]

		Set Console Off

		If This.chkFC <> "1" Or Not This.GetFilePath(@plcFile)		&& foxcodeplus.ini
			Return 0
		Endif

		lnItemsFound = 0

		If Aprocinfo(laProc, plcFile, 3) > 0

			Alines(laCodePrg, Filetostr(plcFile))
			lnTotItems = Iif(Alen(This.Items,1)=1 And Empty(This.Items[1,1]), 0, Alen(This.Items,1))
			lnItemsFound = 0
			lnImage = 12

			For lnx = 1 To Alen(laProc,1)
				lcConst = Alltrim(laProc[lnx,1])

				*if this.IncrementalResult
				If Not This.ChkIncremental(This.LastWord,lcConst) Or Empty(lcConst)		&&- IntelliSense incremental
					Loop
				Endif
				*endif

				*- adiciono ou atualizo o item
				lnRowFound = This.SeekItem(lcConst, lnImage)
				If lnRowFound = 0
					lnItemsFound = lnItemsFound + 1
					Dimension This.Items[lnTotItems + lnItemsFound, 4]
					lnRowFound = lnTotItems + lnItemsFound
				Endif

				lcTextLine = This.TreatLine(laCodePrg[laProc[lnx,2]])
				lcTextLine = Strtran(lcTextLine,"  ", " ")
				lcToolTip = "Constant " + Substr(lcTextLine, At(" ",lcTextLine)+1) + Chr(13) + "File " + Upper(plcFile)

				This.Items[lnRowFound,1] = lcConst			&&- Descricao
				This.Items[lnRowFound,2] = lnImage 			&&- No. imagem
				This.Items[lnRowFound,3] = lcToolTip 		&&- Tooltip
				This.Items[lnRowFound,4] = ""
			Endfor

		Endif

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Ao pressionar DOT "." abre um novo IntelliSense para objetos,
	*/ tabelas/cursores, campos de uma tabela ou lista de variaveis
	*/ NOTE: Uma parte do codigo do metodo THIS.MAIN() é executado qdo pressiona "."
	*/------------------------------------------------------------------------------------------------
	Procedure GetDot
		Local lnLines, lcText, lcLastWord, lcLastWord2, lnWordCount, lcObjReference, lcObjName, loTempObj, ;
			lnRowFound, llCheckFoxCodePlus, lnCountTables, lcOnError, lnCountVars, lcFileDbf, lcAlias, ;
			llReturn, lcSetTalk, lcSetNotify, lcSetMessage, lcSetExact, lcSavecliptext, lcTextSelected, ;
			lcFullTextControl, loControl, lcControlName, lcParentControl, loParentControl, lcAuxControl, ;
			llSqlFields, lcTable

		Local Array laItemsVars[1,4], laItemsTables[1,2], laItemsTablesEnvironment[1,2], laControl[1]

		Set Console Off

		lcSetTalk = Set("talk")
		lcSetNotify = Set("Notify",2)
		lcSetMessage = Set("Message")
		lcSetExact = Set("exact")

		Set Talk Off
		Set Notify Cursor Off
		Set Exact On
		Set Message To ""

		llSqlFields = .F.
		llReturn = .T.
		laControl[1] = .Null.

		If Isnull(This.IntelliSense)
			This.IntelliSense = Newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")
		Endif


		*- verifico em qual lugar do vfp estou acionando o IntelliSense
		If Not This.SetWontop()

			*- Nao estou com o foco no editor, ou seja, estou com foco em algum form do VFP ou em algum form de uma aplicao em run-time
			*- verifico se é um objeto do vfp
			Local loCurrentObject, lcActiveClass, lnSelStart, llReadOnly, lcValue, lcDecimalValue, lcTypedBeforeDot, lnx, lnSetDecimal
			Try
				loCurrentObject = _Screen.ActiveForm.ActiveControl
				lcActiveClass = loCurrentObject.BaseClass
				lnSelStart    = loCurrentObject.SelStart
				llReadOnly    = loCurrentObject.ReadOnly
			Catch
				loCurrentObject = .Null.
				lcActiveClass = ""
				lnSelStart = 0
				llReadOnly = .F.
			Endtry

			*- se for um objeto do vfp faço tratamento do "."
			If Lower(lcActiveClass) $ "//textbox//editbox//combobox//spinner" And lnSelStart > 0 And Not llReadOnly
				Do Case
					Case Lower(lcActiveClass) = "combobox"
						lcValue = loCurrentObject.Text
						loCurrentObject.DisplayValue = Substr(lcValue, 1, lnSelStart) + "." + Substr(lcValue, lnSelStart+1)
						Keyboard '{RightArrow}{home}'+Replicate('{RightArrow}',lnSelStart+1) Plain

					Case Lower(lcActiveClass) $ "textbox//editbox//spinner"
						lcValue = loCurrentObject.Value
						*- string textbox
						If Vartype(lcValue) = "C"
							loCurrentObject.Value = Substr(lcValue, 1, lnSelStart) + "." + Substr(lcValue, lnSelStart+1)
							Keyboard '{RightArrow}{home}'+Replicate('{RightArrow}',lnSelStart+1) Plain
						Else
							*- numeric textbox or spinner
							*- caso o campo tenho separacao decimal utilizando o "." faço o tratamento abaixo
							If Vartype(lcValue) = "N"
								*- verifico se existe valor nas casas decimais
								lcDecimalValue = Iif("." $ loCurrentObject.Text, Substr(loCurrentObject.Text, Rat(".",loCurrentObject.Text)+1), "")
								lcDecimalValue = Iif(Empty(lcDecimalValue), Strtran(lcDecimalValue," ","0"), lcDecimalValue)

								*- capturo o valor inteiro
								lcValue = ""
								lcTypedBeforeDot = Alltrim(Substr(loCurrentObject.Text,1,loCurrentObject.SelStart))

								For lnx = 1 To Len(lcTypedBeforeDot)
									If Substr(lcTypedBeforeDot,lnx,1) = "."
										Exit
									Else
										lcValue = lcValue + Iif(Isdigit(Substr(lcTypedBeforeDot,lnx,1)), Substr(lcTypedBeforeDot,lnx,1), "")
									Endif
								Endfor

								lnSetDecimal = Set("Decimals")
								Set Decimals To Len(lcDecimalValue)
								lcValue = lcValue + "." + lcDecimalValue
								loCurrentObject.Value = Val(lcValue)
								Set Decimals To lnSetDecimal

								Keyboard '{end}'

								If Not Empty(lcDecimalValue) And Len(lcDecimalValue)>=2
									Keyboard Replicate('{leftarrow}',Len(lcDecimalValue)-1)
								Endif
							Endif
						Endif

						*- nao previsto
					Otherwise

				Endcase
			ELSE
				_EdInsert(This.EditorHwnd, ".", 1)
				*Keyboard "." Plain
			Endif

			llReturn = .F.
		Endif


		If llReturn
			*- se tiver texto selecionado eu deleto mas antes checo se nao foi uma auto-selecao da IDE do VFP.
			*- a auto-selecao acontece qdo fechamos parentes ou colchetes ")" ou "]" para indicar a funcao ou array que o abriu.
			lcSavecliptext = _Cliptext

			_EdCopy(This.EditorHwnd)
			lcTextSelected = _Cliptext
			lcTextSelected = Iif(Substr(lcTextSelected,1,1)="[", Stuff(lcTextSelected,1,1,"("), lcTextSelected)
			lcTextSelected = Iif(Right(lcTextSelected,1)="]", Stuff(lcTextSelected,Len(lcTextSelected),1,")"), lcTextSelected)

			If	( Not Empty(lcTextSelected) And Substr(lcTextSelected,1,1) = "(" And Right(lcTextSelected,1) = ")" )
				lcTextSelected = lcTextSelected + "."
				_EdPaste(This.EditorHwnd)
			Else
				_EdDelete(This.EditorHwnd)
			Endif

			_Cliptext = lcSavecliptext


			*- se estou dentro de aspas nao abro o intellisense
			lcText = This.TreatLine(This.GetTextLine())
			If This.IsInQuotes(lcText)
				Keyboard "." Plain
				llReturn = .F.
			Else
				*- se estou dentro de um Text...EndText e for uma instrução SQL (MS Sql Server or another database)
				*- abro o IntelliSense com os campos do SQL
				If This.GetTextEndText(lcText)
					If This.IsSqlIntelliSense And Not Inlist(Lower(Getwordnum(lcText,Getwordcount(lcText)-1)), "from", "join", "into", "update") And Not Inlist(Lower(Getwordnum(lcText,Getwordcount(lcText)-2)), "from", "join", "into", "update") And Not Empty(Right(This.GetTextLine(),1))
						*- selecionei pressionando "." com o intellisense aberto
						If This.IntelliSense.Showed
							This.SelectItem(0)
							lcText = Strtran(Strtran(This.TreatWords(This.TreatLine(This.GetTextLine())),"[ ","[")," ]","]")
							lcLastWord2 = Getwordnum(lcText, Getwordcount(lcText))

							*- pressionei "." sem o IntelliSense estar aberto
						Else
							lcText = Strtran(Strtran(This.TreatWords(lcText),"[ ","[")," ]","]")
							lcLastWord2 = Getwordnum(lcText, Getwordcount(lcText))
							If Inlist(Right(lcLastWord2,1), '"', ']')
								lcLastWord2 =  Getwordnum(lcText, Getwordcount(lcText)-1) + " " + lcLastWord2
							Endif
							This.LastWord = lcLastWord2
						Endif

						*-
						lcSqlTable = lcLastWord2
						This.IntelliSense = Newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")

						*- se a tabela for um alias eu capturo o nome real da tabela
						If This.GetSqlTablesInCmd(lcSqlTable, 2, .F., .T.) > 0
							lnRowFound = Ascan(This.ItemsTables, lcSqlTable, -1,-1, 1, 15)
							If lnRowFound > 0 And Getwordnum(This.ItemsTables[lnRowFound,2],1) == "Table" And " As " $ This.ItemsTables[lnRowFound,2]
								lcSqlTable = Strextract(This.ItemsTables[lnRowFound,2],"Table "," As ")
							Endif
						Endif

						*- capturo os campos da tabela SQL
						This.IncrementalResult = .F.
						lnLines = This.GetSqlFields(lcSqlTable)
						This.IncrementalResult = .T.

						llSqlFields = Iif(lnLines>0, .T., .F.)
					Endif

					_EdInsert(This.EditorHwnd, ".", 1)
					llReturn = .F.
				Endif
			Endif
		Endif


		*- se o IntelliSense do foxcodeplus estiver aberto e pressiono "." seleciono o item da lista.
		*- valido somente se for membros de um objeto ou campos de uma tabela.
		If llReturn
			If This.IntelliSense.Showed
				If Inlist(This.IntelliSense.ActiveImage, 1,3,4,5,6,7,8,9,10,11,13,14,15,16,20)
					This.SelectItem(0)
				Endif
			Endif

			*- busco o ultima palavra digitada antes do ponto
			lcText = This.TreatLine(This.GetTextLine()+".")
			If This.IsComment Or Right(lcText,2) = ".."
				_EdInsert(This.EditorHwnd, ".", 1)
				llReturn = .F.
			Endif
		Endif


		If llReturn
			This.IntelliSense = Newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")

			lcText = This.TreatWords(lcText)
			lnWordCount = Getwordcount(lcText)
			lcLastWord = Getwordnum(lcText, lnWordCount)
			lcLastWord2 = Substr(lcLastWord,1,Len(lcLastWord)-1)

			*- substituicao do IntelliSense nativo da IDE do VFP em form e class designer
			lcFullTextControl = ""
			If This.EditorSource = 10 And This.chkMngDesignTime = "1"
				lcFullTextControl = Iif(" " $ lcText, Substr(lcText,Rat(" ",lcText)+1), lcText)
				lcFullTextControl = Iif(Lower(Substr(lcFullTextControl,1,5)) == "this." Or Lower(Substr(lcFullTextControl,1,9)) == "thisform." Or Lower(Substr(lcFullTextControl,1,12)) == "thisformset.",  lcFullTextControl, "")
				lcFullTextControl = Iif(Right(lcFullTextControl,1) = ".", Substr(lcFullTextControl,1,Len(lcFullTextControl)-1), lcFullTextControl)
			Endif

			*- controlo as propriedades do foxcodeplus
			This.LastWord = lcLastWord2
			This.IncrementalResult = .F.

			*- crio a lista de tabelas (write-time prg)
			This.GetProcInfo(6,1,.F.)									&&- aqui tambem valorizo this.ProcBaseClass
			lnCountTables = Alen(This.ItemsTables,1)
			If lnCountTables > 0
				Dimension laItemsTables[lnCountTables,2]
				Acopy(This.ItemsTables, laItemsTables)
			Endif

			*- crio a lista de tabelas do dataenvironment
			lnCountTablesDataEnvironment = This.GetTablesDataEnvironment(lcLastWord2,.F.)
			If lnCountTablesDataEnvironment > 0
				Dimension laItemsTables[lnCountTablesDataEnvironment + lnCountTables,2]
				For lnx = 1 To lnCountTablesDataEnvironment
					laItemsTables[lnCountTables+lnx,1] = This.ItemsTables[lnx,1]
					laItemsTables[lnCountTables+lnx,2] = This.ItemsTables[lnx,2]
				Endfor
			Endif
			lnCountTables = lnCountTables + lnCountTablesDataEnvironment


			*- crio a lista de variaveis
			lnCountVars = This.GetProcInfo(5,1,.F.)
			Dimension laItemsVars[alen(this.Items,1),4]
			Acopy(This.Items, laItemsVars)								&&- copio as variaveis encontradas para outro array

			*- executar alguns metodos acima retorno o valor da propriedade
			This.IncrementalResult = .T.

			*- Indico ao IntelliSense que "." esta sendo executado
			This.HasDot = .T.

			*- se não estou dentro de uma classe e invoco "this." não prossigo e apresento o erro
			If ( Lower(Substr(lcLastWord,1,5)) == "this." Or Lower(Substr(lcLastWord,1,9)) == "thisform." Or Lower(Substr(lcLastWord,1,12)) == "thisformset."  ) And ;
					this.EditorSource<>10 And Empty(This.ProcBaseClass)

				_EdInsert(This.EditorHwnd, ".", 1)
				If This.IntelliSense.Showed
					This.IntelliSense.Hide()
				Endif

				This.ShowErrorWriteTime(1929, Upper(Substr(lcLastWord,1,At(".",lcLastWord)-1))) 	&&- This, Thisform or ThisformSet can only be used within a method"
				llReturn = .F.
			Endif
		Endif


		*- verifico se devo abrir IntelliSense do FoxCodeplus, caso contrario fica por conta da IDE padrão do VFP.
		If llReturn
			If	( Substr(lcLastWord,1,1) = "." And ( Substr(Lower(lcLastWord),1,2) <> ".." And Lower(lcLastWord) <> ".null." And ;
					lower(lcLastWord) <> ".t." And Lower(lcLastWord) <> ".f.") And Lower(lcLastWord) <> ".and." And ;
					lower(lcLastWord) <> ".or." And Lower(lcLastWord) <> ".not." ) Or ;
					( Lower(lcLastWord) == "m." And Occurs(".",lcLastWord)=1 ) Or ;
					( Not Empty(lcFullTextControl) And This.EditorSource=10 ) Or ;
					( Lower(lcLastWord) == "this." And This.EditorSource<>10 And Occurs(".",lcLastWord)=1 ) Or ;
					( Lower(Substr(lcLastWord,1,5)) == "this." And This.EditorSource<>10 And Occurs(".",lcLastWord)=2 ) Or ;
					( lnCountVars > 0 And This.IsObjInWriteTime(lcLastWord2,@laItemsVars) ) Or ;
					( Lower(lcLastWord) <> "this." And Type(lcLastWord2) = "O" ) Or ;
					( ("[" $ lcText And "]." $ lcText) Or ("(" $ lcText And ")." $ lcText) ) Or ;
					( Used(lcLastWord2) ) Or ;
					( lnCountTables > 0 And Ascan(laItemsTables, lcLastWord2, -1,-1, 1, 15) > 0 And Not Empty(laItemsTables[1,1]))

				_EdInsert(This.EditorHwnd, ".", 1)
				llCheckFoxCodePlus = .T.
			Else
				llCheckFoxCodePlus = .F.
			Endif

			*- não devo abrir o foxcodeplus
			If Not llCheckFoxCodePlus
				This.IntelliSense.Hide()
				Keyboard "." Plain
				llReturn = .F.
			Endif
		Endif


		*- inicio a verificacao
		If llReturn
			This.IntelliSense.Found = .F.
			lnLines = 0

			Do Case

					*------------ dentro de um "WITH...ENDWITH" ------------*
					*- se a string começa com "." componho o "with...endwith"
				Case Substr(lcLastWord,1,1) = "." And ;
						( Substr(Lower(lcLastWord),1,2)<>".." And Lower(lcLastWord) <> ".null." And Lower(lcLastWord) <> ".t." And ;
						lower(lcLastWord) <> ".f." And Lower(lcLastWord) <> ".and." And Lower(lcLastWord) <> ".or." And ;
						lower(lcLastWord) <> ".not." )

					lcObjReference = This.GetWithHierarchy()
					If "&" $ lcObjReference Or Empty(lcObjReference)		&&- se o with/endwith trabalha com macro substituição não prossigo
						llReturn = .F.
					Endif

					If llReturn
						This.IncrementalResult = .F.

						*- WITH...ENDWITH DENTRO DE PRGS -*
						*- editor modify command (e outros diferente do form ou class designer)
						If This.EditorSource<>10
							Do Case
									*- this (classe definida no define class)
								Case lcObjReference == "this" And Not Empty(This.ProcBaseClass)
									lnLines = This.GetProcInfo(0,0,.F.)											&&-PMEs que eu adicionei a classe
									lnLines = lnLines + This.GetMembers(This.ProcBaseClass,.T.,.F.)				&&-PMEs nativas da classe

									*- somente objetos adicionado ao define class (add object ... as ...)
									*- abjetos criados dentro de objetos são desconsiderados pois aqui esta em writetime
								Case Substr(lcObjReference,1,4) == "this" And Occurs(".",lcObjReference) = 1 And Not Empty(This.ControlClassName)
									If Empty(This.ControlOleClass)
										lnLines = This.GetMembers(This.ControlClassName,.T.,.T.)
									Else
										Try
											loTempObj = Createobject("TempOleClass",This.ControlOleClass)
											lnLines = This.GetMembers(loTempObj.xOleControl,.T.,.T.)
											loTempObj = .Null.
										Catch
											lnLines = 0
										Endtry
									Endif

									*- objetos instanciados em write-time usando createobject(), createobjectex() ou newobject()
								Case lnCountVars > 0 And This.IsObjInWriteTime(lcObjReference, @laItemsVars)
									This.IncrementalResult = .F.
									lnLines = This.GetObjectWriteTime(lcObjReference, @laItemsVars)
									This.IncrementalResult = .T.

									*- objeto em memória (run-time)
								Otherwise
									lnLines = 0
									If Type("evaluate(lcObjReference)")="O"
										lnLines = This.GetMembers(Evaluate(lcObjReference),.T.,.T.)
									Endif
							Endcase

							*- WITH...ENDWITH DENTRO DO FORM OU CLASS DESIGNER-*
							*- editor do form ou class designer
						Else

							Do Case
									*- se for "this", "thisform" ou "thisformset" no form ou class desinger
								Case Not Empty(lcObjReference) And (;
										substr(lcObjReference,1,5) == "this." Or Substr(lcObjReference,1,9) == "thisform." Or Substr(lcObjReference,1,12) == "thisformset." Or ;
										inlist(lcObjReference, "this","thisform","thisformset"))

									lnLines = This.GetMembersDesignerTime(lcObjReference)

									*- objetos instanciados em write-time usando createobject(), createobjectex() ou newobject()
								Case lnCountVars > 0 And This.IsObjInWriteTime(lcObjReference, @laItemsVars)
									This.IncrementalResult = .F.
									lnLines = This.GetObjectWriteTime(lcObjReference, @laItemsVars)
									This.IncrementalResult = .T.

									*- objeto em memória (run-time)
								Otherwise
									lnLines = 0
									If Type("evaluate(lcObjReference)")="O"
										lnLines = This.GetMembers(Evaluate(lcObjReference),.T.,.T.)
									Endif
							Endcase
						Endif


						*- se encontrei itens para o IntelliSense limpo a propriedade para permitir que o IntelliSense seja apresentado..E
						If lnLines > 0
							This.TextLine = ""
						Endif

						This.IncrementalResult = .T.
					Endif


					*------------- FORA DO WITH...ENDWITH -----------------*
					*- DAQUI EM DIANTE TUDO FUNCIONA EM PRGS, FORMS E ETC -*
					*- lista de variaveis "m."
				Case Lower(lcLastWord) == "m." And Occurs(".",lcLastWord)=1
					This.IncrementalResult = .F.
					lnLines = This.GetProcInfo(5,1)
					This.IncrementalResult = .T.


					*- se for "this", "thisform" ou "thisformset" form ou class desinger
				Case Not Empty(lcFullTextControl) And This.EditorSource=10
					lnLines = This.GetMembersDesignerTime(lcFullTextControl)


					*- se for "this" - PMEs adicionadas e nativas da classe em um DEFINE CLASS no PRG
				Case Lower(lcLastWord) == "this." And This.EditorSource<>10 And Occurs(".",lcLastWord)=1

					This.IncrementalResult = .F.
					*- no prg busco somente as PME criadas em write-time da classe que estou posicionado
					*- e na sequencia busco as PMEs da baseclass da classe que estou posicionado
					lnLines = This.GetProcInfo(0,0,.F.)
					If Not Empty(This.ProcBaseClass)
						lnLines = lnLines + This.GetMembers(This.ProcBaseClass,.T.,.F.)		&&-ProcBaseClass foi capturada pela this.GetProcInfo(0,0) na linha acima
					Endif
					This.IncrementalResult = .T.


					*- se for "this" - PMEs nativas do objeto adicionado em um DEFINE CLASS no PRG
				Case Lower(Substr(lcLastWord,1,5)) == "this." And This.EditorSource<>10 And Occurs(".",lcLastWord)=2

					This.IncrementalResult = .F.
					This.GetProcInfo(4,0,.F.)

					lcObjName = Substr(lcLastWord,At(".",lcLastWord)+1)
					lcObjName = Substr(lcObjName,1, At(".",lcObjName)-1)
					lnRowFound = Ascan(This.ItemsObjects, lcObjName, -1,-1, 1, 15)

					If lnRowFound > 0
						This.ControlClassName = Alltrim(This.ItemsObjects[lnRowFound,2])
						This.ControlOleClass = Alltrim(This.ItemsObjects[lnRowFound,3])
					Else
						This.ControlClassName = ""
						This.ControlOleClass = ""
					Endif

					If Not Empty(This.ControlClassName)
						If Empty(This.ControlOleClass)
							lnLines = This.GetMembers(This.ControlClassName,.T.,.T.)
						Else
							Try
								loTempObj = Createobject("TempOleClass",This.ControlOleClass)
								lnLines = This.GetMembers(loTempObj.xOleControl,.T.,.T.)
								loTempObj = .Null.
							Catch
								lnLines = 0
							Endtry
						Endif
					Endif
					This.IncrementalResult = .T.


					*- objetos instanciados em write-time usando createobject(), createobjectex(), newobject() ou "for each"
				Case lnCountVars > 0 And This.IsObjInWriteTime(lcLastWord2, @laItemsVars)
					This.IncrementalResult = .F.
					lnLines = This.GetObjectWriteTime(lcLastWord2, @laItemsVars)
					This.IncrementalResult = .T.


					*- verifico se é um objeto de classe já instanciada em memória
				Case Lower(lcLastWord) <> "this." And Type(lcLastWord2) = "O"
					This.IncrementalResult = .F.
					lnLines = This.GetMembers(Evaluate(lcLastWord2),.T.,.T.)
					This.IncrementalResult = .T.


					*- collections ( _screen.forms[1] or grid.columns[1]. or grid.columns[1].header1. )
				Case ("[" $ lcText And "]." $ lcText) Or ("(" $ lcText And ")." $ lcText)
					lcObjName = Strtran(Strtran(Strtran(Strtran(Strtran(Alltrim(lcText),"  "," "), "[ ","["), " ]", "]"), "( ","("), " )", ")")
					lcObjName = Getwordnum(lcObjName, Getwordcount(lcObjName))
					lcObjName = Lower(Substr(lcObjName,1,Len(lcObjName)-1))

					If This.EditorSource = 10 And Substr(lcObjName,1,5) == "this." Or Substr(lcObjName,1,9) == "thisform." Or Substr(lcObjName,1,12) == "thisformset."
						Aselobj(laControl, 1)
						lcObjName = "laControl[1]." + Substr(lcObjName,At(".",lcObjName)+1)
					Endif

					If Type("evaluate(lcObjName)") = "O"
						This.IncrementalResult = .F.
						lnLines = This.GetMembers(Evaluate(lcObjName),.T.,.T.)
						This.IncrementalResult = .T.
					Endif


					*- campos de uma tabela ou cursor em run-time
				Case Used(lcLastWord2)
					This.IncrementalResult = .F.
					lnLines = This.GetFields(lcLastWord2)
					This.IncrementalResult = .T.


					*- campos de uma tabela, cursor ou no dataenvironment de forms ou reports em write-time
				Case Ascan(laItemsTables, lcLastWord2, -1,-1, 1, 15) > 0 And Not Empty(laItemsTables[1,1])
					This.IncrementalResult = .F.
					Local lnRowFound, lcAlias
					lnRowFound = 0
					lcAlias = ""

					lnRowFound = Ascan(laItemsTables, lcLastWord2, -1,-1, 1, 15)
					If lnRowFound > 0

						*- se for tabela do dataenvironment (free table)
						If  "cursorsource" $ Lower(laItemsTables[lnRowFound,2]) And "alias" $ Lower(laItemsTables[lnRowFound,2]) And "dataenvironment" $ Lower(laItemsTables[lnRowFound,2])
							lcFileDbf = Strextract(laItemsTables[lnRowFound,2],"CursorSource ",Chr(13))

							Try
								lcAlias = Alias()

								Use (lcFileDbf) Alias (lcLastWord2) In 0 Shared Again
								lnLines = This.GetFields(lcLastWord2)
								Use In (lcLastWord2)

								If Used(lcAlias)
									Select (lcAlias)
								Endif
							Catch
							Endtry

							*- obtenho os campos
						Else
							Try
								lcAlias =  Alias()
								Execscript(laItemsTables[lnRowFound,2])
								lnLines = This.GetFields(lcLastWord2)
								Use In (lcLastWord2)
								If Used(lcAlias)
									Select (lcAlias)
								Endif
							Catch
							Endtry
						Endif
					Endif

					This.IncrementalResult = .T.
			Endcase
		Endif

		If Type("This.IntelliSense") == "O" Then

			If (llReturn Or llSqlFields) And lnLines > 0
				This.IntelliSense.LastFind = ""
				This.IntelliSense.Find(This.LastWord)
				This.IntelliSense.Show()
			Else
				If This.IntelliSense.Showed
					This.IntelliSense.Hide()
				Endif
			Endif

		Endif

		*- retorno a configuracao do ambiente
		If lcSetTalk = "ON"
			Set Talk On
		Endif

		If lcSetNotify = "ON"
			Set Notify Cursor On
		Endif

		If lcSetExact = "OFF"
			Set Exact Off
		Endif

		Set Message To lcSetMessage

		Return llReturn
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Identifica se uma variavel tem um objeto instanciado em write-time.
	*/ atraves das funcoes "-createobject()", "createobjectex()" or "newobject()"
	*/------------------------------------------------------------------------------------------------
	Protected Procedure IsObjInWriteTime
		Lparameters plcWord, plaItemsVars
		Local lnRowFound, lcVarContent, llReturn

		Set Console Off

		If "." $ plcWord
			plcWord = Substr(plcWord,1,At(".",plcWord)-1)
		Endif

		lnRowFound = Ascan(plaItemsVars, plcWord, -1,-1, 1, 15)
		If lnRowFound > 0

			lcVarContent = Strtran(Substr(plaItemsVars[lnRowFound,4], At("=",plaItemsVars[lnRowFound,4])+1), " ", "")

			Do Case
					*- variaveis com objetos instanciados com as funcoes createobject(), createobjectex() or newobject()
				Case Inlist(Lower(Substr(lcVarContent,1,At("(",lcVarContent))), "createobject(", "createobjectex(", "newobject(")
					llReturn = .T.

					*- variaveis referenciadas com objetos ja instanciados (run-time ou designer-time)
				Case Not Isdigit(lcVarContent) And Not Inlist(Substr(lcVarContent,1,1),"'",'"',"[") And Not "<FOREACH>" $ plaItemsVars[lnRowFound,4]
					llReturn = .T.

					*- variaveis criadas com "for each"
				Case "<FOREACH>" $ plaItemsVars[lnRowFound,4]
					llReturn = .T.

					*- variavel nao contem um objeto em write-time
				Otherwise
					llReturn = .F.
			Endcase
		Endif

		Return llReturn
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busco as PMEs de um objeto instanciado em write-time.
	*/ Suporte para classes nativas do VFP, activex, classes em prg e vcx.
	*/ classes em exe, app, dll e fxp não são suportadas.
	*/------------------------------------------------------------------------------------------------
	Protected Procedure	GetObjectWriteTime
		Lparameters plcVarName, plaItemsVars
		Local lnRowFound, lcClassName, loTempObj, loTempSubObj, lnLines, lcAlias, lcVarName, ;
			lcFile, lnCountClasses, lcchkErrorToolTip, lcVarContent

		Local Array laVFPnativeClass[1]

		Set Console Off

		lnLines = 0
		lcVarName = plcVarName

		If "." $ plcVarName
			lcVarName = Substr(lcVarName,1,At(".",lcVarName)-1)
		Endif

		lnRowFound = Ascan(plaItemsVars, lcVarName, -1,-1, 1, 15)
		If lnRowFound = 0
			Return 0
		Endif

		lcVarContent = Lower(Strtran(Substr(plaItemsVars[lnRowFound,4], At("=",plaItemsVars[lnRowFound,4])+1), " ", ""))

		Do Case
			Case Inlist(Substr(lcVarContent,1,At("(",lcVarContent)), "createobject(", "createobjectex(", "newobject(")
				lcClassName = Lower(lcVarContent)
				lcClassName = Substr(lcClassName, At("(",lcClassName)+1)
				lcClassName = Alltrim(Substr(lcClassName, 1, At(Iif("," $ lcClassName,",",")"), lcClassName) -1))

				*- nome da classe literal
				If	(Substr(lcClassName,1,1) = "[" And Right(lcClassName,1) = "]") Or ;
						(Substr(lcClassName,1,1) = "'" And Right(lcClassName,1) = "'") Or ;
						(Substr(lcClassName,1,1) = '"' And Right(lcClassName,1) = '"')

					lcClassName = This.ClearQuotes(lcClassName)

					*- nome da classe atraves de variavel ou concatenacao
				Else
					lcClassName = "(NONE)"
				Endif

				If "." $ plcVarName
					lcClassName = lcClassName + Substr(plcVarName,At(".",plcVarName))
				Endif

				lcAlias = Alias()
				Select Code From FoxCodePlus Where Type = "T2" And Lower(Code) == lcClassName Into Array laVFPnativeClass

				Use In FoxCodePlus
				If Used(lcAlias)
					Select (lcAlias)
				Endif

				Do Case
						*- se for uma classe nativa do vfp ou uma activex
						*- funciona para createobject() e newobject()
					Case Inlist(Substr(lcVarContent,1,At("(",lcVarContent)), "createobject(", "newobject(") And Not "," $ lcVarContent
						If Alltrim(Lower(Evl(laVFPnativeClass[1],""))) == lcClassName Or "." $ lcClassName
							Try
								loTempObj = Createobject("TempObjWt",lcVarContent)

								If "." $ plcVarName
									loTempSubObj = Evaluate( "loTempObj.xObj"+Substr(plcVarName,At(".",plcVarName))  )
									loTempObj.xObj = loTempSubObj
								Endif

								lnLines = This.GetMembers(loTempObj.xObj,.T.,.T.)
								loTempObj = .Null.
								loTempSubObj = .Null.
							Catch
							Endtry

							*- se nao for uma classe nativa e nao tem o nome do arquivo
							*- verifico dentro do prg corrente se a classe existe.
						Else
							lcSafety = Set("Safety")
							Set Safety Off
							lcThisCode = This.GetText(0,130000)			&&- capturo o texto da janela atual
							Strtofile(lcThisCode, This.Tmpfile)
							Set Safety &lcSafety

							lcchkErrorToolTip = This.chkErrorToolTip
							This.chkErrorToolTip = "0"
							lnLines = This.GetProcInfoPrg(lcClassName, This.Tmpfile, .T., .T.)
							This.chkErrorToolTip = lcchkErrorToolTip
						Endif


						*- Se não achei nada acima, então...
						*- verifico se a classe informada pertence a um PRG invocado pelo SET PROCEDURE TO ou
						*- verifico se a classe informada pertence a uma VCX invocada pelo SET CLASSLIB TO
						If lnLines = 0
							lnCountClasses = This.GetProcInfo(8,1,.F.)									&&- aqui tambem valorizo this.ProcBaseClass
							If lnCountClasses > 0
								For lnx = 1  To lnCountClasses
									If This.Items[lnx,2] = 1 And Lower(This.Items[lnx,1]) == Lower(lcClassName)
										If "SET PROCEDURE TO" $ This.Items[lnx,4]
											lnLines = This.GetProcInfoPrg(lcClassName, Strextract(This.Items[lnx,4],"<FILE>","</FILE>"), .T., .T.)	&& getwordnum(this.Items[lnx,4],4)
										Else
											lnLines = This.GetProcInfoVcx(lcClassName, Strextract(This.Items[lnx,4],"<FILE>","</FILE>"))
										Endif
									Endif
								Endfor
							Endif
						Endif

						If lnLines = 0
							This.ShowErrorWriteTime(1733, Upper(lcClassName)) 	&&- class doesn't exist
						Endif



						*- se for funcao newobject() com um prg|vcx, faço a analise no arquivo para obter as pmes da classe.
						*- obs: caso exista somente o fxp nada é feito.
					Case Substr(lcVarContent,1,At("(",lcVarContent)) == "newobject(" And "," $ lcVarContent

						lcFile = Substr(lcVarContent, At(",",lcVarContent)+1)
						lcFile = Alltrim(Substr(lcFile, 1, At(Iif("," $ lcFile,",",")"), lcFile) -1))

						*- nome do arquivo literal
						If	(Substr(lcFile,1,1) = "[" And Right(lcFile,1) = "]") Or ;
								(Substr(lcFile,1,1) = "'" And Right(lcFile,1) = "'") Or ;
								(Substr(lcFile,1,1) = '"' And Right(lcFile,1) = '"')

							lcFile = This.ClearQuotes(lcFile)

							*- nome do arquivo atraves de variavel ou concatenacao
						Else
							lcFile = "(NONE)"
						Endif

						*- se nao tem extensao padronizo como vcx
						If Empty(Justext(lcFile))
							lcFile = Forceext(lcFile,"vcx")
						Endif

						*- se for uma compilação troco para prg para tentar achar o programa
						If Inlist(Lower(Justext(lcFile)),"fxp","app","exe","dll")
							lcFile = Forceext(lcFile,"prg")
						Endif

						*- busco as pmes conforme tipo do arquivo informado
						Do Case
							Case Inlist(Lower(Justext(lcFile)),"prg","mpr","spr","qpr")
								lnLines = This.GetProcInfoPrg(lcClassName, lcFile, .T., .T.)

							Case Lower(Justext(lcFile)) = "vcx"
								lnLines = This.GetProcInfoVcx(lcClassName, lcFile)

							Otherwise
								* no support
						Endcase
				Endcase


				*- objetos referenciados atraves de variaveis
			Case Not Isdigit(lcVarContent) And Not Inlist(Substr(lcVarContent,1,1),"'",'"',"[") And Not "<FOREACH>" $ plaItemsVars[lnRowFound,4]

				If Lower(plaItemsVars[lnRowFound,4]) # (This.LastWord) And "." $ This.LastWord
					lcVarContent = lcVarContent + "." + Substr(This.LastWord,At(".",This.LastWord)+1)
				Endif

				If This.EditorSource = 10
					lnLines = This.GetMembersDesignerTime(lcVarContent)
				Else
					If Type("evaluate(lcVarContent)") = "O"
						lnLines = This.GetMembers(Evaluate(lcVarContent),.T.,.T.)
					Endif
				Endif


				*- for each
			Case "<FOREACH>" $ plaItemsVars[lnRowFound,4]

				lcVarContent = Strextract(plaItemsVars[lnRowFound,4],"<FOREACH>","</FOREACH>") + "[1]"
				If Lower(plaItemsVars[lnRowFound,4]) # (This.LastWord) And "." $ This.LastWord
					lcVarContent = lcVarContent + "." + Substr(This.LastWord,At(".",This.LastWord)+1)
				Endif


				If This.EditorSource = 10
					lnLines = This.GetMembersDesignerTime(lcVarContent)
				Else
					If Type("evaluate(lcVarContent)") = "O"
						lnLines = This.GetMembers(Evaluate(lcVarContent),.T.,.T.)
					Endif
				Endif

		Endcase

		Return lnLines
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ busco os membro de um controle em um form ou class designer em tempo de designer
	*/ usado para abrir o IntelliSense em with..endwith e substituir o IntelliSense de forms e classes
	*/ plcFullTextControl -> Ex: thisform / this
	*/                           thisform.cmdOpen / this.parent
	*/                           thisformset.formx.grid1
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetMembersDesignerTime
		Lparameters plcFullTextControl
		Local lnLines, loControl, lcParentControl, lcAuxControl, lcWhichThis
		Local Array laControl[1]

		Set Console Off

		lcWhichThis = Lower(Iif("." $ plcFullTextControl, Substr(plcFullTextControl,1,At(".",plcFullTextControl)-1), plcFullTextControl)) + "."
		lnLines = 0
		This.ProcBaseClass = ""
		This.ProcClass = ""

		Aselobj(laControl, 1)
		loControl = laControl[1]

		If Vartype(loControl)="O"
			*- THISFORM or THISFORMSET
			If Inlist(lcWhichThis, "thisform.", "thisformset.")
				lcParentControl = "loControl"
				Do While .T.
					loParentControl = Iif(Type("evaluate(lcParentControl)")="O", Evaluate(lcParentControl), .Null.)
					If Vartype(loParentControl) = "O" And Lower(loParentControl.BaseClass) = Iif(lcWhichThis == "thisform.", "form", "formset")
						Exit
					Endif
					lcParentControl = lcParentControl + ".parent"
				Enddo
				lcAuxControl = lcParentControl + Iif("." $ plcFullTextControl, Substr(plcFullTextControl,At(".",plcFullTextControl)), "")

				*- THIS
			Else
				lcAuxControl = _wtitle(This.EditorHwnd)
				lcAuxControl = Substr(lcAuxControl,1,At(".",lcAuxControl)-1)
				loControl = Iif(Type("loControl.activecontrol")="O", loControl.ActiveControl, loControl)

				If Lower(loControl.Name) == Lower(lcAuxControl)
					lcAuxControl = "loControl"
				Else
					lcAuxControl = "loControl."+lcAuxControl
				Endif
				lcAuxControl = lcAuxControl + Iif("." $ plcFullTextControl, Substr(plcFullTextControl,At(".",plcFullTextControl)), "")
			Endif

			*- busco os membros do controle
			If Type("evaluate(lcAuxControl)") = "O"
				lnLines = This.GetMembers(Evaluate(lcAuxControl),.T.,.T.)
			Endif
		Endif

		Return lnLines
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Ao pressionar igual "=" abre um novo IntelliSense para algumas propriedades
	*/ Essa funcao interage diretamente com o Foxcode.App e Foxcode.dbf e só funciona se o
	*/ Foxcodeplus estiver ativado.
	*/------------------------------------------------------------------------------------------------
	Procedure GetEqual
		Local lcProperty, lcFullProperty, lcTextLine, loControl, lcSavecliptext, lcTextSelected, lcSetTalk, ;
			lcSetNotify, lcSetMessage, lcSetExact, llReturn

		Set Console Off

		lcSetTalk = Set("talk")
		lcSetNotify = Set("Notify",2)
		lcSetMessage = Set("Message")
		lcSetExact = Set("exact")

		Set Talk Off
		Set Notify Cursor Off
		Set Exact On
		Set Message To ""

		llReturn = .T.
		loControl = .Null.
		This.LoadScriptBoolean = .F.


		If Isnull(This.IntelliSense)
			This.IntelliSense = Newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")
		Endif

		*- verifico em qual lugar do vfp estou acionando o IntelliSense
		If Not This.SetWontop("00//01//08//10//12")
			*- verifico se é um objeto do vfp
			Local loCurrentObject, lcActiveClass, lnSelStart, llReadOnly
			Try
				loCurrentObject = _Screen.ActiveForm.ActiveControl
				lcActiveClass = loCurrentObject.BaseClass
				lnSelStart    = loCurrentObject.SelStart
				llReadOnly    = loCurrentObject.ReadOnly
			Catch
				loCurrentObject = .Null.
				lcActiveClass = ""
				lnSelStart = 0
				llReadOnly = .F.
			Endtry

			*- se for um objeto do vfp faço tratamento do "."
			If Lower(lcActiveClass) $ "//textbox//editbox//combobox//" And lnSelStart > 0 And Not llReadOnly
				If Lower(lcActiveClass) = "combobox"
					lcValue = loCurrentObject.Text
					loCurrentObject.DisplayValue = Substr(lcValue, 1, lnSelStart) + "=" + Substr(lcValue, lnSelStart+1)
				Else
					If Vartype(loCurrentObject.Value) = "C"
						lcValue = loCurrentObject.Value
						loCurrentObject.Value = Substr(lcValue, 1, lnSelStart) + "=" + Substr(lcValue, lnSelStart+1)
					Endif
				Endif
				Keyboard '{RightArrow}{home}'+Replicate('{RightArrow}',lnSelStart+1) Plain
			Else
				Keyboard "=" Plain
			Endif

			llReturn = .F.
		Endif


		If llReturn
			*- se o IntelliSense do foxcodeplus estiver aberto e pressiono "="
			*- seleciono o item da lista.
			If This.IntelliSense.Showed
				This.SelectItem(0)
			Endif

			*- se tiver texto selecionado eu deleto mas antes checo se nao foi uma auto-selecao da IDE do VFP.
			*- a auto-selecao acontece qdo fechamos parentes ou colchetes ")" ou "]" para indicar a funcao ou array que o abriu.
			lcSavecliptext = _Cliptext
			_EdCopy(This.EditorHwnd)
			lcTextSelected = _Cliptext
			lcTextSelected = Iif(Substr(lcTextSelected,1,1)="[", Stuff(lcTextSelected,1,1,"("), lcTextSelected)
			lcTextSelected = Iif(Right(lcTextSelected,1)="]", Stuff(lcTextSelected,Len(lcTextSelected),1,")"), lcTextSelected)
			_Cliptext = lcSavecliptext

			If	( Not Empty(lcTextSelected) And Substr(lcTextSelected,1,1) <> "(" And Right(lcTextSelected,1) <> ")" )
				_EdDelete(This.EditorHwnd)
			Endif

			*- busco o ultima palavra digitada antes do "=" e verifico se a linha é um comentario
			lcTextLine = This.TreatLine(This.GetTextLine()+"=")
			If This.IsComment
				This.CursorPos = _EdGetPos(This.EditorHwnd)
				llReturn = .F.
			Endif

			*- se estiver no editor do form ou classe
			If This.EditorSource=10
				Local Array laControl[1]
				laControl[1] = ""

				Aselobj(laControl, 1 )
				If Vartype(laControl[1]) = "O"
					loControl = laControl[1]
				Endif
			Else
				*- obtenho o objeto para verificar se a propriedade em questão é boolean
				This.GetWithHierarchy()		&&-capturo a classe e controle
				Try
					Do Case
						Case Not Empty(This.ControlClassName)
							loControl = Createobject(This.ControlClassName)
						Case Not Empty(This.ProcBaseClass)
							loControl = Createobject(This.ProcBaseClass)
						Otherwise
							loControl = Createobject("empty")
					Endcase
				Catch
					loControl = Createobject("empty")
				Endtry
			Endif

			*- funciona somente se for uma propriedade que possa obter um valor lógico
			If "." $ lcTextLine And "=" $ lcTextLine And Empty(Substr(lcTextLine,At("=",lcTextLine)+1))
				lcProperty = Substr(lcTextLine, Rat(".",lcTextLine))
				lcProperty = Lower(Alltrim(Substr(lcProperty, 1, At("=",lcProperty)-1)))
				lcFullProperty = Alltrim(Substr(lcTextLine, 1, At("=",lcTextLine)-1))

				*- boolean properties (true or false)
				If	Type("EVALUATE('_screen"+lcProperty+"')") = "L" Or;
						type("EVALUATE('_vfp"+lcProperty+"')") = "L" Or;
						type("EVALUATE('application"+lcProperty+"')") = "L" Or;
						type("EVALUATE('loControl"+lcProperty+"')") = "L" Or;
						( Not Empty(lcFullProperty) And Type("EVALUATE(lcFullProperty)") = "L" )

					This.LoadScriptBoolean = .T.

					*- aciono o IntelliSense do foxcode.app e foxcode.dbf
					_EdInsert(This.EditorHwnd, "=", 1)

					Keyboard " "
				ELSE
					_EdInsert(This.EditorHwnd, "=", 1)
*					Keyboard "=" Plain
				Endif
			ELSE
				_EdInsert(This.EditorHwnd, "=", 1)
*				Keyboard "=" Plain
			Endif
		Endif

		Try
			This.CursorPos = _EdGetPos(This.EditorHwnd)
		Catch
		Endtry

		*- retorno a configuracao do ambiente
		If lcSetTalk = "ON"
			Set Talk On
		Endif

		If lcSetNotify = "ON"
			Set Notify Cursor On
		Endif

		If lcSetExact = "OFF"
			Set Exact Off
		Endif

		Set Message To lcSetMessage

		Return llReturn
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ trata as palavras de uma string separando-as de sinais matematicos e outros
	*/------------------------------------------------------------------------------------------------
	Protected Procedure TreatWords
		Lparameters plcText
		plcText = Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(plcText,"+"," + "), "-"," - "), "*"," * "), "/"," / "), "="," = "), "#"," # "), ";"," ; "), ",", " , ")
		plcText = Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(plcText,"<","< "), ">"," >"), "[","[ "), "]"," ]"), "(","( "), ")"," )"), "  ", " "), "$", " $ ")
		Return plcText
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ trata linha retirando identação, quebra de linha, etc.
	*/------------------------------------------------------------------------------------------------
	Protected Procedure TreatLine
		Lparameters plcLineText, plnLineNo, plaCodePrg

		Set Console Off

		plnLineNo = Evl(plnLineNo, -1)
		This.IsComment = .F.

		*- removo os TABs and quebra de linha
		plcLineText = Alltrim(Strtran(Strtran(Strtran(plcLineText, Chr(9)," "), Chr(10),""), Chr(13),""))

		*- desconsidero os comentario contidos na mesma linha do codigo
		If "&"+"&" $ plcLineText
			plcLineText = Alltrim( Substr(plcLineText, 1, At("&"+"&", plcLineText)-2) )
		Endif

		*- caso tenha ";" ao final da linha retiro a quebra e coloco todo o codigo na mesma linha
		If plnLineNo >= 1
			Do While Right(plcLineText ,1) = ";"
				plcLineText = Substr(plcLineText, 1, Len(plcLineText)-1) + " "
				If plnLineNo+1 < Alen(plaCodePrg,1)
					plnLineNo = plnLineNo + 1
					plcLineText = plcLineText + This.TreatLine(plaCodePrg[plnLineNo])		&&- recursive called
				Endif
			Enddo
		Endif

		*- se a linha for um comentario retorno vazio
		If Substr(plcLineText,1,1)="*" Or Substr(plcLineText,1,2)="&"+"&" Or Substr(Lower(plcLineText),1,5)="note "
			This.IsComment = .T.
			plcLineText = ""
		Endif

		Return plcLineText
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ se for TEXT...ENDTEXT desconsidero o bloco de codigo
	*/------------------------------------------------------------------------------------------------
	Protected Procedure SkipTextEndText
		Lparameters plcLineText, plnLineNo, plaCodePrg

		Set Console Off

		If Substr(Lower(plcLineText),1,5) == "text "
			Do While .T.
				plnLineNo = plnLineNo + 1
				If Substr(Lower(plaCodePrg[plnLineNo]),1,5) == "endte" Or plnLineNo = Alen(plaCodePrg,1)
					Exit
				Else
					plaCodePrg[plnLineNo] = ""
				Endif
			Enddo
			Return .T.
		Else
			Return .F.
		Endif
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Adiciona um item a lista do IntelliSense (somente no tabpage1)
	*/------------------------------------------------------------------------------------------------
	Protected Procedure AddItem
		Lparameters plcItem, plnImage, plcToolTip
		Local lnItemWidth, llReturn

		Set Console Off

		*- acho a maior largura entre os itens
		With This.IntelliSense.Tabs.TabPage1.Items
			lnItemWidth = Txtwidth(plcItem, .FontName, .FontSize) * This.IntelliSense.AvgCharWidth
			This.MaxWidth = Max(This.MaxWidth, lnItemWidth)
		Endwith

		llReturn = This.IntelliSense.AddItem(plcItem, plnImage, plcToolTip)
		Return llReturn
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Compilacao em write-time
	*/ Retorna Array com os erros encontrados apos compilação em write-time
	*/ A compilação é feita somente em PRGs
	*/------------------------------------------------------------------------------------------------
	Procedure GetErrorList
		Set Console Off

		If This.chkErrorList <> "1" Or Vartype(This.FormErrorList)<>"O"
			Return
		Endif

		Local lcSafety, lcNotify, lcLogerrors, lcErrors, lcThisCode, lcCodePrg, lnx, lnz, lny, laAuxErrors, lcDescr, lnLine, ;
			lcTextLine, lnTotLines, lnCurrentLine, lnStartLine, lnEndLine, llFoundProc, lnTotVars

		lcSafety = Set("Safety")
		lcNotify = Set("Notify")
		lcLogerrors = Set("Logerrors")
		lnz = 0

		Local Array laErrors[1,3], laAuxErrors[1], laCodePrg[1]
		Set Notify Off
		Set Logerrors On
		Set Safety Off

		lcThisCode = This.GetText(0,130000)			&&- capturo o texto da janela atual

		If This.EditorSource = 10
			lcThisCode = "define class TmpClass as form"+Chr(13)+;
				"procedure TmpProc"+Chr(13)+;
				lcThisCode+Chr(13)+;
				"endproc"+Chr(13)+;
				"enddefine"
		Endif

		Strtofile(lcThisCode, This.Tmpfile)
		Compile (This.Tmpfile)

		Set Notify &lcSafety
		Set Logerrors &lcNotify
		Set Safety &lcLogerrors

		*----------------------
		*- erros de compilação e write-time
		*----------------------
		If File(Forceext(This.Tmpfile,"err"))
			lcErrors = Filetostr(Forceext(This.Tmpfile,"err"))
			Alines(laAuxErrors,lcErrors)

			lcCodePrg = ""

			For lnx = 1 To Alen(laAuxErrors,1)
				If Lower(Substr(laAuxErrors[lnx],1,13)) == "error in line"	And ":" $ laAuxErrors[lnx]
					lcDescr = Substr(laAuxErrors[lnx],At(" ",laAuxErrors[lnx],3))
					lnLine = Val(Substr(lcDescr,1,At(":",lcDescr)-1)) - Iif(This.EditorSource = 10, 2, 0)
					lcDescr = Alltrim(Substr(lcDescr,At(":",lcDescr)+1))
					lnz = lnz + 1

					Dimension laErrors[lnz,3]
					laErrors[lnz,1] = 1
					laErrors[lnz,2] = lcDescr + ".. " + lcCodePrg
					laErrors[lnz,3] = lnLine
					lcCodePrg = ""
				Else
					lcCodePrg = Alltrim(laAuxErrors[lnx])
				Endif
			Endfor
		Endif

		*----------------------
		*- warnings
		*----------------------
		*- still not avaible


		*- insiro o que encontrei
		This.FormErrorList.LoadErrorList(@laErrors, Iif(This.EditorSource = 10, _wtitle(This.EditorHwnd), This.EditorFileName), This.EditorHwnd)
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Inclue um summary para um metodo, evento ou função quando invocado por "***"
	*/ para funcionar devo estar obrigatorimente uma linha acima de uma procedure
	*/------------------------------------------------------------------------------------------------
	Protected Procedure SetSummary
		Set Console Off

		Local lcFullLine, lnLineNo
		lnLineNo = This.GetLineNo()
		lcFullLine = This.GetTextLine(lnLineNo,.T.)

		*- verifico se estou invocando o summary
		If Alltrim(Strtran(Strtran(Strtran(lcFullLine, Chr(9),""), Chr(10),""), Chr(13),"")) == "***"
			Local lcIndent, lnx, lcParamName, lcSummary, lcParamNames, lnGoToPos, lnCountPars
			lcIndent = ""
			lcParamName = ""
			lcParamNames = ""
			lnx = 0

			*- capturo a identação
			For lnx = 1 To Len(lcFullLine)
				If Inlist(Substr(lcFullLine,lnx,1), Chr(9), Chr(32))
					lcIndent = lcIndent + Substr(lcFullLine,lnx,1)
				Else
					Exit
				Endif
			Endfor

			*- capturo a linha abaixo para verificar se é uma procedure
			lnx = 0
			lcFullLine = ""
			Do While .T.
				lnx = lnx + 1
				lcFullLine = lcFullLine + This.GetTextLine(lnLineNo+lnx ,.T.)
				lcFullLine = Alltrim(Strtran(Strtran(Strtran(lcFullLine, Chr(9),""), Chr(10),""), Chr(13),""))
				If Right(lcFullLine ,1) = ";"
					lcFullLine = Strt(lcFullLine, ";","")
				Else
					Exit
				Endif
			Enddo

			*- se for uma procedure insiro o Summary
			If This.IsProc(lcFullLine)

				*- capturo os parametros da procedure quando declarado na mesma linha de criacao da procedure
				If "(" $ lcFullLine And ")" $ lcFullLine
					lcFullLine = Strtran(Strtran(lcFullLine,"(",","),")",",")
					For lnCountPars = 1 To 99
						lcParamName = Alltrim(Strextract(lcFullLine,",",",",lnCountPars))
						lcParamName = Getwordnum(lcParamName,1)		&&- se for parametro tipado considero somente o nome do parametro
						If Not Empty(lcParamName)
							lcParamNames = lcParamNames + lcIndent + [*** <param name="]+lcParamName+["></param>]+Chr(13)
						Else
							Exit
						Endif
					Endfor

					*- parametros declarados com "parameters" ou "lparameters"
				Else
					*- continuo a captura das linhas abaixo para verificar se é uma declaração de parametros
					lcFullLine = ""
					Do While .T.
						lnx = lnx + 1
						lcFullLine = lcFullLine + This.GetTextLine(lnLineNo+lnx,.T.)
						lcFullLine = Alltrim(Strtran(Strtran(Strtran(lcFullLine, Chr(9),""), Chr(10),""), Chr(13),""))

						If Right(lcFullLine ,1) = ";"
							lcFullLine = Strt(lcFullLine, ";","")
						Else
							Exit
						Endif
					Enddo

					If Getwordcount(lcFullLine) >= 2 And Inlist(Lower(Substr(Getwordnum(lcFullLine,1),1,4)),"lpar","para")

						lcFullLine = ","+Alltrim(Substr(lcFullLine, At(" ",lcFullLine)))+","
						For lnCountPars = 1 To 99
							lcParamName = Alltrim(Strextract(lcFullLine,",",",",lnCountPars))
							lcParamName = Getwordnum(lcParamName,1)		&&- se for parametro tipado considero somente o nome do parametro
							If Not Empty(lcParamName)
								lcParamNames = lcParamNames + lcIndent + [*** <param name="]+lcParamName+["></param>]+Chr(13)
							Else
								Exit
							Endif
						Endfor

					Endif
				Endif

				*- insiro o summary
				lcSummary =	[ <summary>]+Chr(13)+;
					lcIndent+[*** ]+Chr(13)+;
					lcIndent+[*** </summary>]+Chr(13)+;
					lcParamNames+;
					lcIndent+[*** <remarks></remarks>]

				_EdInsert(This.EditorHwnd, lcSummary, Len(lcSummary))


				*- ja posiciono no lugar correto para digitacao do texto summary
				lnGoToPos = _EdGetLPos(This.EditorHwnd, lnLineNo)
				lnGoToPos = lnGoToPos + Len(lcIndent) + 4
				_EdSetPos(This.EditorHwnd, lnGoToPos)
			Endif
		Endif

		Return
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Faço a leitura do summary da procedure para inserir no tooltip do IntelliSense
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetSummary
		Lparameters plaCodePrg, plnx, pllNoParam
		Local lcFullLine, lnx, lnLineStartComment, lnLineEndComment, lcText, lcSummary, lnCountPars, lcPar, lcPars, lcReturn

		lcFullLine = ""
		lcReturn = ""
		lnLineStartComment = 0
		lnLineEndComment = 0

		Set Console Off

		If plnx <= 1
			Return ""
		Endif

		*- verifico se a linha acima onde a procedure, metodo ou evento está criada pode ser um summary
		For lnx = 1 To 99
			If (plnx - lnx) <= 0
				Exit
			Endif

			lcFullLine = Alltrim(Strtran(Strtran(Strtran(Strtran(plaCodePrg[plnx - lnx], Chr(9),""), Chr(10),""), Chr(13),""), "  "," "))

			If Substr(lcFullLine,1,3) == "***"
				*- linha onde começa a comentario do summary
				If Lower(Substr(lcFullLine,1,14)) == "*** <summary>"
					lnLineStartComment = (plnx - lnx)
					Exit
				Endif

				*- linha onde termina a comentario
				If lnLineEndComment = 0
					lnLineEndComment = (plnx - lnx)
				Endif
			Else
				If lnLineEndComment = 0
					Return ""
				Else
					Exit
				Endif
			Endif
		Endfor

		If lnLineStartComment = 0
			Return ""
		Endif


		*- capturo o bloco de texto do comentario
		lcText = ""
		For lnx = lnLineStartComment To lnLineEndComment
			lcText = lcText + Alltrim(Strtran(plaCodePrg[lnx], Chr(9),""))
		Endfor
		lcText = Strtran(lcText,"*","")

		*- capturo o <summary>
		lcSummary = Alltrim(Strextract(lcText,"<summary>","</summary>"))

		*- capturo o <remarks>
		lcRemarks = Alltrim(Strextract(lcText,"<remarks>","</remarks>"))

		*- capturo os <param name="">
		lnCountPars = Occurs("<param name",lcText)
		lcPars = ""
		If lnCountPars > 0
			For lnx = 1 To lnCountPars
				lcPar = Strtran(Alltrim(Strextract(lcText,"<param name=","</param>",lnx)), "  ", " ")
				lcPar = Strtran(This.ClearQuotes(lcPar), ">", ": ")
				lcPars = lcPars + lcPar + Iif(lnx<lnCountPars, Chr(13), "")
			Endfor
		Endif

		*- Texto para o tooltip
		lcReturn = Iif(!Empty(lcSummary), Chr(13)+lcSummary, "") + ;
			iif(!Empty(lcPars) And !pllNoParam, Chr(13)+lcPars, "")


		Return lcReturn
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Verifica se o comando ou funcao tem code snippet
	*/ para ativar deve-se pressionar SPACE ao lado do item selecionado
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetCodeSnippets
		Lparameters plcWord
		Local lnItemsFound, lnLines, lnx, lcItem, lcToolTip

		Set Console Off

		If This.chkCodeSnippet <> "1" Or Empty(plcWord)
			Return 0
		Endif

		lnLines = Alen(This.ItemsCodeSnippets,1)
		lnItemsFound = 0

		*- functions
		If This.WordCount >= 2
			For lnx = 1 To lnLines
				lcItem = Proper(Alltrim(This.ItemsCodeSnippets[lnx,1]))
				If ("( )" $ This.ItemsCodeSnippets[lnx,2] Or "()" $ This.ItemsCodeSnippets[lnx,2]) And This.ChkIncremental(plcWord,lcItem)
					lnItemsFound = lnItemsFound + 1
					lcToolTip = "Code Snippet " + lcItem + " for expansion of "+Alltrim(This.ItemsCodeSnippets[lnx,2])
					This.AddItem(lcItem, 18, lcToolTip)
				Endif
			Endfor

			*- commands
		Else
			For lnx = 1 To lnLines
				lcItem = Proper(Alltrim(This.ItemsCodeSnippets[lnx,1]))
				If This.ChkIncremental(plcWord,lcItem)
					lnItemsFound = lnItemsFound + 1
					lcToolTip = "Code Snippet " + lcItem + " for expansion of "+Alltrim(This.ItemsCodeSnippets[lnx,2])
					This.AddItem(lcItem, 18, lcToolTip)
				Endif
			Endfor
		Endif

		Return lnItemsFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ apresento o erro em forma de tooltip conforme o codigo do plnCode
	*/------------------------------------------------------------------------------------------------
	Procedure ShowErrorWriteTime
		Lparameters plnCode, plcInfo1, plcInfo2, plcInfo3
		plcInfo1 = Evl(plcInfo1,"")
		plcInfo2 = Evl(plcInfo2,"")
		plcInfo3 = Evl(plcInfo3,"")

		Set Console Off

		*- foxcode.app
		If This.chkErrorToolTip <> "1"
			Return
		Endif

		*- custom messages
		If Empty(plnCode)
			*- plcInfo1 -> Message
			*- plcInfo2 -> Title
			This.ShowToolTipEditor(plcInfo1, plcInfo2, 3, .T., .T.)

			*- VFP default messages
		Else
			Local lcAlias, lcMessage, lcTitle
			lcAlias = Alias()

			Use foxcodeplus3 Alias foxcodeplus3 In 0 Shared
			If Seek(plnCode,"foxcodeplus3","code")
				lcMessage = Strtran(Strtran(Strtran(Alltrim(foxcodeplus3.Message), "@1", plcInfo1), "@2", plcInfo2), "@3", plcInfo3)
				lcTitle = Strtran(Strtran(Strtran(Alltrim(foxcodeplus3.Title), "@1", plcInfo1), "@2", plcInfo2), "@3", plcInfo3)
				This.ShowToolTipEditor(lcMessage,  lcTitle + " (Error "+Alltrim(Str(foxcodeplus3.Code))+")", 3, .T., .T.)
			Endif
			Use In foxcodeplus3

			If Used(lcAlias)
				Select (lcAlias)
			Endif
		Endif
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Busco as assinaturas das funcoes e metodos da linha atual
	*/------------------------------------------------------------------------------------------------
	Procedure GetSignature
		Local lcText, lcProc, lcMethod, lnLines, lnLen, lnx, lnPos, lnRowFound, lnOpenedProc, lcChar, lcSignature ,;
			lcProcPar, lnProcPar, lcToolTip, lcDescription, lcParName, llGetDescription, lcObj, llProcFound

		Local Array laItems[1,3]

		Set Console Off

		*- verifica se o VFP esta parametrizado para apresentar tooltip para assinatura de metodos e funcoes
		If Not "Q" $ _vfp.EditorOptions 	&&- Quick Info
			Return
		Endif

		*- capturo a linha até a posicao do cursor no texto
		lcText = This.TreatLine(This.GetTextLine())

		*- identifico o nome da funcao ou metodo ao qual estou posicionado
		lnLen = Len(lcText)
		lnOpenedProc = 0
		lcProc = ""

		For lnx = 1 To lnLen
			lnPos = (lnLen - lnx) + 1
			lcChar = Substr(lcText,lnPos,1)

			If lcChar=")" And Not This.IsInQuotes(Substr(lcText,1,lnPos))
				lnOpenedProc = lnOpenedProc + 1
			Else
				If lcChar="(" And Not This.IsInQuotes(Substr(lcText,1,lnPos))
					If lnOpenedProc > 0
						lnOpenedProc = lnOpenedProc - 1
						Loop
					Endif
					lcProc = Alltrim(This.TreatWords( Substr(lcText,1,lnPos-1) ))
					If " " $ lcProc
						lcProc = Alltrim(Substr(lcProc, Rat(" ",lcProc)))
					Endif
					Exit
				Endif
			Endif
		Endfor

		If Empty(lcProc) Or Not Isalpha(lcProc)
			Return
		Endif


		*- identifico o numero do parametro que estou posicionado dentro da funcao que estou posicionado
		lcProcPar = Substr(lcText, lnPos+1)
		lnProcPar = 1
		lnOpenedProc = 0

		If Not Empty(lcProcPar) 				&&and "," $ lcProcPar
			For lnx = 1 To Len(lcProcPar)

				*- verifico se no parametro contem funcao para buscar a "," correta e nao a "," que esta dentro da funcao
				If Substr(lcProcPar,lnx,1) = "(" And Not This.IsInQuotes(Substr(lcProcPar,1,lnx))
					lnOpenedProc = lnOpenedProc + 1
					Do While .T.
						lnx = lnx + 1
						If lnx > Len(lcProcPar)
							Exit
						Endif

						If Substr(lcProcPar,lnx,1) = "(" And Not This.IsInQuotes(Substr(lcProcPar,1,lnx))
							lnOpenedProc = lnOpenedProc + 1
						Else
							If Substr(lcProcPar,lnx,1) = ")" And Not This.IsInQuotes(Substr(lcProcPar,1,lnx))
								lnOpenedProc = lnOpenedProc - 1
								If lnOpenedProc <= 0
									Exit
								Endif
							Endif
						Endif
					Enddo
					Loop
				Endif

				*- se for virgula incremento a variavel que controla o numero do parametro posicionado.
				If Substr(lcProcPar,lnx,1) = ","
					If Not This.IsInQuotes( Substr(lcProcPar,1,lnx) )
						lnProcPar = lnProcPar + 1
					Endif
				Endif
			Endfor
		Endif


		lcSignature = ""
		lcDescription = ""
		llProcFound = .F.

		*- busco a assinatura para o METODO encontrado na classe corrente do PRG
		If "." $ lcProc And This.EditorSource <> 10
			lcObj = Alltrim(Lower(Substr(lcProc, 1, Rat(".", lcProc)-1)))
			lcMethod = Substr(lcProc, Rat(".", lcProc)+1)
			lcSignature = ""
			llGetDescription = .T.

			This.IncrementalResult = .F.

			Do Case
					*- "this" dentro de "define class" no prg
				Case lcObj == "this" And Occurs(".",lcProc) = 1
					lnLines = This.GetProcInfo(0,0,.F.)

					*- outros pontos ainda serão criados abaixo
					*- others inclusions should be included below
				Case 1 = 0

				Otherwise
					Return
			Endcase

			This.IncrementalResult = .T.

			lnRowFound = This.SeekItem(lcMethod, 3)
			If lnRowFound > 0
				llProcFound = .T.
				lcProc = Lower(This.ProcClass) + "." + Alltrim(This.Items[lnRowFound,1])
				lcSignature = Alltrim(Strextract(This.Items[lnRowFound,3], "(",")"))
				lcDescription = Alltrim(Strextract(This.Items[lnRowFound,3], Chr(13), Chr(13)))
			Endif


			*- busco a assinatura para a FUNCAO encontrada
		Else
			lcSignature = ""
			lnRowFound = Ascan(This.FoxcodeFunctions, lcProc, -1,-1, 1, 15)

			*- verifico se é uma funcao nativa do VFP
			*if not empty(laItems[1,2])
			If Not Empty(lnRowFound)
				*llGetDescription = .f.
				*lnLines = 1
				*lcProc = alltrim(laItems[1,2])
				*lcSignature = alltrim(strtran(strtran(strtran(strextract(laItems[1,3], "(",")"), chr(13), ""), chr(10), ""), chr(9), "") )
				*lcDescription = alltrim(substr(laItems[1,3],at(chr(13),laItems[1,3])+1))
				Return

				*- verifico se a funcao existe no prg corrente ou "set procedure to" invocado no prg corrente
			Else
				llGetDescription = .T.
				This.IncrementalResult = .F.
				lnLines = This.GetProcInfo(0,1,.F.)			&&- procuro no prg corrente
				lnRowFound = This.SeekItem(lcProc, 19)
				If lnRowFound = 0
					This.GetProcInfo(8,1,.F.)				&&- procuro nos SET PROCEDURES do prg ou method corrente
					lnRowFound = This.SeekItem(lcProc, 19)
				Endif
				This.IncrementalResult = .T.

				lnRowFound = This.SeekItem(lcProc, 19)
				If lnRowFound > 0
					llProcFound = .T.
					lcProc = Alltrim(This.Items[lnRowFound,1])
					lcSignature = Alltrim(Strextract(This.Items[lnRowFound,3], "(",")"))
					lcDescription = Alltrim(Strextract(This.Items[lnRowFound,3], Chr(13), Chr(13)))
				Endif
			Endif
		Endif

		If Not llProcFound
			Return
		Endif

		*- apago o tooltip nativo caso esteja aberto.
		*- se cheguei ate aqui é pq é necessario substituir o tooltip nativo do VFP, ou, nao existe tooltip nativo.
		_wSelect(This.EditorHwnd)

		*- quantidade de parametros encotrados
		lnCountPars = Iif(","$lcSignature, Occurs(",",lcSignature)+1, 1)
		lnCountPars = Iif(Empty(lcSignature), 0, lnCountPars)

		*- sem assinatura
		If Empty(lcSignature)
			lcToolTip = lcProc+"( )" + Chr(13) + lcDescription

			*- com assinatura
		Else
			*- capturo a assinatura e o nome do parametro posicionado
			lcSignature = Strtran(Strtran(Strtran(lcSignature,"[",""), "]",""), " ","")
			lcSignature = "<PAR>" + Strtran(lcSignature,",","</PAR><PAR>") + "</PAR>"
			lcParName = Strextract(lcSignature,"<PAR>","</PAR>",lnProcPar)

			*- destaco o parametro posicionado
			lcSignature = Strtran(lcSignature,"<PAR>"+lcParName+"</PAR>","<PAR><<"+lcParName+">></PAR>")
			lcSignature = Strtran(Strtran(Strtran(lcSignature,"</PAR><PAR>",", "), "<PAR>", ""), "</PAR>", "")

			*- descricao do parametro posicionado
			If Empty(lcParName)
				lcAuxParName = Alltrim(Str(lnProcPar)) + ". (INVALID PARAMETER)"
			Else
				If llGetDescription
					lcAuxParName = Strtran(Strextract(This.Items[lnRowFound,3]+Chr(13), lcParName+":", Chr(13)), Chr(10), "")
					lcAuxParName = Alltrim(Str(lnProcPar)) + ". " + lcParName + ":" + lcAuxParName
				Else
					lcAuxParName = Alltrim(Str(lnProcPar)) + ". " + lcParName
				Endif
			Endif

			*- monto o tooltip
			lcDescription = Iif(lcParName+":" $ lcDescription, "", Strtran(Strtran(lcDescription, Chr(13), ""), Chr(10), "") )
			lcToolTip = lcProc+"("+lcSignature+")" + ;
				iif(!Empty(lcDescription), Chr(13) + lcDescription, "") + ;
				iif(!Empty(lcAuxParName), Chr(13) + lcAuxParName, "")
		Endif


		*- apresento o tooltip
		If lnProcPar <= lnCountPars Or lnCountPars = 0
			This.ShowToolTipEditor( lcToolTip )

			*- apresento o erro indicando que a quantidade de parametros excedeu a quantidade definida na funcao
		Else
			This.ShowErrorWriteTime(1230, lcToolTip)
		Endif
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ faz a checagem do pode ser incluido na lista do IntelliSense de acordo com o que foi
	*/ definido no IntelliSense manager no checkbox "Incremental starded by [ ]"
	*/------------------------------------------------------------------------------------------------
	Protected Procedure ChkIncremental
		Lparameters plcLastWord, plcItem
		Local llReturn
		llReturn = .T.

		If This.IncrementalResult
			*- se contem o que foi digitado
			If This.cboSearch = "0"
				If Not Lower(plcLastWord) $ Lower(plcItem)
					llReturn = .F.
				Endif

				*- se inicia pelo o que foi digitado
			Else
				If Not Lower(plcLastWord) $ Lower(Substr(plcItem,1,Len(plcLastWord)))
					llReturn = .F.
				Endif
			Endif
		Endif

		Return llReturn
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Apresento tooltip para o editor corrente conforme a posição dentro do texto.
	*/ NOTE: O ToolTip usado na lista do IntelliSense é outro.
	*/------------------------------------------------------------------------------------------------
	Procedure ShowToolTipEditor
		Lparameters plcText, plcTitle, plnIcon, pllBallonTip, pllCloseButton
		Local lnLeft, lnTop

		Set Console Off

		*- posicao do cursor na janela do editor
		This.IntelliSense.GetCursorPos(@lnTop,@lnLeft)
		lnTop  = lnTop  + _Screen.Top  + Fontmetric(1, This.EditorFontName, This.EditorFontSize) + 1
		lnLeft = lnLeft + _Screen.Left + 2

		*- apresento o tooltip na posição correta
		With _Screen.FoxCodePlus.EditorToolTip
			*- configuro o tooltip
			.HWnd = _Screen.HWnd
			.Text = plcText
			.Title = Evl(plcTitle,"")
			.Icon = Evl(plnIcon,0)
			.BalloonTip = pllBallonTip
			.CloseButton = pllCloseButton
			.MousePos = .F.

			*- monto e apresento o tooltip nas coordenadas indicadas
			.Show(lnTop , lnLeft)
		Endwith
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Functions returns whether current editor position is at a location within open quote so
	*/ that it will be part of a string when close quote is added.
	*/------------------------------------------------------------------------------------------------
	Protected Procedure IsInQuotes
		Lparameters plcText
		Local i, lcChar, lInQuote, lcQuoteChar

		Set Console Off

		For i = 1 To Len(plcText)
			lcChar = Substr(plcText,m.i,1)
			If !lInQuote
				If Inlist(lcChar,'"',"'","[")
					lInQuote = .T.
					lcQuoteChar = lcChar
				Endif
			Else
				If (lcQuoteChar="[" And lcChar="]") Or (lcChar==lcQuoteChar And lcQuoteChar#"[")
					lInQuote = .F.
				Endif
			Endif
		Endfor
		Return lInQuote
	Endfunc


	*/------------------------------------------------------------------------------------------------
	*/ Functions returns whether current editor position is at a location within Text...EndText
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetTextEndText
		Lparameters plcText
		Local lcThisCode, lnLines, lnx, lcWord1, lnLineStartText, lnLineEndText, llReturn
		Local Array laCodePrg[1]

		Set Console Off

		lcThisCode = This.GetText(0,130000)
		lnLines = Alines(laCodePrg, lcThisCode)
		lnLineEndText = 0
		This.IsSqlIntelliSense = .F.

		*- se a linha q estou posicionado for a declaracao do text permito o IntelliSense
		If Lower(Getwordnum(plcText,1)) == "text"
			Return .F.
		Endif

		*- verifico onde fecha ou abre o text...endtext
		For lnx = 1 To This.CursorLine
			lnLineStartText = This.CursorLine-lnx
			If lnLineStartText <= 0 Or This.IsProc(laCodePrg[lnx]) Or lnLineStartText > Alen(laCodePrg,1)
				Return .F.
			Endif

			lcWord1 = Lower(Getwordnum(laCodePrg[lnLineStartText], 1))

			*- identifiquei q estou dentro de um text...endtext
			If lcWord1 == "text"
				llReturn = .T.
				This.TextEndBlock = ""

				*- se estou dentro dos delimitadores permito a abertura do IntelliSense com conteudo do vfp
				For lnx = 1 To Len(plcText)
					*- fora dos delimiters
					If Substr(plcText,Len(plcText)-lnx+1,3) == "> >"
						llReturn = .T.
						Exit
					Else
						*- dentro dos delimiters
						If Substr(plcText,Len(plcText)-lnx+1,3) == "< <" Or Substr(plcText,Len(plcText)-lnx+1,2) == "<<"
							llReturn = .F.
							Exit
						Endif
					Endif
				Endfor

				If llReturn
					*- se cheguei aqui é pq encontrei um bloco TEXT...ENDTEXT porem
					*- agora descubro onde o mesmo foi fechado e capturo todo o bloco
					lnLineEndText = Ascan(laCodePrg, "endtext", This.CursorLine, -1, 1, 15)
					If lnLineEndText > 0
						This.TextEndBlock = This.GetText(lnLineStartText, lnLineEndText-2)
						This.TextEndBlock = Strtran(Strtran(This.TextEndBlock, Chr(13)+Chr(10), " "), Chr(9), " ")
					Endif

					*- verifico se é uma instrucao SQL (somente para database conectado ex: MS SQL Server)
					If	( Inlist(Lower(Getwordnum(This.TextEndBlock,1)), "select", "update") ) Or ;
							( Lower(Getwordnum(This.TextEndBlock,1)) == "delete" And Lower(Getwordnum(This.TextEndBlock,2)) == "from" ) Or ;
							( Lower(Getwordnum(This.TextEndBlock,1)) == "insert" And Lower(Getwordnum(This.TextEndBlock,2)) == "into" )

						This.IsSqlIntelliSense = .T.
					Endif
				Endif

				Return llReturn
			Else
				If Substr(lcWord1,1,5) == "endte"
					Return llReturn
				Endif
			Endif
		Endfor

		Return llReturn
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ Retorna verdadeiro se a linha do texto do prg passado como parametro, indica que é a criacao
	*/ de uma procedure, metodo ou evento
	*/------------------------------------------------------------------------------------------------
	Protected Procedure IsProc
		Lparameters plcTextLine
		Local llReturn

		llReturn = ( Getwordcount(plcTextLine) >= 2 And Inlist(Lower(Substr(Getwordnum(plcTextLine,1),1,4)),"proc","func") ) Or ;
			( Getwordcount(plcTextLine) >= 3 And Inlist(Lower(Substr(Getwordnum(plcTextLine,1),1,4)),"hidd","prot") And Inlist(Lower(Substr(Getwordnum(plcTextLine,2),1,4)),"proc","func") )

		Return llReturn
	Endproc



	*/------------------------------------------------------------------------------------------------
	*/ Configuro o ambiente do VFP conforme parametro.
	*/ plnMode = 1 seta novos atribuito a IDE
	*/			 0 retorna os atributos configurados na IDE pelo desenvolvedor
	*/------------------------------------------------------------------------------------------------
	Procedure SetFoxcodePlusEnvironment
		Lparameters plnMode

		Set Console Off

		If plnMode =  1
			This.Environment[ 1] = Set("TALK")
			This.Environment[ 2] = Set("NOTIFY",2)
			This.Environment[ 3] = Set("ESCAPE")
			This.Environment[ 4] = Set("EXCLUSIVE")
			This.Environment[ 5] = Set("UDFPARMS")
			This.Environment[ 6] = Set("EXACT")
			This.Environment[ 7] = Set("MESSAGE",1)
			This.Environment[ 8] = 0		&&sys(2030)
			This.Environment[ 9] = _Tally

			Set Talk Off
			Set Notify Cursor Off
			Set Escape Off
			Set Exclusive Off
			Set Udfparms To Value
			Set Exact On
			*sys(2030,0)
		Else
			If This.Environment[1]="ON"
				Set Talk On
			Endif
			If This.Environment[2]="ON"
				Set Notify Cursor On
			Endif
			If This.Environment[3]="ON"
				Set Escape On
			Endif
			If This.Environment[4]="ON"
				Set Exclusive On
			Endif
			If This.Environment[5]="REFERENCE"
				Set Udfparms To Reference
			Endif
			If This.Environment[6]="OFF"
				Set Exact Off
			Endif
			*sys(2030,int(val(this.Environment[8])))
			_Tally = This.Environment[9]
		Endif
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Procura a informação no array de itens do IntelliSense.
	*/ pesquisa por categoria (property, method, etc)
	*/------------------------------------------------------------------------------------------------
	Protected Procedure SeekItem
		Lparameters plcItem, plnImage
		Local lnRowFound, lnStarElement

		Set Console Off

		lnStarElement = 1
		lnRowFound = 0

		Do While .T.

			*- procuro na 1a. coluna
			lnRowFound = Ascan(This.Items, plcItem, lnStarElement, -1, 1, 15)

			If lnRowFound <= 0
				Exit
			Else
				*- verifico se o tipo do item é o correto
				*- caso contrario, volto e faço uma nova pesquisa iniciando na proxima linha
				Do Case
						*- class
					Case plnImage = 1 And This.Items[lnRowFound,2] = 1
						Exit

						*- propriedades
					Case Inlist(plnImage,7,8,9,10) And Inlist(This.Items[lnRowFound,2],7,8,9,10)
						Exit

						*- metodos
					Case Inlist(plnImage,3,14,15,4,5,6) And Inlist(This.Items[lnRowFound,2],3,14,15,4,5,6)
						Exit

						*- variable
					Case plnImage = 11 And This.Items[lnRowFound,2] = 11
						Exit

						*- visual control
					Case plnImage = 13 And This.Items[lnRowFound,2] = 13
						Exit

						*- procedure (procedural)
					Case plnImage = 19 And This.Items[lnRowFound,2] = 19
						Exit

						*- object locked (_screen, this, activex or object instantied)
					Case plnImage = 20 And This.Items[lnRowFound,2] = 20
						Exit

					Otherwise
						lnStarElement = lnRowFound + 1
						If lnStarElement > Alen(This.Items,1)
							lnRowFound = 0
							Exit
						Endif
				Endcase
			Endif

		Enddo

		Return lnRowFound
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ plcFile deve ser passado como referencia
	*/ Retorna .T. se o arquivo foi encontrado
	*/------------------------------------------------------------------------------------------------
	Protected Procedure GetFilePath
		Lparameters plcFile
		Local llReturn, lfName

		*- verifico se o arquivo existe.
		*- OBS: A funcao nativa File() verifica se o arquivo existe no diretorio corrente,
		*- caso nao encontre o arquivo faz a busca nas pastas informadas no SET PATH
		If File(plcFile)
			*- se informei o arquivo sem a pasta, entao faço a verificação pelo
			*- SET PATH usando a função LOCFILE() para capturar o nome completo com a pasta.
			*- isso assegura a abertura do arquivo quando esta fora do diretorio corrente.
			If Not "\" $ plcFile
				plcFile = Locfile(plcFile,Justext(plcFile),"FoxcodePlus Search")
			Endif
			llReturn = .T.
		Else
			If This.chkChkSetPCls == "1" Then
				*- 09/03/2016 - Marcelo Junior (mInternauta)
				*- Forçar um Getfile se o não achar o arquivo para que o usuário informe o caminho do arquivo
				lfName  = Alltrim(Justfname(plcFile))
				If !Empty(lfName) Then
					plcFile = Getfile(Justext(plcFile), "Search for " + Justfname(plcFile), "Open", 0, "Foxcodeplus Search")
					llReturn = File(plcFile)
				Else
					llReturn = .F.
				Endif
			Else
				llReturn = .F.
			Endif
		Endif

		*- faço isso para garantir abertura de arquivos com large name.
		plcFile = ["]+This.ClearQuotes(plcFile)+["]

		Return llReturn
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ Abre o List Error
	*/------------------------------------------------------------------------------------------------
	Procedure showerrorlist
		If Vartype(This.FormErrorList) <> "O"
			This.FormErrorList = Newobject("ErrorList","FoxCodePlusIntelliSense.vcx")
		Endif
		This.FormErrorList.Show()
	Endproc


	*/------------------------------------------------------------------------------------------------
	*/ ao destruir esse objeto
	*/------------------------------------------------------------------------------------------------
	Protected Procedure Destroy
		On Key Label "."
		On Key Label "="
		Try
			Erase (This.Tmpfile)
		Catch
		Endtry
		Set Message To ""
	Endproc



	*------------------------------------------------------------------------------*
	* funcao arrayToString
	* recebe um array por parametro (@) e devolve todo o conteudo em String
	* exemplo:	AERROR(meuArray)
	*			arrayToString(@meuArray)
	*------------------------------------------------------------------------------*
	Function arrayToString
		Lparameters laArray

		Local lcRetorno
		If Type("Alen(laArray, 1)") = "U"
			Dimension laArray[1,1]
		Endif

		lcRetorno = ""
		Try
			Local lnCol, lnLin, lnContaCol, lnContaLin
			lnCol = Alen( laArray, 2)
			lnLin = Alen( laArray, 1)

			For lnContaLin = 1 To lnLin
				For lnContaCol = 1 To lnCol  && Exibe todos os elementos da matriz
					lcRetorno = lcRetorno + "(" + Alltrim(Str(lnContaLin)) + ", " + Alltrim(Str(lnContaCol)) + ") " + Transform(laArray[lnContaLin, lnContaCol]) + Chr(13) + Chr(10)
				Endfor
			Endfor
		Catch
		Endtry

		Return lcRetorno
	Endfunc

	*/------------------------------------------------------------------------------------------------
	*/ Se ocorrer erro no FoxCodePlus...
	*/------------------------------------------------------------------------------------------------
	Procedure Error
		Lparameters plnError, plcMethod, plnLine

		Local lcMsgPilha, lcAerror

		lcAerror = ""
		If Aerror(gaerror) > 0
			lcAerror = "** AERROR() **" + CRLF + This.arrayToString(@gaerror)
		Endif

		Astackinfo(PILHA_EXECUCAO)
		lcMsgPilha = "** Astackinfo() **" + CRLF + This.arrayToString(@PILHA_EXECUCAO)

		Set Memowidth To 8000

		Local lcError, loVersion, lcLevelError, lnx
		Local Array laStackInfo[1,2]

		Set Console Off

		*- desconsidero a apresentação desse erro, pois nao é um erro e sim um bug do vfp,
		*- até porque nao existe codigo nesse metodo"
		If Lower(plcMethod) == "foxcodeplusintellisense.tabs.tabpage1.items.mousedown"
			&&- do nothing
		Else
			Try
				*- objeto que contem as informacoes da versao do foxcodeplus
				*loVersion = Createobject("xfcpVersion")

				*- acho o level correto do erro
				lnx = 1
				lcLevelError = ""
				Do While Len(Sys(16,lnx)) <> 0
					If ".ERROR " $ Sys(16,lnx)
						lcLevelError = Sys(16,lnx-1)
						Exit
					Endif
					lnx = lnx + 1
				Enddo


				*- monto a messagem de erro

				lcError = "**-------------------------------------------" + CRLF + ;
					"A FoxcodePlus error has occured. Press CTRL+C to copy this error or send the file foxcodeplus.err to the author. " + CRLF + ;
					"" + CRLF + ;
					"Line contents: " + Message(1) + CRLF + ;
					"OS version: "+ Os(1) + CRLF + ;
					"Wontop: "+ Wontop() + CRLF + ;
					"Localization: " + Iif(Empty(lcLevelError), "UNKNOWN", lcLevelError) + CRLF + ;
					"Method: " + plcMethod + CRLF + ;
					"Line: " + Transform(plnLine) + CRLF + ;
					"Error message: " + Message() + CRLF + ;
					"Error number: " + Transform(plnError) + CRLF + ;
					"FoxcodePlus version: "+ ""  + CRLF + ;
					lcAerror + CRLF + lcMsgPilha + CRLF + CRLF

				loVersion = .Null.

				Strtofile(Ttoc(Datetime()) + CRLF + lcError + CRLF + CRLF, "foxcodeplus.err", 1)
				*Messagebox(lcError,16)
			Endtry
		Endif
	Endproc

Enddefine




*/------------------------------------------------------------------------------------------------
*/ Classe para controlar a versao do FoxcodePlus
*/------------------------------------------------------------------------------------------------
Define Class xfcpVersion As Custom
	Version = "Beta 3.13.2"
	Datetime = Ttoc( Iif(File(Addbs(Home(1)) + "foxcodeplus.app"), Fdate(Addbs(Home(1))+"foxcodeplus.app",1), "") )
	Author = "Rodrigo Duarte Bruscain"
	CountryAndCity = "kitchener ON - Canada"
	Url = "http://vfpx.codeplex.com/wikipage?title=FoxcodePlus"
	Email = "bruscain@hotmail.com"

	Protected Procedure Version_ASSIGN(newvalue)
		Error 1740, "Number"
	Endproc

	Protected Procedure DateTime_ASSIGN(newvalue)
		Error 1740, "DateTime"
	Endproc

	Protected Procedure Author_ASSIGN(newvalue)
		Error 1740, "Author"
	Endproc

	Protected Procedure CountryAndCity_ASSIGN(newvalue)
		Error 1740, "CountryAndCity"
	Endproc

	Protected Procedure Url_ASSIGN(newvalue)
		Error 1740, "Url"
	Endproc

	Protected Procedure Email_ASSIGN(newvalue)
		Error 1740, "Email"
	Endproc
Enddefine


*/------------------------------------------------------------------------------------------------
*/ Classe usada para obter o objeto OleControl dentro de um DEFINE CLASS adicionado por ADD OBJET
*/------------------------------------------------------------------------------------------------
Define Class TempOleClass As Form
	Procedure Init
		Lparameters plcOleControl
		Set Console Off
		If Not Empty(plcOleControl)
			This.AddObject("xOleControl","olecontrol",plcOleControl)
		Endif
	Endproc
Enddefine


*/------------------------------------------------------------------------------------------------
*/ Classe usada para obter as PMEs da objeto instanciando em write-time pelas funcoes
*/ CreateObject(), CreateObjectEx() e NewObject()
*/------------------------------------------------------------------------------------------------
Define Class TempObjWt As Custom
	Procedure Init
		Lparameters plcCommand
		Set Console Off
		If Not Empty(plcCommand)
			This.AddProperty("xObj",Evaluate(plcCommand))
		Endif
	Endproc
Enddefine





*- functions created to avoid errors
Procedure _EdGetPos
Procedure _EDPOSINVI
Procedure _EdGetChar
Procedure _wSelect
Procedure _wontop
Procedure _EdGetEnv
Procedure _EdSetEnv
Procedure _WZOOM
Procedure _wtitle
Procedure _EdGetLNum
Procedure _EdGetLPos
Procedure _EdGetStr
Procedure _EdSelect
Procedure _EdDelete
Procedure _EdInsert
Procedure _EdSetPos
Procedure _WHTOHWND
Procedure _EDSTOSEL
Procedure _EdCopy
Procedure _EdPaste


	*- just for test
	*procedure mc(plcFileName)
	*	modify command &plcFileName
	*endproc


	*unbindevents(0, WM_KEYUP)
	*sys(2030,1)
	*set step on
