#include <WinAPIFiles.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <WindowsConstants.au3>
#include <Date.au3>
#include <File.au3>
#include <Math.au3>


#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 615, 438, 192, 124)
$List1 = GUICtrlCreateList("", 8, 8, 345, 110, BitOR($LBS_NOTIFY, $LBS_NOSEL, $WS_VSCROLL, $WS_BORDER))
$List2 = GUICtrlCreateList("", 16, 144, 337, 292, BitOR($LBS_NOTIFY, $WS_VSCROLL, $WS_BORDER))
$Button1 = GUICtrlCreateButton("RAZ", 384, 16, 153, 49)
$Button2 = GUICtrlCreateButton("Quitter", 384, 72, 153, 49)
#EndRegion ### END Koda GUI section ###


Global $lign = 0

list2("Application lancée le " & _Now())


Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
Local $sFilePath = @ScriptDir & "\config.ini"
Local $source = IniRead($sFilePath, "General", "Source", @ScriptDir & "\SOURCE\")
Local $dest = IniRead($sFilePath, "General", "Destination", @ScriptDir & "\HPR2HM\DEST\")
Local $copie = IniRead($sFilePath, "General", "DestinationCopie", @ScriptDir & "\SOURCE\")
Local $deplace = IniRead($sFilePath, "General", "DestinationDeplace", @ScriptDir & "\HPR2HM\DEST\")
Local $OBX_cons = StringSplit(IniRead($sFilePath, "General", "OBX_a_conserver", ""), "|")
Local $OBX_cop = StringSplit(IniRead($sFilePath, "General", "OBX_a_copier", ""), "|")
Local $OBX_dep = StringSplit(IniRead($sFilePath, "General", "OBX_a_deplacer", ""), "|")
Local $delai = Number(IniRead($sFilePath, "General", "Delai", "1000"))


if $source = $dest Then
	MsgBox(8192, "Erreur", "Répertoire source et destination identique !")
	exit
EndIf

if FileExists($source) = 0 Then
	MsgBox(8192, "Erreur", "Répertoire source non existant / non accessible")
	exit
EndIf

if FileExists($dest) = 0 Then
	MsgBox(8192, "Erreur", "Répertoire destination non existant / non accessible")
	exit
EndIf

if FileExists($copie) = 0 Then
	MsgBox(8192, "Erreur", "Répertoire copie non existant / non accessible")
	exit
EndIf

if FileExists($deplace) = 0 Then
	MsgBox(8192, "Erreur", "Répertoire deplace non existant / non accessible")
	exit
EndIf

GUICtrlSetData($List1, "Source : " & $source)
GUICtrlSetData($List1, "Destination : " & $dest)
GUICtrlSetData($List1, "Copie : " & $copie)
GUICtrlSetData($List1, "Deplace : " & $deplace)

Local $txt = "OBX à conserver :"
for $k = 1 to $OBX_cons[0]
	$txt = $txt & " " & $OBX_cons[$k]
Next
Local $OBX_cons_t = False
if $txt = "OBX à conserver : " Then $OBX_cons_t = True
GUICtrlSetData($List1, $txt)


Local $txt = "OBX à copier : "
for $k = 1 to $OBX_cop[0]
	$txt = $txt & " " & $OBX_cop[$k]
Next
GUICtrlSetData($List1, $txt)

Local $txt = "OBX à deplacer : "
for $k = 1 to $OBX_dep[0]
	$txt = $txt & " " & $OBX_dep[$k]
Next


GUICtrlSetData($List1, $txt)
GUICtrlSetData($List1, "Délai : " & $delai & " ms")





GUISetState(@SW_SHOW)

$hTimer = TimerInit()


While 1

	if TimerDiff($hTimer) >= $delai Then
;~ 		list2("Lecture du répertoire le " & _Now())
		$hTimer = TimerInit()
		$asource = _FileListToArray($source, "*.OK", 1, 1) ; Lecture des fichiers ok du répertoire source
		if @error = 0 Then ; Si pas d'erreur
			for $i = 1 to $asource[0] ; Pour chaque fichier ok
				if FileExists($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR") AND FileExists($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".OK") AND FileExists($source & "T_" & StringSplit(_PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3], "-")[3] & ".pdf") Then ; Si les fichiers ok, hprim et pdf existent
					list2("Traitement du fichier : " & $asource[$i])
					Local $hFileOpen2 = FileOpen($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR", $FO_READ) ; $hFileOpen2 = Fichier source
					Local $hFileOpen3 = FileOpen($dest & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR", $FO_OVERWRITE) ; $hFileOpen3 = Fichier destination
					$acopier = False
					$adeplacer = False
					$txtobr=""
					$txtpre=""
					$codepre=""


							While 1 = 1
						$txt = FileReadLine($hFileOpen2) ; Lecture de la ligne du fichier source
						if @error = -1 then ExitLoop ; Si on est à la fin, on quitte la bouche
						if StringSplit($txt, "|")[1] = "OBX" Then ; Si le début est OBX
						   if $txtobr<>"" Then ; Si le buffer $txtobr est plein, on l'écrit et on le supprime
						   FileWriteLine($hFileOpen3, $txtobr) ; On écrit l'OBR
						   $txtobr=""
						   EndIf

						   for $j = 1 to $OBX_cons[0] ; On vérifie sur le code OBX est à conserver
								if StringSplit(StringSplit($txt, "|")[4], "~")[1] = $OBX_cons[$j] OR $OBX_cons_t Then ; Si oui, on recopie la ligne dans le fichier source
								   									FileWriteLine($hFileOpen3, $txt)
								EndIf
							Next

							for $j = 1 to $OBX_cop[0] ; On vérifie sur le code OBX doit permettre de copier le fichier source
								if StringSplit(StringSplit($txt, "|")[4], "~")[1] = $OBX_cop[$j] AND $OBX_Cop[$j]<> "" Then
									$acopier = True
									list2($asource[$i]&" à copier car OBX "&$OBX_cop[$j]&" trouvé")
								EndIf
							Next

							for $j = 1 to $OBX_dep[0] ; On vérifie sur le code OBX doit permettre de déplacer le fichier source
								if StringSplit(StringSplit($txt, "|")[4], "~")[1] = $OBX_dep[$j] AND $OBX_dep[$j]<>"" Then
									$adeplacer = True
									list2($asource[$i]&" à déplacer car OBX "&$OBX_dep[$j]&" trouvé")
								EndIf
							 Next
$codepre="OBX"

						Elseif StringSplit($txt, "|")[1] = "C" Then ; Si le début est C ==> On suprime la ligne
$codepre="C"

						Elseif StringSplit($txt, "|")[1] = "OBR" Then ; Si le début est OBR, on modifie l'entete d'affichage
						$txtpre=$txt
							Local $atemp = StringSplit($txt, "|")
							Local $atemp2 = StringSplit($atemp[5], "^")


							$txt2="CR_SGL~COMPTE RENDU PDF^"
;~ 							for $j = 1 to $atemp2[0]
;~ 							   $tt=StringSplit($atemp2[$j],"~")
;~ 							   if @error<>1 then $txt2 = $txt2 & $tt[2] & "^"

;~ 							 Next
							 $atemp[5]=StringLeft($txt2, StringLen($txt2) - 1)

							$txt = ""
							for $j = 1 to $atemp[0]
								$txt = $txt & $atemp[$j] & "|"
							 Next

							$txtobr=StringLeft($txt, StringLen($txt) - 1)
$codepre="OBR"

						Elseif StringSplit($txt, "|")[1] = "A" Then ; Si le début est A

						if $codepre="OBR" Then ; Si le code précédent est OBR
						   $txt=$txtpre&StringRight($txt, StringLen($txt) - 2)
						   $txtpre=$txt

						   Local $atemp = StringSplit($txt, "|")
							Local $atemp2 = StringSplit($atemp[5], "^")


							$txt2="CR_SGL~COMPTE RENDU PDF^"
;~ 							for $j = 1 to $atemp2[0]
;~ 								$tt=StringSplit($atemp2[$j],"~")
;~ 							   if @error<>1 then $txt2 = $txt2 & $tt[2] & "^"
;~ 							 Next
							 $atemp[5]=StringLeft($txt2, StringLen($txt2) - 1)

							$txt = ""
							for $j = 1 to $atemp[0]
								$txt = $txt & $atemp[$j] & "|"
							Next
						   $txtobr=StringLeft($txt, StringLen($txt) - 1)
						Else ; Sinon on supprime la ligne
						   EndIf





						Else ; Si le code est ni OBX, ni C, ni A, ni OBR, on recopie la ligne telle quelle
							FileWriteLine($hFileOpen3, $txt)
							$codepre=""
						EndIf



						if @error = -1 then ExitLoop ; Si on est à la fin, on quitte la bouche
					WEnd

					FileClose($hFileOpen2)
					FileClose($hFileOpen3)

					if $acopier Then ; Si le fichier source est à copier
						FileCopy($source & "T_" & StringSplit(_PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3], "-")[3] & ".pdf", $copie & "T_" & StringSplit(_PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3], "-")[3] & ".pdf", 1)
						FileCopy($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR", $copie & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR", 1)
						FileCopy($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".OK", $copie & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".OK", 1)
					EndIf

					if $adeplacer Then ; Si le fichier source est à déplacer
						Filemove($source & "T_" & StringSplit(_PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3], "-")[3] & ".pdf", $deplace & "T_" & StringSplit(_PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3], "-")[3] & ".pdf",1)
						FileMove($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR", $deplace & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR", 1)
						FileDelete($dest & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR")
						FileMove($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".OK", $deplace & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".OK", 1)


					Else ; Sinon on déplace les fichiers pdf et ok et on supprime le fichier source
						FileMove($source & "T_" & StringSplit(_PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3], "-")[3] & ".pdf", $dest & "T_" & StringSplit(_PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3], "-")[3] & ".pdf", 1)
						FileDelete($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".HPR")
						FileMove($source & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".OK", $dest & _PathSplit($asource[$i], $sDrive, $sDir, $sFileName, $sExtension)[3] & ".OK", 1)
					EndIf
				EndIf
			Next
		EndIf
	EndIf

	$nMsg = GUIGetMsg()
	Switch $nMsg

		Case $Button2
			Exit
		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch

WEnd


Func list2($txt)
	Local $hFileOpen = FileOpen("log.txt", $FO_APPEND)
	FileWrite($hFileOpen, _now()&" : "&$txt & @CRLF)
	FileClose($hFileOpen)
	if $lign = 500 Then
		GUICtrlSetData($List2, "")
		$lign = -1
	EndIf
	$lign = $lign + 1
	_GUICtrlListBox_SetTopIndex($List2, _GUICtrlListBox_GetListBoxInfo($List2) - 1)
	GUICtrlSetData($List2, $txt, -1)


EndFunc   ;==>list2


i = 0
While 1

	$nMsg = GUIGetMsg()
	Switch $nMsg

	case $Button1
	   GUICtrlSetData($List2, "")
	   		$lign = 0

		case $Button2
			Exit

		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch
WEnd
