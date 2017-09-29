'LIEN GITHUB: https://github.com/ValoremDev/Wind-Level3

'=======================================================================================================
'DEFINITION DES VARIABLES GLOBALES
'=======================================================================================================
Public DSCR_sculptage As Double             ' DSCR de sculptage
Public Gearing_sup As Boolean               ' Booléen de dépassement de Gearing Maximum
Public prod As Integer                      ' Scénario de production
Public Type_Amortissement As String         ' Dette sculptée ou K+I
Public TRI_Renta As Double                  ' TRI visé pour visée tarif
Public Cas_Scenario As Integer              ' Numéro de scénario actif
Public TRI_Actionnaire As Double            ' TRI Actionnaire calculé
Public debugBool As Boolean                 ' Booléen indiquant l'utilisation ou non du mode débug
Public debugLoopValue As Integer            ' Numéro de loop du Crawler
Public scenario As Integer                  ' Placeholder du scenario
Public CrawlerJson As Object                ' Objet ayant pour but de contenir le JSON du Crawler Debugger
Public delta_CAPEX As Object                ' Placeholder du range marge d'Input pour évolution dasn Debugger


' PRIVATE VARIABLES
Private Up As Boolean                       ' Décrit un passage au DESSUS de la zone de recherche de TRI
Private Down As Boolean                     ' Décrit un passage au DESSOUS de la zone de recherche de TRI
Private step_Tarif As Double                ' Step appliqué au tarif lors de la recherche du TRI


'=======================================================================================================
' DEFINITION DES CONSTANTES GLOBALES
' /!\ Noms des Constantes en MAJUSCULES /!\
'=======================================================================================================

' On fixe le pas avec lequel on fait varier le DSCR en cas de gearing trop élevé à 0.01 (standard)
Public Const STEP_DSCR_1 = 0.01
Public Const STEP_DSCR_2 = 0.001
Public Const STEP_DSCR_3 = 0.0001
Public Const STEP_DSCR_4 = 0.00001

' On fixe la précision sur le gearing en cas de gearing trop élevé à 0.001% (standard)
Public Const PRECISION_GEARING_1 = 0.001
Public Const PRECISION_GEARING_2 = 0.0001
Public Const PRECISION_GEARING_3 = 0.00001

' On fixe le Pas de base sur la recherche de Tarif pour TRI souhaité
Public Const STEP_TARIF_INITIAL = 2

' On fixe la précision souhaitée sur le TRI pour la recherche de tarif pour TRI souhaité
Public Const PRECISION_TRI = 0.00001

' On fixe une durée post copier coller pour éviter les crash excel qui peuvent découler d'un trop grand nombre d'opérations
Public Const MSECONDS = 0.000000011574


'=======================================================================================================
' DEFINITION DES ELEMENTS EXTERNES
' LIBRAIRIES ANNEXES SI BESOIN
'=======================================================================================================
'VBA-JSON : https://github.com/VBA-tools/VBA-JSON
'VBA-Dictionary : https://github.com/VBA-tools/VBA-Dictionary
'Permet d'éviter Windows Scripting Runtime

Public Const xRef As String = "C:\WINDOWS\System32\scrrun.dll"  'Référence au Microsoft Scripting Runtime

