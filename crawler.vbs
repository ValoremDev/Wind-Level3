Sub crawler(ByVal BoxStep, ByVal Boxiter)
    '---------------------------------------------------
    ' Crawl sur le Xème scénario afin de tester le modèle et les différentes configurations
    ' :BoxStep: Provient du Form Debug/Test -> Step up / down à appliquer à chaque itération sur la Marge
    ' :Boxiter: Provient du Form Debug/Test -> Nombre d'itérations à effectuer. Répartie 50/50 entre montant et descendant
    '---------------------------------------------------
 
    Set fso = CreateObject("Scripting.FileSystemObject")   
    Dim CurrDir As String: CurrDir = ActiveWorkbook.Path
    
    Dim logFile As Object
    Dim fso As Object
    Set CrawlerJson = JsonConverter.ParseJson("{}")

    Dim file As String: file = CurrDir & "\" & Format(Date, "yyyy-mm-dd") & " - Log.txt"
    Set logFile = UseOverwriteFile(file)

    scenario = Range("Cas_Scenario").Value
    Set delta_CAPEX = Range("delta_CAPEX").Offset(0, 1 - scenario)
    Dim scenarioName As String: scenarioName = Range("scenario_Name").Offset(0, 1 - scenario).Value

    Debug.Print "-------------------------"
    Debug.Print "Running DEBUG modèle"
    Debug.Print "-------------------------"

    step = BoxStep
    debugBool = True
    initial = delta_CAPEX.Value

    For i = 1 To Boxiter / 2
        debugLoopValue = i
        delta_CAPEX.Value = delta_CAPEX.Value + step
        Call Sortie_BP
    Next
    
    ' Réinitialisation du delta_CAPEX pour explorer à la baisse
    delta_CAPEX.Value = initial
    
    For i = Boxiter / 2 + 1 To Boxiter
        debugLoopValue = i
        delta_CAPEX.Value = delta_CAPEX.Value - step
        Call Sortie_BP
    Next
    
    delta_CAPEX.Value = initial

    logFile.WriteLine scenarioName & " - " & Now & " - Logger"
    logFile.WriteLine "Exit Code - " & "Delta CAPEX - " & "DSCR Calc - " & "Gearing Calc"
    logFile.WriteLine "--------------------------------------------------------"
    logFile.WriteLine ""

    For Each strkey In CrawlerJson.Keys()
        logFile.WriteLine strkey & " - " & CrawlerJson(strkey)("Exit Code") & " - " & CrawlerJson(strkey)("Delta CAPEX") & " - " & CrawlerJson(strkey)("DSCR Calc") & " - " & CrawlerJson(strkey)("Gearing Calc")
    Next
    
    logFile.WriteLine ""
    
    debugBool = False

    Debug.Print "-------------------------"
    Debug.Print "ENDING DEBUG MODE"
    Debug.Print "-------------------------"
End Sub



