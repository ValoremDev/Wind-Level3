Sub crawler(ByVal BoxStep, ByVal Boxiter)
    '---------------------------------------------------
    ' Crawl sur le Xème scénario afin de tester le modèle et les différentes configurations / tournicoti
    '---------------------------------------------------
    
    Dim CurrDir As String: CurrDir = ActiveWorkbook.Path
    
    Dim logFile As Object
    Dim fso As Object

    debugBool = True
    scenario = Range("Cas_Scenario").Value
    Set delta_CAPEX = Range("delta_CAPEX").Offset(0, 1 - scenario)
    initial = delta_CAPEX.Value
    
    step = BoxStep
    
    Set CrawlerJson = JsonConverter.ParseJson("{}")
    
    Debug.Print "-------------------------"
    Debug.Print "Running DEBUG MODE"
    Debug.Print "-------------------------"
    
    For i = 1 To Boxiter / 2
        debugLoopValue = i
        delta_CAPEX.Value = delta_CAPEX.Value + step
        Call Sortie_BP

    Next
    
    delta_CAPEX.Value = initial
    
    For i = Boxiter / 2 + 1 To Boxiter
        debugLoopValue = i
        delta_CAPEX.Value = delta_CAPEX.Value - step
        Call Sortie_BP

    Next
    
    delta_CAPEX.Value = initial

    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As String: file = CurrDir & "\" & Format(Date, "yyyy-mm-dd") & " - Log.txt"
    
    Set logFile = UseOverwriteFile(file)
    
    Dim scenarioName As String: scenarioName = Range("scenario_Name").Offset(0, 1 - scenario).Value
    
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



