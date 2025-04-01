' API_Request.bas - Module VBA pour interagir avec une API Web
' Exemple : Récupération des taux de change depuis une API

Attribute VB_Name = "API_Request"

Sub GetExchangeRates()
    Dim http As Object
    Dim JSON As Object
    Dim url As String
    Dim ws As Worksheet
    Dim i As Integer
    
    ' Définir l'URL de l'API (Exemple : API de taux de change)
    url = "https://api.exchangerate-api.com/v4/latest/USD"
    
    ' Initialiser l'objet HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send
    
    ' Vérifier la réponse de l'API
    If http.Status <> 200 Then
        MsgBox "Erreur de récupération des données !", vbExclamation, "Erreur API"
        Exit Sub
    End If
    
    ' Parser la réponse JSON
    Set JSON = JsonConverter.ParseJson(http.responseText)
    
    ' Sélectionner la feuille active
    Set ws = ActiveSheet
    
    ' Insérer les données dans la feuille
    ws.Cells(1, 1).Value = "Devise"
    ws.Cells(1, 2).Value = "Taux de change"
    
    i = 2
    Dim key As Variant
    For Each key In JSON("rates")
        ws.Cells(i, 1).Value = key
        ws.Cells(i, 2).Value = JSON("rates")(key)
        i = i + 1
    Next key
    
    MsgBox "Données mises à jour avec succès !", vbInformation, "Succès"
End Sub
