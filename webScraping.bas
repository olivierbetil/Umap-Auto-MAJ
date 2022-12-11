Attribute VB_Name = "webScraping"
'****************************************************************************************************
'@file: webScraping.bas
'@brief: Librairie pour utiliser Selenium dans le cadre du projet
'@author: Betil Olivier
'@version: v1.0
'@date: 06/12/2022
'****************************************************************************************************

Option Explicit

'@brief fonction qui met à jour la carte Umap présente dans le tableur avec le fichier geoJson crée
'
'@param void
'@return void
Public Sub umapUpdate()

    'création d'une instance de Chrome (headless = invisible)
    Dim bot As New WebDriver
    bot.AddArgument "--headless"
    bot.Start "chrome", "https://www.google.com"
    
    'recherche l'adresse présente à la case J3
    bot.Get (ActiveSheet.Cells(3, 10))
    
    'active le mode modification
    bot.FindElementByXPath("/html/body/div[1]/div[2]/div[2]/div[1]/a").Click
    
    'supprime le calque
    bot.FindElementByXPath("/html/body/div[1]/div[2]/div[1]/div[9]/a").Click
    bot.FindElementByXPath("/html/body/div[1]/div[2]/div[1]/div[9]/div/ul/li/i[5]").Click
    bot.SwitchToAlert.Accept
    bot.Wait (500)
    
    'importe le fichier en geojson sur Umap
    bot.FindElementByXPath("/html/body/div[1]/div[2]/div[2]/div[3]/ul/li[1]/a").Click
    On Error Resume Next
    bot.FindElementByXPath("/html/body/div[1]/div[4]/div[1]/div/div/div[2]/input").SendKeys Environ$("USERPROFILE") & "\Downloads" & "\" & ActiveSheet.Name & ".geojson"
    bot.FindElementByXPath("/html/body/div[1]/div[4]/div[1]/div/div/div[2]/input").SendKeys Environ$("USERPROFILE") & "\Téléchargements" & "\" & ActiveSheet.Name & ".geojson"
    On Error GoTo 0
    bot.Wait (100)
    
    'enregistre les changements et ferme l'instance de Chrome
    bot.FindElementByXPath("/html/body/div[1]/div[4]/div[1]/div/div/input[2]").Click
    bot.FindElementByXPath("/html/body/div[1]/div[2]/div[7]/a[1]").Click
    bot.Wait (1000)
    bot.Close
End Sub

'@brief récupère les coords des villes où c'est nécessaire
'
'@param nbLignes Entier le nombre de villes/elements présents dans le tableau
'@return void
Public Sub recupCoord(nbLignes As Integer)
    Sheets("Tabelle1").Select
    
    'crée une nouvelle instance de Chrome
    Dim bot As New WebDriver
    bot.AddArgument "--headless"
    bot.Start "chrome", "https://www.google.com"
    
    'recherche vers le site web "coordonnees-gps.fr"
    bot.Get ("https://www.coordonnees-gps.fr/")
    bot.Wait (1000)
    
    'accepte les cookies
    bot.FindElementByXPath("/html/body/div[3]/nav/div/div/div/a[1]").Click
    bot.Wait (1000)
    
    'boucle pour aller chercher et sauvgarder les coords
    Dim compteurCoord As Integer
    compteurCoord = 6
    Do While compteurCoord <= nbLignes
        If IsEmpty(ActiveSheet.Cells(compteurCoord, 5).Value) = True Then
            bot.FindElementByXPath("/html/body/div[1]/div[2]/div[2]/div[1]/form[1]/div[1]/div/input").Clear
            bot.FindElementByXPath("/html/body/div[1]/div[2]/div[2]/div[1]/form[1]/div[1]/div/input").SendKeys ActiveSheet.Cells(compteurCoord, 4).Value
            bot.FindElementByXPath("/html/body/div[1]/div[2]/div[2]/div[1]/form[1]/div[2]/div/button").Click
            bot.Wait (200)
            ActiveSheet.Cells(compteurCoord, 5) = bot.FindElementByXPath("/html/body/div[1]/div[2]/div[2]/div[1]/form[2]/div[1]/div/input").Attribute("value")
            ActiveSheet.Cells(compteurCoord, 6) = bot.FindElementByXPath("/html/body/div[1]/div[2]/div[2]/div[1]/form[2]/div[2]/div/input").Attribute("value")
        End If
        compteurCoord = compteurCoord + 1
    Loop
    
    'ferme l'instance de Chrome
    bot.Close
End Sub

