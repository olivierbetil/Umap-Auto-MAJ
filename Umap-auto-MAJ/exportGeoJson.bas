Attribute VB_Name = "exportGeoJson"
'****************************************************************************************************
'file: exportGeoJson.bas
'@brief: Programme pour mettre à jour le carte Umap présente dans l'excel selon le contenu de l'excel
'@author: Betil Olivier
'version: v1.0
'date: 06/12/2022
'****************************************************************************************************

Option Explicit

Sub GEOJSON()
    'on choisis le tableur principal
    Sheets("Tabelle1").Select
    
    'Fonction pour mettre à jour le driver pour Chrome si besoin
    With New SeleniumWebDriverUpdate
        .UpdateDriver (Chrome)
    End With
    
    'création d'un objet qui va représenter le fichier qu'on enregistrera (et on commence à écrire dedans)
    Dim objStream As ADODB.Stream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.WriteText "{" & vbLf
    objStream.WriteText "  ""type"": ""FeatureCollection""," & vbLf
    objStream.WriteText "  ""features"": [" & vbLf
    
    'récupere le nombre de données à gérer
    Dim intStop As Integer
    Dim nbLignes As Integer
    nbLignes = 6
    intStop = 0
    Do While intStop = 0
        If IsEmpty(ActiveSheet.Cells(nbLignes, 2).Value) = True Then
            intStop = 1
        End If
        nbLignes = nbLignes + 1
    Loop
    nbLignes = nbLignes - 2
    
    'si cases coordonnées vides, recherches des coords avec le nom de la ville/de l'élement
    intStop = 0
    recupCoord (nbLignes)
    
    'boucle principale qui va remplir l'objet selon les règles du geoJson
    Dim compteur As Integer
    compteur = 7
    Do While compteur <= nbLignes
        Dim lat1 As String, lng1 As String, ville As String
        ville = ""
        ville = ville + ActiveSheet.Cells(compteur, 4).Value
        ville = ville + " "
        ville = ville + CStr(ActiveSheet.Cells(compteur, 3).Value)
        lat1 = ActiveSheet.Cells(compteur, 5).Value
        lng1 = ActiveSheet.Cells(compteur, 6).Value
        objStream.WriteText "    {" & vbLf
        objStream.WriteText "      ""type"": ""Feature""," & vbLf
        objStream.WriteText "      ""properties"": {" & vbLf
        objStream.WriteText "        ""name"": """ & Replace(ActiveSheet.Cells(compteur, 4).Value, vbLf, " ") & """," & vbLf
        Dim i As Integer
        Dim description As String
        description = ""
        i = 2
        Do While i <= 17
            If i <> 13 And i <> 14 And i <> 5 And i <> 6 Then
                description = description & "**" & ActiveSheet.Cells(6, i).Value & " :** " & ActiveSheet.Cells(compteur, i).Value & "\n"
            End If
            i = i + 1
        Loop
        i = 33
        
        'boucle pour rajouter la "phase" dans la description (en fonction de comment est remplie la seconde partie du tableau)
        Do While i >= 21
            If IsEmpty(ActiveSheet.Cells(compteur, i).Value) = True Then
                i = i - 1
            Else
                description = description + "**Phase :** " & ActiveSheet.Cells(6, i).Value
                i = 0
            End If
        Loop
        description = Replace(description, """", "'")
        objStream.WriteText "        ""description"": """ & Replace(description, vbLf, " ") & """" & vbLf
        objStream.WriteText "      }," & vbLf
        objStream.WriteText "      ""geometry"": {" & vbLf
        objStream.WriteText "        ""type"": ""Point""," & vbLf
        objStream.WriteText "        ""coordinates"": [" & vbLf
        objStream.WriteText "        " & """" & lng1 & """" & "," & vbLf
        objStream.WriteText "        " & """" & lat1 & """" & vbLf
        objStream.WriteText "        ]" & vbLf
        objStream.WriteText "      }" & vbLf
        If compteur = nbLignes Then
            objStream.WriteText "    }" & vbLf
        Else
            objStream.WriteText "    }," & vbLf
        End If
        compteur = compteur + 1
    Loop
    objStream.WriteText "  ]" & vbLf
    objStream.WriteText "}"
    Dim error As Integer
    'enregistrer l'objet en un fichier en geoJson dans le dossier téléchargements de l'utilisateur
    On Error Resume Next
    objStream.SaveToFile Environ$("USERPROFILE") & "\Downloads" & "\" & ActiveSheet.Name & ".geojson", adSaveCreateOverWrite
    objStream.SaveToFile Environ$("USERPROFILE") & "\Téléchargements" & "\" & ActiveSheet.Name & ".geojson", adSaveCreateOverWrite
    On Error GoTo 0
    'fonction pour mettre à jour la carte Umap
    umapUpdate
    
    'Supprime le fichier en geoJson crée auparavant
    On Error Resume Next
    Kill Environ$("USERPROFILE") & "\Downloads" & "\" & ActiveSheet.Name & ".geojson"
    Kill Environ$("USERPROFILE") & "\Téléchargements" & "\" & ActiveSheet.Name & ".geojson"
    On Error GoTo 0
End Sub
