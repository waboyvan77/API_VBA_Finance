Attribute VB_Name = "Module1"
Option Explicit ' mention facultative qui rend la déclaration de variables obligatoire
                ' minimise ensuite le risque d'erreur
                
Sub RecupData()
    Const URL = "https://fr.wikipedia.org/wiki/JavaScript_Object_Notation"
'    Const URL = "https://api.ipify.org/?format=json"
    
    On Error GoTo err_RecupData
    Dim oRequest As Object
    Dim sResponse As String
    
    ' Création d'un objet requête HTTP
    Set oRequest = CreateObject("MSXML2.XMLHTTP")

    ' Paramétrage de la requête
    oRequest.Open "GET", URL, False
    
    ' Envoi de la requête au serveur
    oRequest.Send
    
    ' Récupération de la réponse
    sResponse = oRequest.ResponseText
    
    ' Exploitation de la réponse
    
    ' MsgBox sResponse
    ' Debug.Print sResponse
    ' Ouverture d'un fichier texte en écriture
    Open "H:\2025-2026\fichier.txt" For Output As #1
    ' Ecriture du HTML dans le fichier
    Print #1, sResponse
    ' Fermeture du fichier
    Close #1
    ' on peut ensuite ouvrir le fichier dans le bloc-notes
    ' ou le renommer en .html et l'ouvrir dans un navigateur
    
    Exit Sub
err_RecupData:
    MsgBox Err.Description
End Sub

Function VBCurl(sURL As String) As String
    On Error GoTo err_VBCurl
    
    ' La fonction VBCurl agit comme la commande cUrl, elle reçoit une URL et renvoie la réponse du serveur
    Dim oRequest As Object
    
    ' Création d'un objet requête HTTP
'    Set oRequest = CreateObject("MSXML2.XMLHTTP") ' ancienne version (deprecated)
    Set oRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0") ' nouvelle version
    
    ' Paramétrage de la requête
    oRequest.Open "GET", sURL, False
    
    ' Envoi de la requête au serveur
    oRequest.Send
    
    ' Récupération de la réponse
    VBCurl = oRequest.ResponseText
    
    Exit Function
err_VBCurl:
    VBCurl = ""
End Function

Sub DisplayIP()
    ' En utilisant une api afficher votre ip publique dans la cellule A2
    Dim s As String
    
    s = VBCurl("https://api.ipify.org/?format=json")
    
    ' Debug.Print s
    ' exemple de réponse de l'api :
    ' {"ip":"148.60.178.200"}
    ' on souhaite ne récupérer que l'adresse IP
    Cells(2, 1) = Mid(s, 8, Len(s) - 7 - 2)
    ' 8 correspond à la position du 1ère chiffre de l'adresse IP
    ' 7 correspond au nombre de caractères inutiles en début de json
    ' 2 correspond au nombre de caractères inutiles en fin de json
End Sub

Sub GetCountry()
    ' Ce programme interroge l'api IP-API pour récupérer le pays correspondant à une adresse IP donnée
    ' Il récupère l'adresse dans la cellule B4
    ' Et affiche le pays en B5
    Dim s As String
    Dim n1 As Long
    Dim n2 As Long
    
    s = VBCurl("http://ip-api.com/json/" & Cells(4, 2))
    ' (l'URL de l'api est directement copiée/collée depuis le navigateur)
    ' Debug.Print s
    ' exemple de réponse de l'api :
    ' {"status":"success","country":"United States","countryCode":"US","region":"TX","regionName":"Texas","city":"Richardson","zip":"75080","lat":32.9918,"lon":-96.7108,"timezone":"America/Chicago","isp":"AT\u0026T Enterprises, LLC","org":"ATT Global Anycast Prefix","as":"AS7018 AT\u0026T Enterprises, LLC","query":"99.99.99.99"}
    ' le nom du pays se situe entre "country":"  et ","countryCode"
    
    ' en cas d'échec (ip invalide), voici le json retourné :
    ' {"status":"fail","message":"invalid query","query":"150.60.70.280"}
    
    If InStr(1, s, """status"":""success""") > 0 Then
        ' L'appel est concluant
        n1 = InStr(1, s, """country"":""")  ' on double les guillements contenus dans la chaine recherchée
        If n1 > 0 Then
            ' on a trouvé le début du motif
            n2 = InStr(n1, s, """,""countryCode""")
            If n2 > 0 Then
                Cells(5, 2) = Mid(s, n1 + Len("""country"":"""), n2 - n1 - Len("""country"":"""))
            End If
        End If
    Else
        ' L'appel a échoué
        Cells(5, 2) = "Echec"
    End If
    
End Sub

Sub DisplayCountry()
    ' v2 de GetCountry en utilisant une fonction dédiée pour récupérer les valeurs dans le json
    Dim s As String
    
    s = VBCurl("http://ip-api.com/json/" & Cells(4, 2))
    If JsonValue(s, "status") = "success" Then
        Cells(5, 2) = JsonValue(s, "country")
    Else
        Cells(5, 2) = JsonValue(s, "message")
    End If

End Sub

Function JsonValue(jsonString As String, fieldName As String) As String
    ' Fonction récupérée sur Mistral AI
    ' Prompt : j'ai besoin d'une fonction vba JsonValue à laquelle je passe du json
    '          et le nom d'un champ et qui me renvoie sa valeur
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String

    ' Nettoyer le nom du champ (enlever les espaces avant/après)
    fieldName = Trim(fieldName)

    ' Trouver la position du champ dans la chaîne JSON
    startPos = InStr(1, jsonString, """" & fieldName & """", vbTextCompare)

    ' Si le champ n'est pas trouvé, retourner une chaîne vide
    If startPos = 0 Then
        JsonValue = ""
        Exit Function
    End If

    ' Trouver le début de la valeur (après le ":")
    startPos = InStr(startPos, jsonString, ":")
    If startPos = 0 Then
        JsonValue = ""
        Exit Function
    End If
    startPos = startPos + 1

    ' Ignorer les espaces après le ":"
    Do While Mid(jsonString, startPos, 1) = " "
        startPos = startPos + 1
    Loop

    ' Déterminer si la valeur est une chaîne (entre guillemets)
    If Mid(jsonString, startPos, 1) = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, jsonString, """")
        If endPos = 0 Then
            JsonValue = ""
            Exit Function
        End If
        value = Mid(jsonString, startPos, endPos - startPos)
    Else
        ' Cas d'un nombre ou d'un booléen (simplifié)
        endPos = InStr(startPos, jsonString, ",")
        If endPos = 0 Then endPos = InStr(startPos, jsonString, "}")
        If endPos = 0 Then endPos = Len(jsonString) + 1
        value = Trim(Mid(jsonString, startPos, endPos - startPos))
    End If

    JsonValue = value
End Function

Sub DisplayXY()
    ' Le programme utilise l'api IP-API.com dans sa version XML
    ' pour récupérer les coordonnées correspondant à une adresse IP
    Dim sXml As String
    
    sXml = VBCurl("http://ip-api.com/xml/" & Cells(4, 2))
    ' Debug.Print sXml
    ' Exemple de XML :
        '<?xml version="1.0" encoding="UTF-8"?>
        '<query>
        '  <status>success</status>
        '  <country>Japan</country>
        '  <countryCode>JP</countryCode>
        '  <region>13</region>
        '  <regionName>Tokyo</regionName>
        '  <city>Chiyoda City</city>
        '  <zip>100-8111</zip>
        '  <lat>35.6906</lat>
        '  <lon>139.77</lon>
        '  <timezone>Asia/Tokyo</timezone>
        '  <isp>KDDI Web Communications Inc.</isp>
        '  <org>CPI</org>
        '  <as>AS9597 KDDI Web Communications Inc.</as>
        '  <query>150.60.70.80</query>
        '</query>
    If XmlValue(sXml, "status") = "success" Then
        Cells(9, 2) = XmlValue(sXml, "lat")
        Cells(10, 2) = XmlValue(sXml, "lon")
    Else
        Cells(9, 2) = XmlValue(sXml, "message")
        Cells(10, 2) = ""
    End If
End Sub

Function XmlValue(xml As String, nodeName As String) As String
    ' Fonction générée par Mistral en réponse au prompt :
    ' il me faudrait la même fonction pour du XML
    Dim posDebut As Long
    Dim posFin As Long
    Dim valeur As String

    ' Nettoyer le nom du nœud
    nodeName = Trim(nodeName)

    ' Chercher l'ouverture du nœud
    posDebut = InStr(1, xml, "<" & nodeName & ">", vbTextCompare)
    If posDebut = 0 Then
        XmlValue = "Nœud non trouvé"
        Exit Function
    End If

    ' Trouver la fin de la balise d'ouverture
    posDebut = posDebut + Len(nodeName) + 2

    ' Trouver la balise de fermeture correspondante
    posFin = InStr(posDebut, xml, "</" & nodeName & ">")
    If posFin = 0 Then
        XmlValue = "Balisage XML invalide"
        Exit Function
    End If

    ' Extraire la valeur entre les balises
    valeur = Mid(xml, posDebut, posFin - posDebut)
    XmlValue = Trim(valeur)
End Function

Sub SunriseSunset()
    ' Utilisation de l'api https://api.sunrise-sunset.org pour récupérer
    ' les heures de lever et de coucher du soleil
    
    Dim sURL As String
    Dim sJson As String
    
    ' Construction de l'URL - URL Building
    sURL = "https://api.sunrise-sunset.org/json"
    sURL = sURL & "?lat=" & Replace(Cells(1, 2), ",", ".") ' Le 1er paramètre est introduit par un ?
                                                           ' 1st param starts with a ?
    sURL = sURL & "&lng=" & Replace(Cells(2, 2), ",", ".") ' Les paramètres suivants commencent par &
                                            ' Next params start with &
    sURL = sURL & "&date=" & Format(Cells(3, 2), "yyyy-mm-dd")
    
    ' Appel de l'API - API Call
    sJson = VBCurl(sURL)
    ' Debug.Print sJson
    ' Exemple de réponse (example answer) :
    ' {"results":{"sunrise":"5:33:32 AM","sunset":"6:43:02 PM","solar_noon":"12:08:17 PM","day_length":"13:09:30","civil_twilight_begin":"5:04:06 AM","civil_twilight_end":"7:12:28 PM","nautical_twilight_begin":"4:26:52 AM","nautical_twilight_end":"7:49:42 PM","astronomical_twilight_begin":"3:47:25 AM","astronomical_twilight_end":"8:29:09 PM"},"status":"OK","tzid":"UTC"}
    ' Affichage des résultats (display results)
    Cells(5, 2) = JsonValue(sJson, "sunrise")
    Cells(6, 2) = JsonValue(sJson, "sunset")
End Sub


Sub getCAC40()
    ' Cette macro met à jour la feuille avec les derniers cours du CAC40
    ' https://www.abcbourse.com/marches/indice_cac40
    ' The macro gets the Paris stock exchange prices
    Dim sURL As String
    Dim sHtml As String
    Dim n1 As Long
    Dim n2 As Long
    Dim sHtmlAction As String
    Dim sAction As String
    Dim sValeur As String
    Dim nRow As Long
    
    sURL = "https://www.abcbourse.com/marches/indice_cac40"
    sURL = sURL & "?param=" & Format(Now, "yyyymmddhhnnss")
    ' on ajoute un paramètre inutile pour être certain que notre requête est bien exécutée
    ' (We add a useless parameter to make sure our request is actually executed)
    sHtml = VBCurl(sURL)
    
    ' En regardant le code source de la page, on constate que la liste des actions est dans
    ' un tableau défini ainsi :
    ' (By looking at the page's source code, we can see that the list of stocks is in a table
    ' defined as follows :)
    ' <table class="tablesorter tbl100_6 mt5" id="tabQuotes">
    
    n2 = InStr(1, sHtml, "<table class=""tablesorter tbl100_6 mt5"" id=""tabQuotes"">")
    If n2 = 0 Then Exit Sub ' html non valide, on quitte la macro
    ' On a trouvé le début du tableau contenant les actions
    ' (We found the beginning of the table containing the stocks)
    ' On constate que chaque action est décrite dans une ligne de tableau comme celle-ci :
    ' (Each stock is described in a table row like this one : )
'                    <tr data-sx="AC_25" data-name="ACp">
'                        <td class="srd"><a href="/cotation/ACp">Accor Hotels</a></td>
'                        <td>41,40</td>
'                        <td>41,46</td>
'                        <td>40,77</td>
'                        <td >136230</td>
'                        <td>41,05</td>
'                        <td class="bold">40,90</td>
'                        <td class="quote_downb">-0,37%</td>
'                    </tr>
    ' Toutes les lignes du tableau commencent par <tr data-sx=
    ' (Every row in the table start with <tr data-sx=)
    ' Et se terminent par </tr> (and end with </tr>)
    nRow = 2
    Do
        n1 = InStr(n2, sHtml, "<tr data-sx=")
        If n1 > 0 Then
            n2 = InStr(n1, sHtml, "</tr>")
            ' on récupère le HTML d'1 action
            sHtmlAction = Mid(sHtml, n1, n2 - n1)
            ' on extrait le nom et le dernier cours de cette action
            TraiteAction sHtmlAction, sAction, sValeur
            ' on écrit ces données dans le classeur
            nRow = nRow + 1
            Cells(nRow, 1) = sAction
            Cells(nRow, 2) = sValeur
        End If
    Loop Until n1 = 0
End Sub

Sub TraiteAction(ByVal sHtml As String, ByRef sNom As String, ByRef sDernierCours As String)
    ' Ce sous-programme reçoit le HTML décrivant une action et renvoie son nom et son dernier cours
    ' This subroutine receives the HTML describing a stock and returns its name and last price
    ' exemple de HTML :
'                    <tr data-sx="AC_25" data-name="ACp">
'                        <td class="srd"><a href="/cotation/ACp">Accor Hotels</a></td>
'                        <td>41,40</td>
'                        <td>41,46</td>
'                        <td>40,77</td>
'                        <td >136230</td>
'                        <td>41,05</td>
'                        <td class="bold">40,90</td>
'                        <td class="quote_downb">-0,37%</td>
'                    </tr>
    Dim n1 As Long
    Dim n2 As Long
    
    sNom = ""
    sDernierCours = ""
    
    ' Le nom de l'action si situe entre le 3ème > et </a>
    ' The name of the stock is located between the third > and </a>
    ' Recherche du 1er >
    n1 = InStr(1, sHtml, ">")
    If n1 = 0 Then Exit Sub ' html invalide, on quitte la procedure
    ' Recherche du 2ème >
    n1 = InStr(n1 + 1, sHtml, ">")
    If n1 = 0 Then Exit Sub ' html invalide, on quitte la procedure
    ' Recherche du 3ème >
    n1 = InStr(n1 + 1, sHtml, ">")
    If n1 = 0 Then Exit Sub ' html invalide, on quitte la procedure
    ' Recherche de </a>
    n2 = InStr(n1 + 1, sHtml, "</a>")
    If n2 = 0 Then Exit Sub ' html invalide, on quitte la procedure
    ' On récupère le nom de l'action
    sNom = Mid(sHtml, n1 + 1, n2 - n1 - 1)
    
    ' Le dernier cours de l'action est compris entre <td class="bold"> et </td>
    ' (The stock's last price is located between <td class="bold"> and </td>)
    n1 = InStr(n2, sHtml, "<td class=""bold"">")
    If n1 = 0 Then Exit Sub ' html invalide, on quitte la procedure
    n2 = InStr(n1, sHtml, "</td>")
    If n2 = 0 Then Exit Sub ' html invalide, on quitte la procedure
    ' On récupère le dernier cours de l'action
    sDernierCours = Mid(sHtml, n1 + Len("<td class=""bold"">"), n2 - n1 - Len("<td class=""bold"">"))
    ' On transforme les , en . car Excel considère les , comme un séparateur de milliers
    ' We replace commas with periods because Excel treats commas as thousand separators
    sDernierCours = Replace(sDernierCours, ",", ".")
End Sub






