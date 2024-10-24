Attribute VB_Name = "Module1"
Sub CallSAPProcessAutomationAPI()
    Dim objHTTP As Object
    Dim strURL As String
    Dim strAccessToken As String
    Dim strPayload As String
    Dim strResponse As String
    
    ' Create HTTP object
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Set the URL for the API call
    strURL = "https://spa-api-gateway-bpi-us-prod.cfapps.us10.hana.ondemand.com/workflow/rest/v1/workflow-instances"
    
    ' Get the access token
    strAccessToken = GetAccessToken()
    
    ' Set the payload
    strPayload = "{""definitionId"": ""us10.yk2lt6xsylvfx4dz.zreadinvoicecopy1.getInvoiceDetails"", ""context"": {""outlook"": ""Inbox"", ""temporary"": ""C:\\Users\\jibin\\Downloads\\Work-BTP\\Msitek\\SAP Build process Automation\\Outlook Folder""}}"
    
    ' Open the connection
    objHTTP.Open "POST", strURL, False
    
    ' Set the headers
    objHTTP.setRequestHeader "Authorization", "Bearer " & strAccessToken
    objHTTP.setRequestHeader "Accept", "application/json"
    objHTTP.setRequestHeader "Content-Type", "application/json"
    
    ' Send the request
    objHTTP.Send strPayload
    
    ' Get the response
    strResponse = objHTTP.responseText
    
    ' Display the response (you can modify this part as needed)
    Debug.Print strResponse
    
    ' Clean up
    Set objHTTP = Nothing
End Sub

Function GetAccessToken() As String
    Dim objHTTP As Object
    Dim strTokenURL As String
    Dim strClientID As String
    Dim strClientSecret As String
    Dim strAuth As String
    Dim strResponse As String
    
    ' Create HTTP object
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Set the token URL
    strTokenURL = "https://yk2lt6xsylvfx4dz.authentication.us10.hana.ondemand.com/oauth/token"
    
    ' Set client credentials
    strClientID = "sb-4516453a-ad79-400f-b87c-c5d47a354173!b220961|xsuaa!b49390"
    strClientSecret = "7891dd5c-912b-42e5-a243-b4125acd6d34$PEutIDKTAD-RUECciDCTZA4WqAkFaDc43egOCufx0IU="
    
    ' Create the authorization string
    strAuth = Base64Encode(strClientID & ":" & strClientSecret)
    
    ' Open the connection
    objHTTP.Open "POST", strTokenURL, False
    
    ' Set the headers
    objHTTP.setRequestHeader "Authorization", "Basic " & strAuth
    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Send the request
    objHTTP.Send "grant_type=client_credentials"
    
    ' Parse the response to get the access token
    strResponse = objHTTP.responseText
    GetAccessToken = ParseAccessToken(strResponse)
    
    ' Clean up
    Set objHTTP = Nothing
End Function

Function ParseAccessToken(strResponse As String) As String
    ' This is a simple parser. You might want to use a JSON parser for more robust handling
    Dim arrParts As Variant
    arrParts = Split(strResponse, """")
    Dim i As Integer
    For i = 0 To UBound(arrParts)
        If arrParts(i) = "access_token" Then
            ParseAccessToken = arrParts(i + 2)
            Exit Function
        End If
    Next i
End Function

Function Base64Encode(sText)
    Dim bArr() As Byte
    Dim sBase64 As String
    Dim oXML As Object
    
    bArr = StrConv(sText, vbFromUnicode)
    Set oXML = CreateObject("MSXML2.DOMDocument")
    With oXML.createElement("B64")
        .dataType = "bin.base64"
        .nodeTypedValue = bArr
        sBase64 = .text
    End With
    Set oXML = Nothing
    Base64Encode = Replace(sBase64, vbLf, "")
End Function
