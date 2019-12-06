Attribute VB_Name = "FIXER_VBA_API_JSON"
'=============================================
' VBA API for Fixer.io foreign exchange rates
'
' Uses VBA-JSON v2.3.1 @ https://github.com/VBA-tools/VBA-JSON
'
' @author hagend87@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'=============================================

Sub FIXER_VBA_API_JSON()

'Declare variables
Dim xml_obj As MSXML2.XMLHTTP60

'Create a new Request Object
Set xml_obj = New MSXML2.XMLHTTP60

    'Define URL Components
    base_url = "http://data.fixer.io/api"
    endpoint = "/latest?"
    
    param_api = "access_key="
    param_api_value = CStr(Worksheets("api_key").Range("A1").Value)
    
        ' Latest Rates Endpoint
        ' ? access_key = API_KEY
        ' & base = USD
        ' & symbols = GBP,JPY,EUR
        param_base = "&base="
        param_base_value = "EUR"
        
        param_symbols = "&symbols="
        param_symbols_value = "USD"
    
    'Combine all the different components into a single URL
    api_url = base_url + endpoint + _
              param_api + param_api_value + _
              param_base + param_base_value + _
              param_symbols + param_symbols_value
              
    Debug.Print (api_url)
    
    
    'Open a new request using our URL
    xml_obj.Open bstrMethod:="GET", bstrURL:=api_url
    
    'Send the request
    xml_obj.send
    
    'Print the status code in case something went wrong
    Debug.Print "The Request was " + CStr(xml_obj.Status)
    
    'Define a few object variables
    Dim Json As Object
    Dim result As Dictionary
    
    'Parse the response
    Set Json = JsonConverter.ParseJson(xml_obj.responseText)
    
    Debug.Print JsonConverter.ConvertToJson(Json)
    Debug.Print Json("rates")("USD")

End Sub
