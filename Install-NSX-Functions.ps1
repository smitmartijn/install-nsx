

# This function is from: http://sharpcodenotes.blogspot.nl/2013/03/how-to-make-http-request-with-powershell.html
function Http-Web-Request([string]$method,[string]$encoding,[string]$server,[string]$path,$headers,[string]$postData)
{
  $return_value = New-Object PsObject -Property @{httpCode = ""; httpResponse = ""}

  ## Compose the URL and create the request
  $url = "$server/$path"
  [System.Net.HttpWebRequest] $request = [System.Net.HttpWebRequest] [System.Net.WebRequest]::Create($url)

	# Ignore SSL certificate errors
  [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

  ## Add the method (GET, POST, etc.)
  $request.Method = $method
  ## Add an headers to the request
  foreach($key in $headers.keys)
  {
    $request.Headers.Add($key, $headers[$key])
  }

  ## We are using $encoding for the request as well as the expected response
  $request.Accept = $encoding
  ## Send a custom user agent if you want
  $request.UserAgent = "PowerShell script"

  ## Create the request body if the verb accepts it (NOTE: utf-8 is assumed here)
  if ($method -eq "POST" -or $method -eq "PUT") {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($postData)
    $request.ContentType = $encoding
    $request.ContentLength = $bytes.Length

    [System.IO.Stream] $outputStream = [System.IO.Stream]$request.GetRequestStream()
    $outputStream.Write($bytes,0,$bytes.Length)
    $outputStream.Close()
  }

  ## This is where we actually make the call.
  try
  {
    [System.Net.HttpWebResponse] $response = [System.Net.HttpWebResponse] $request.GetResponse()
    $sr = New-Object System.IO.StreamReader($response.GetResponseStream())
    $txt = $sr.ReadToEnd()
    ## NOTE: comment out the next line if you don't want this function to print to the terminal
    #Write-Host "CONTENT-TYPE: " $response.ContentType
    ## NOTE: comment out the next line if you don't want this function to print to the terminal
    #Write-Host "RAW RESPONSE DATA:" . $txt
    ## Return the response body to the caller
    $return_value.httpResponse = $txt
    $return_value.httpCode = [int]$response.StatusCode

    return $return_value
  }
  ## This catches errors from the server (404, 500, 501, etc.)
  catch [Net.WebException] {
    [System.Net.HttpWebResponse] $resp = [System.Net.HttpWebResponse] $_.Exception.Response
    ## NOTE: comment out the next line if you don't want this function to print to the terminal
    #Write-Host $resp.StatusCode -ForegroundColor Red -BackgroundColor Yellow
    ## NOTE: comment out the next line if you don't want this function to print to the terminal
    #Write-Host $resp.StatusDescription -ForegroundColor Red -BackgroundColor Yellow
    ## Return the error to the caller
    $return_value.httpResponse = $resp.StatusDescription
    $return_value.httpCode = [int]$resp.StatusCode

    return $return_value
  }
}

# Execute NSX API Calls using Http-Web-Request
function NSX-API-Call([string]$NSX_Manager_IP, [string]$NSX_Manager_Username, [string]$NSX_Manager_Password, [string]$method, [string]$URL, [string]$postData)
{
  # Format authentication header
  $auth    = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($NSX_Manager_Username + ":" + $NSX_Manager_Password))
  $headers = @{ Authorization = "Basic $auth" }

  $result = Http-Web-Request $method "application/xml" "https://$NSX_Manager_IP" $URL $headers $postData
  return $result
}

# Generic function to release memory
function Release-Ref ($ref) {
    ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
