# Authored by: Ryan Johnson 
# Special thanks to Nick DiMartinis for his valuable ideas on my script and for being a helpful mentor!

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Initialization of IP mode variables to keep track of what mode we are in
$SingleIPMode = $false 
$ManualIPMode = $false
$TextFileMode = $false

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Allows the browse feature to work for each input box (as well as setting the text of each box to match the path selected)
function FileDialogInput
{
$FileBrowser.ShowDialog()
$global:inputPath = $FileBrowser.FileName 
$inputPathBox.Text = $inputPath
$inputPathBox.Refresh()
}

# Allows the browse feature to work for each output box (as well as setting the text of each box to match the path selected)
function FileDialogOutput
{
$FileBrowser.ShowDialog()
$global:outputPath = $FileBrowser.FileName
if ($TextFileMode -eq $true) {
$outputPathBox.Text = $outputPath
$outputPathBox.Refresh()
}
elseif (($SingleIPMode -eq $true) -or ($ManualIPMode -eq $true)) {
$outputPathBoxSingle.Text = $outputPath
$outputPathBoxSingle.Refresh()
}
}


# Used when Run button is pressed; Uses conditional logic to identify which mode we are currently in and executes the appropriate function as a result
Function Run
{
    $outputBox.Text = "Plese wait (Note: If the program is unresponsive, it's still working, don't close it.)"
    $outputBox.Refresh()

    # Checks if excel is opened; if it is, the user is notified and recieves an error message
    $excel = Get-Process Excel -ErrorAction SilentlyContinue
    if ($excel) {
        $outputBox.Text = "Please close excel before running"
        $outputBox.Refresh()
            Exit
        }
    Remove-Variable Excel 

    # Checks what mode the user is currently in and executes the correct function based on that info
    if ($TextFileMode -eq $true) {
        if (($inputPathBox.Text -eq "") -or ($outputPathBox.Text -eq "")) { # check for if a path is missing
            $outputBox.Text = "Error: Path not selected."
            $outputBox.Refresh()
            break
        }
    $outputBox.Text = TxtBulkLookupIPAddresses -filePath $inputPath -outputFile $outputPath | Out-String
    $outputBox.Refresh()
    }
    elseif ($SingleIPMode -eq $true) {
        if (($SingleIPInputBox.Text -eq "") -or ($outputPathBoxSingle.Text -eq "")) { # check for if a path is missing
            $outputBox.Text = "Error: Path not selected."
            $outputBox.Refresh()
            break
        }
    $outputBox.Text = SingleLookupIPAddress -ipAddress $SingleIPInputBox.text -outputFile $outputPath | Out-String
    $outputBox.Refresh()
    }
    elseif ($ManualIPMode -eq $true) {
        if (($manualInputBox.Text -eq "") -or ($outputPathBoxSingle.Text -eq "")) { # check for if a path is missing
            $outputBox.Text = "Error: Path not selected."
            $outputBox.Refresh()
            break
        }
    $outputBox.Text = ManualBulkLookupIPAddresses -outputFile $outputPath -inputArray $inputArray | Out-String
    $outputBox.Refresh()
    }
    else {
    $outputBox.Text = "No mode selected. Please try again."
    $outputBox.Refresh()
    }

}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Function to clear the contents of each control in text file mode (executed when we switch to a different mode to keep things clean)
Function TextFileModeClear {
$outputBox.Text = ""
$inputPathBox.Text = ""
$outputPathBox.Text = ""
}

# Changes the visibility of the text file controls (toggleable; hides them when mode is not activated and turns them on when activated)
Function TextFileVisibility {
$TxtFileBtn.Visible = -not($TxtFileBtn.Visible)
$BrowseButton1.Visible = -not($TxtFileBtn.Visible)
$BrowseButton2.Visible = -not($TxtFileBtn.Visible)
$Label1.Visible = -not($TxtFileBtn.Visible)
$Label2.Visible = -not($TxtFileBtn.Visible)
$inputPathBox.Visible = -not($TxtFileBtn.Visible)
$outputPathBox.Visible = -not($TxtFileBtn.Visible)
}

# Runs when text file mode is clicked on and selected (sets up controls and updates GUI)
Function TextFileOnClick {

if ($SingleIPMode -eq $true) {
-not(SingleIPVisibility)
SingleIPClear
}
elseif ($ManualIPMode -eq $true) {
-not(ManualIPVisibility)
ManualIPClear
}

# Global changes are made to these variables so that they are kept track of appropriately
$Global:ManualIPMode = $false
$Global:SingleIPMode = $false
$Global:TextFileMode = -not($TextFileMode) # Each change has to be set to global to work

# change text file controls to be visibile when this button is clicked
TextFileVisibility 
}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Function to clear the contents of each control in single ip mode (executed when we switch to a different mode to keep things clean)
Function SingleIPClear {
$outputBox.Text = ""
$SingleIPInputBox.Text = ""
$outputPathBoxSingle.Text = ""
}

# Changes the visibility of the single IP mode controls (toggleable; hides them when mode is not activated and turns them on when activated)
Function SingleIPVisibility {
$SingleIPBtn.Visible = -not($SingleIPBtn.Visible)
$outputPathBoxSingle.visible = -not($SingleIPBtn.Visible)
$Label2Single.visible = -not($SingleIPBtn.Visible)
$BrowseButton2Single.Visible = -not($SingleIPBtn.Visible)
$SingleIPInputBox.Visible = -not($SingleIPBtn.Visible)
$SingleIPInputLabel.Visible = -not($SingleIPBtn.Visible)
}

# Runs when single IP mode is clicked on and selected (sets up controls and updates GUI)
Function SingleIPOnClick {

if ($TextFileMode -eq $true) { 
-not(TextFileVisibility)
TextFileModeClear
}
elseif ($ManualIPMode -eq $true) {
-not(ManualIPVisibility)
ManualIPClear
}

# Global changes are made to these variables so that they are kept track of appropriately
$Global:ManualIPMode = $false
$Global:TextFileMode = $false
$Global:SingleIPMode = -not($SingleIPMode) # Each change has to be set to global to work

# change single IP mode controls to be visibile when this button is clicked
SingleIPVisibility
}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Array is initialized to keep track of added IPs
$inputArray = @()

# Function to clear the contents of each control in manual ip mode (executed when we switch to a different mode to keep things clean)
Function ManualIPClear {
ClearClick
$outputBox.Text = ""
$manualInputBox.Text = ""
$outputPathBoxSingle.Text = ""
}

# Changes the visibility of the manual IP mode controls (toggleable; hides them when mode is not activated and turns them on when activated)
Function ManualIPVisibility {
$ManualIPBtn.Visible = -not($ManualIPBtn.Visible)
$outputPathBoxSingle.Visible = -not($ManualIPBtn.Visible)
$Label2Single.Visible = -not($ManualIPBtn.Visible)
$BrowseButton2Single.Visible = -not($ManualIPBtn.Visible)
$manualInputBox.Visible = -not($ManualIPBtn.Visible)
$AddButton.Visible = -not($ManualIPBtn.Visible)
$ClearButton.Visible = -not($ManualIPBtn.Visible)
$ManualIPInputLabel.Visible = -not($ManualIPBtn.Visible)
}

# Runs when manual IP mode is clicked on and selected (sets up controls and updates GUI)
Function ManualIPOnClick {

if ($TextFileMode -eq $true) { 
-not(TextFileVisibility)
TextFileModeClear
}
elseif ($SingleIPMode -eq $true) {
-not(SingleIPVisibility)
SingleIPClear
}

# Global changes are made to these variables so that they are kept track of appropriately
$Global:SingleIPMode = $false
$Global:TextFileMode = $false
$Global:ManualIPMode = -not($ManualIPMode) # Each change has to be set to global to work

# change manual IP mode controls to be visibile when this button is clicked
ManualIPVisibility
}

# Function to add different IPs to the input array
Function AddClick
{
$Global:inputArray += $manualInputBox.Text
$outputBox.Text = "IP Address Added: " + $manualInputBox.Text
$outputBox.Refresh()
}

# Function to clear input array
Function ClearClick
{
$Global:inputArray = @() #resets array
$outputBox.Text = "IP Addresses successfully erased."
$outputBox.Refresh()
}

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Install .Net Assemblies
Add-Type -AssemblyName System.Windows.Forms

# Makes SJI logo the app's icon  (uses base 64 so its visible on all platforms)
$iconBase64 = '/9j/4AAQSkZJRgABAQEBLAEsAAD/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAEAAAAAAAD/7QAsUGhvdG9zaG9wIDMuMAA4QklNA+0AAAAAABABLAAAAAEAAQEsAAAAAQAB/+FBWWh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8APD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4NCjx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuNi1jMTQzIDc5LjE2MTIxMCwgMjAxNy8wOC8xMS0xMDoyODozNiAgICAgICAgIj4NCgk8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPg0KCQk8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIgeG1sbnM6eG1wR0ltZz0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL2cvaW1nLyIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiB4bWxuczppbGx1c3RyYXRvcj0iaHR0cDovL25zLmFkb2JlLmNvbS9pbGx1c3RyYXRvci8xLjAvIiB4bWxuczpwZGY9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGRmLzEuMy8iPg0KCQkJPGRjOmZvcm1hdD5pbWFnZS9qcGVnPC9kYzpmb3JtYXQ+DQoJCQk8ZGM6dGl0bGU+DQoJCQkJPHJkZjpBbHQ+DQoJCQkJCTxyZGY6bGkgeG1sOmxhbmc9IngtZGVmYXVsdCI+NjIwOV9TSklfTG9nb19GdWxsTmFtZV9NYXhIZWlnaHRGbGFtZV8wMi4wMi4xODwvcmRmOmxpPg0KCQkJCTwvcmRmOkFsdD4NCgkJCTwvZGM6dGl0bGU+DQoJCQk8eG1wOk1ldGFkYXRhRGF0ZT4yMDE4LTAyLTE0VDE2OjU5OjU5LTA1OjAwPC94bXA6TWV0YWRhdGFEYXRlPg0KCQkJPHhtcDpNb2RpZnlEYXRlPjIwMTgtMDItMTRUMjE6NTk6NTlaPC94bXA6TW9kaWZ5RGF0ZT4NCgkJCTx4bXA6Q3JlYXRlRGF0ZT4yMDE4LTAyLTE0VDE2OjU5OjU5LTA1OjAwPC94bXA6Q3JlYXRlRGF0ZT4NCgkJCTx4bXA6Q3JlYXRvclRvb2w+QWRvYmUgSWxsdXN0cmF0b3IgQ0MgMjIuMCAoTWFjaW50b3NoKTwveG1wOkNyZWF0b3JUb29sPg0KCQkJPHhtcDpUaHVtYm5haWxzPg0KCQkJCTxyZGY6QWx0Pg0KCQkJCQk8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4NCgkJCQkJCTx4bXBHSW1nOndpZHRoPjE2ODwveG1wR0ltZzp3aWR0aD4NCgkJCQkJCTx4bXBHSW1nOmhlaWdodD4yNTY8L3htcEdJbWc6aGVpZ2h0Pg0KCQkJCQkJPHhtcEdJbWc6Zm9ybWF0PkpQRUc8L3htcEdJbWc6Zm9ybWF0Pg0KCQkJCQkJPHhtcEdJbWc6aW1hZ2U+LzlqLzRBQVFTa1pKUmdBQkFnRUJMQUVzQUFELzdRQXNVR2h2ZEc5emFHOXdJRE11TUFBNFFrbE5BKzBBQUFBQUFCQUJMQUFBQUFFQQ0KQVFFc0FBQUFBUUFCLys0QURrRmtiMkpsQUdUQUFBQUFBZi9iQUlRQUJnUUVCQVVFQmdVRkJna0dCUVlKQ3dnR0JnZ0xEQW9LQ3dvSw0KREJBTURBd01EQXdRREE0UEVBOE9EQk1URkJRVEV4d2JHeHNjSHg4Zkh4OGZIeDhmSHdFSEJ3Y05EQTBZRUJBWUdoVVJGUm9mSHg4Zg0KSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmLzhBQUVRZ0JBQUNvQXdFUg0KQUFJUkFRTVJBZi9FQWFJQUFBQUhBUUVCQVFFQUFBQUFBQUFBQUFRRkF3SUdBUUFIQ0FrS0N3RUFBZ0lEQVFFQkFRRUFBQUFBQUFBQQ0KQVFBQ0F3UUZCZ2NJQ1FvTEVBQUNBUU1EQWdRQ0JnY0RCQUlHQW5NQkFnTVJCQUFGSVJJeFFWRUdFMkVpY1lFVU1wR2hCeFd4UWlQQg0KVXRIaE14Wmk4Q1J5Z3ZFbFF6UlRrcUt5WTNQQ05VUW5rNk96TmhkVVpIVEQwdUlJSm9NSkNoZ1poSlJGUnFTMFZ0TlZLQnJ5NC9QRQ0KMU9UMFpYV0ZsYVcxeGRYbDlXWjJocGFtdHNiVzV2WTNSMWRuZDRlWHA3ZkgxK2YzT0VoWWFIaUltS2k0eU5qbytDazVTVmxwZVltWg0KcWJuSjJlbjVLanBLV21wNmlwcXF1c3JhNnZvUkFBSUNBUUlEQlFVRUJRWUVDQU1EYlFFQUFoRURCQ0VTTVVFRlVSTmhJZ1p4Z1pFeQ0Kb2JId0ZNSFI0U05DRlZKaWN2RXpKRFJEZ2hhU1V5V2lZN0xDQjNQU05lSkVneGRVa3dnSkNoZ1pKalpGR2lka2RGVTM4cU96d3lncA0KMCtQemhKU2t0TVRVNVBSbGRZV1ZwYlhGMWVYMVJsWm1kb2FXcHJiRzF1YjJSMWRuZDRlWHA3ZkgxK2YzT0VoWWFIaUltS2k0eU5qbw0KK0RsSldXbDVpWm1wdWNuWjZma3FPa3BhYW5xS21xcTZ5dHJxK3YvYUFBd0RBUUFDRVFNUkFEOEE5VTRxN0ZYWXE3RlhZcTdGWFlxNw0KRlhZcTdGV0VmbVgrYkhsL3lOWmdYQit1YXhNdkswMHVOZ0hic0hrTy9wcFh1UnYyQjN6SjArbGxsUGwzdU5xTlRIR1BQdVVmeXVsMQ0KN1VrdU5hOHhURjlZdUkxTFc2L0REYlJTa3NzRWNlL0hqd0hJbjRpZXBOQmgxSWlQVEhralRtUjNselo3bUs1VHNWZGlyc1ZkaXJzVg0KZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZVkvbkgrY2xsNUxzenAybWxMbnpMY0pXS0kvRWx1amRKWlI0L3lwMzZuYnJtNg0KVFNISWJQMHVIcXRVTVlvZlUrZC9JOXJmZWFQT2phcnE4cjNqUXQ5YnZKcFR5THlWL2RxZjlsMjZVRk0yMllpRUtEcXNJTTUyWDFEKw0KWFYya2tsL0VXK01pSmxYdVFDd0oraW96UzV4eWR6Z1BObXVZemtPeFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4Vg0KMktzRi9Oejh6TFR5TjVlTXljWmRhdlEwZW1XeDMrSUQ0cFhIOGtkZnBOQjdqSzB1bk9XWGtPYmphblVESEh6ZkcybzZqZTZsZjNGLw0KZlROY1hsMDdTM0V6N3N6c2Frbk9nakVBVUhRU2tTYkwxTDhzN05iSHkvOEFXQ0tTM3JtUW52d1Q0VUg0RS9UbUZxRGN2YzVtbkZSOQ0KNzBmeWo1aVRTOWR0NTVXcGJ5ZnViZzlnajl6L0FLclViNk13OHVQaWk1ZUtmREo3VUNDS2pwbXVkaTdGWFlxN0ZYWXE3RlhZcTdGWA0KWXE3RlhZcTdGWFlxN0ZYWXE3RlZHK3ZiV3hzcDcyN2tFVnJiUnROUEszUlVqVXN6SDVBWVFDVFFRU0FMTDRnL01YenRlZWN2TmQzcg0KTTVaWUdQcFdFRGY3cXQwSjlOTnUvd0MwMytVVG5TYWZDTWNRSG5jK1k1SkVzWnk5cGV6NmVSYVdGdmFxYUNHTkkvOEFnVkF6WFMzTg0KdWZIWVVyL1d6NDRLVGIyRDhyL09jZXBXUTBlN2tIMSsxWDl3VDFsaEhUNXNuVDVmVG12MU9Halk1T2ZwOHRpanpaN21LNVRzVmRpcg0Kc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyeUwvbkpqelMrbGVSNHRKZ2ZoY2E1TjZUMDJQMWVHankwK2JGRlBzVG1mMmZqNA0KcDMzT0QyaGs0WVYzdmxITjY2UlZ0RjVYY0swcnlrVVUrWkdBOGtqbTlVK3MrK1lOT1pidnJQdmpTMnJXZXEzTmxkUlhkcktZcmlCZw0KOFVpOVF3d0dOaWlrU28ySDBINUM4OVdYbWpUNmtyRHFrQUgxdTFCK2oxRXIxUS9oMFBhdXF6NFRBK1R0Y09ZVEhteW5LRzUyS3V4Vg0KMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdmw3L25LblVXbDg1NlhwNE5ZN1N3RXROOW5ubGNOL3dzUzV1dXpZK2duemROMmpMMQ0KZ2VUeFhOazY5ZkEvQ2FOLzVXQis0NENvZWgvV1BmTVNuS3QzMWozeHBiZDlZOThhVzBYcE91NmhwT29RNmhZVEdDNmdQSkhINGdqdQ0KRDBJT1JsQVNGRmxHWmliRDZROGdlZjhBVHZOdW5jazR3YXBBQjljczY5TzNOSzdsRCtIUSsrb3o0RGpQazdiQm5FeDVzcXlodmRpcg0Kc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZktYL09VTVRwK1kxdXpDaXlhYkF5SHhBbG1YOVl6ZWRuSDkzOFhTZG9EOTU4SGtPYg0KQndYWXF6Q3l1L1Z0SXBLN2xSWDVqWS9qbU9SdTNnN0szcmUrQ2syNzF2ZkdsdDNyZStOTGFOMGZYdFIwZlVZZFIwNmN3WGNEY2tjZg0KaXJEb1ZJMklPUm5BU0ZGbENaaWJENlcvTGo4eU5PODM2ZnhQRzMxaTNVZlc3T3ZYdDZrVmR5aC80WG9leE9tMUduT00rVHQ4R29HUQ0KZWJNc3gzSWRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXI1NS81eXYwUnVXZzY0Z0pXa3RqTzNZSGFXSWZUKzh6YmRtVDV4K0xxdQ0KMG9jcFBuck5zNnQyS3BybzEzUldnWTlQaVQrT1Z6RE9KVFAxY2hUTzNlcmpTMjcxY2FXM2VyalMyaXRNMWkrMHkvaHY3Q2RyZTd0Mg0KRHd6SWFFRWZyQjZFSHJrWlFFaFJUR1JCc1BxRDhzUHpQMC96alllak54dDljdDFCdXJVR2djRGIxWXE5VlBjZnMvY1RwZFRwampQOQ0KRjNPbjFBeUQra3puTVp5WFlxN0ZYWXE3RlhZcW9yY3E5MDBDYittb01yZUJiN0svTWpmL0FHOGh4M0ttWmhVYlBWV3liQjJLdXhWMg0KS3NQL0FEYThvbnpYNUQxUFRJVTUzeUo5WnNBTno2OEh4S285M0ZVLzJXWkdseThHUUhvMGFuRnh3STZ2aVFnZ2tFVUkySU9kRzg2Nw0KQ3E2T1JvM1YxNnFhakFxY3hYQ3l4aHgzNmp3T1ZrTmxyL1V4VjNxWXE3MU1WZDZtS29yUzlYdjlLMUNEVU5QbmEydkxaZzhNeUdoQg0KSDZ3ZWhCNjVHVVJJVWVUS01qRTJIMVIrVm41bzJIblRUakZLRnQ5Y3RWQnZMVWZaWWRQVmlyK3lUMUg3SitnblNhblRIR2Y2THV0Tg0KcUJrSDlKbmVZcmt1eFYyS3V4VkE2M3E5dHBHbDNHb1hHNlFyVUpXaFpqc3FqNW5LYytZWW9HUjZOK2wwOHMyUVFqMVFubFJKdjBaNg0KOXp2ZTNEZXJkTi9sdUE5UDloeTRqMkdWYU1IZ3MvVWViZnJ5UEVxUDBqWWZqejVwem1XNExzVmRpcnNWZGlyNUkvNXlEL0w5dkxmbQ0KNTlXdEk2YVJyYk5QR1FQaGp1Q2F6UisxU2VhK3hwMnpmYUhQeHdvOHc2UFc0T0NkamtYbGVaemhPeFZWdDV6RTMrU2Vvd0VKQlRBUw0KMUZRYWc1Qms3MURpcnZVT0t1OVE0cTcxRGlxTzBUWGRUMFRWTGZWTk5tTUY3YXVIaWtIdDFWaDNWaHNRZW95TTRDUW84bVVKbUpzUA0KcjM4dC93QXdOTzg2NkFsOUJTSytocEhxTm5YZUtXblVWM0tQMVUvUjFCelFhakFjY3E2Tzl3WnhramZWbGVVTjdzVmRpcnpQejlxLw0KNlQ4eTJPZ3hOeXRyYVZHdWdPalNIY2ovQUdDZnJPYy8ybG44VExIRU9RTy80OXoxUFpPbjhMQkxNZnFrTnZkKzBzMjBHNURpV0lrVg0KRkhBN211eC9VTTIybW5kaDBXcmhWRk5zeW5EZGlyc1ZkaXJzVlNIeng1UDAzemY1YXU5RHZ4eFdjY29Kd0t0Rk11OGNpLzZwNmp1Sw0KanZsdUhLY2NoSU5XYkVKeDRTK0pQTWZsN1ZQTHV0M2VqYXBGNlY3WnZ3a0g3TERxcnFlNnNwREtmRE9reDVCT0lJNVBQWklHQm9wYg0KazJEc1ZWWVppbnduN1A2c0JDUVVUenlLVytXS3U1WXE3bGlydVdLc2k4aGVlTlQ4bmVZb05Xc3lYakg3dTh0YTBXYUFrY2tQZ2U2bg0Kc2NwejRSa2pSYmNPWTQ1V0gyYm9tdGFkcmVrMnVyYWRMNjFsZVJpV0YraG9lb0k3TXAySTdIT2VuQXhOSG05QkNZa0xISkc1RmtndA0KYTFPTFM5SnV0UWszVzNqTEJUKzAzUlYvMlRFREtzK1VZNEdSNk4rbXdITGtqQWRTOFU4dVRTM1BtSVhVN2M1V01rc2pIdXpBMVAzdA0Kbkk2YVJsbDRqejNlNTFzUkhCd2psc0hvV202bjlWdkk1U2ZncnhrSCtTZXY5YzNlSEx3eXQ1ek5nNDRrTTJWbFpReWtGU0tnamNFSA0KTnc2SWgyS3V4VjJLdXhWMkt2TXZ6dC9LZUx6bnBINlEwNUZYekhwNkg2c2RsK3NSajRqQXg4ZTZFOUQ3SE0zUjZyd3pSK2t1SHE5Tg0KNGdzZlVIeUpORE5CTkpCTWpSVFJNVWxqY0ZXVmxOR1ZnZHdRYzN3TnVpSXBaaFYyS3I0NUN1eDZZS1ZXREE5RGdaTjhzVmR5eFYzTA0KRlhjc1ZlMWY4NDRmbUdkTzFkdktkL0wvQUtEcVRGOVBaanRIZFUzUVY2Q1VEL2dnUEU1cnRmZ3NjWTVoMkdnejBlRThpK2w4MHp0Mw0KbjM1dTZxWTdHejB0RHZjT1pwdjlXUFpRZm14cjlHYVR0ck5VUkR2M2VqOW50UGM1WkQwMkh4L0gyc0Q4cnR4MUluL2l0djFqTkxwVA0KNjNvTmNMeC9GbHZyZStiSGlkUHdzcThxYTRzaWpUNTIrTmY5NTJQY2QxK2p0bXowV292MEg0T28xK2xyMWo0c216WXVyZGlyc1ZkaQ0KcnNWZGlyeFA4OS95WE91eHkrWi9Mc0ZkYWpISy9zMC80K1VVZmJSZjkrcUIwL2FIdjEyV2kxZkQ2WmNuWGF6U2NYcWp6Zk1MS3lzVg0KWUVNRFFnN0VFWnVYVHRZVmRpcmFzUjhzQ3FnWUhGTGRjVmRYRlhWeFZVdDdtZTJ1SXJpM2tNVThMckpGSXBveXVocXJBK0lJd0VYcw0Ka0duMjkrWG5tMkx6WjVQMDdXMW9KcDQrRjNHT2l6eG5oS0tlSElWSHNSbk41OFhCTXhlaXdaZU9BTHpiOHlML0FPdWViTG9BMVMxVg0KTGRQOWlLdC93N05uR2RxWk9MT2ZMWjcvQUxGeGNHbmovUzMvQUI4RW0wU1gwOVJqL3dBb012NFppWVRVbk4xVWJnV1QrdG1keE9xNA0KVzB1WFIxZEdLdXBCVmhzUVIwSXdpVktZQWlpOUI4cytZNDlUaDlHVWhiMklmR09nY0Q5b2Z4emU2VFZESUtQMVBPYTNSbkViSDBsUA0KTXpIQWRpcnNWZGlyc1ZkaXJ3Yjg4dnlOK3ZmV1BOUGxhMy8wN2VUVTlNakg5OTNhYUZSL3V6dXlqN1hVZkY5cmFhUFdWNlpjdWhkWg0Kck5IZnFqemZOcEZOajF6Y09wZGlyc1ZjQ1JpcTROZ1Z1dUtYVnhWMWNWZlFIL09Ldm1aeGNheDVhbGY0R1ZkUXRVUFlxUkZOOTRNZg0KM1pxdTBzZktYd2RuMmRrNXgrTHRZdURjNnRlM0o2elR5eWY4RTVPZVlaNWNVNUh2SmZYdFBEaHh4ajNSSDNJZUNReFRKSVAyR0IrNw0KSUEwYmJKUnNFTWxFb0lxRHNlbVp2RTZ2aGI5WEcxNFZTM3ZacmFkSjRITWNzWnFqanFEa281REUyT2JHZUlTRkViUFN2TFBtYTMxZQ0KRGc5STc2TWZ2WXV6RCtaZmI5V2REcE5XTW9yK0o1Zlc2RTRUWTNpVTh6TWNCMkt1eFYyS3V4VjJLdkNmenQvSW45Sk5jZVovS3NJWA0KVUtHWFVkTVFVOWM5VExDQi91eitaZjJ1M3hmYTJlajF0ZW1YTHZkYnE5SGZxano3bnpXeXNyRldCREEwSU94QkdiaDFEV0ZYWXE3Rg0KV3djQ3Q0cTdGV2Vma1pxemFiK2FPaVBVaU81a2Uwa0hpSjQyUlIvd1pVNWpheU40aTVXamxXUU01enlCOXVkaXFiMkZ6emdDay9FbQ0KeCtYYkw0UzJjUExDaWlmVXlkdFZPOVRHMXBVdDd1YTNtU2VCekhMR2FvNm1oQnlVWm1Kc2MyTThZa0tJc1BUdkt2bTYzMWVNVzg5SQ0KdFFRZkVuUVNBZnRKL0VaMFdqMW95aWp0TDczbGRmMmVjSnNidys1a1daN3JYWXE3RlhZcTdGWFlxOEwvQUQyL0phMHY0Ym56Wm9Lcg0KQnFTL0hmMklGRnVTU0J6anAwbDMzSDdYejY1dUh0T09HUDcwMUR2N25CejlueXltOFk5WGQzdm1sbFpXS3NDR0JvUWRpQ00zd0lJcw0KT2tJcHJDcnNWZGlyZGNWZFhGVTc4alN0RjUyOHZ5cDl1UFVyTmxyNGlkQ01xekQwUzl4Yk1QMWozaDY5UEVZcHBJbTZ4c1ZQelUweg0KeHlRbzArNVJsWUI3MW1Ca3EyMHhpa0Rmc25aaDdZWW1tRTQyRXpFbFJVSFk1ZGJpMDduamEwN25qYTB1aXVKSXBGbGljcEloREk2bQ0KaEJIUWc0UklnMkVTZ0NLUEo2ZDVQODZ4YW1Gc3I1aEhxQTJSdWl5Z2VIZzN0blJhSFhqSjZaZlY5N3kzYVBaaHhldUgwZmQreGxtYg0KTjA3c1ZkaXJzVmRpckNQUDJwTks4ZGhFMVVpL2VUZ2Z6RWZDUG9INjg1RDJoMW9NaGhIVGMrL3ArUE4zL1pHQ2dabnJ5ZUdmbUgrWA0KaTZrc21yYVRHQnFBSEs0dDEyRXdIN1MvNWY4QXhMNTVuZXpudEdjQkdITWYzWFEvemY4QWp2M2U1dyszT3cvRnZMaUg3enFQNTM3Zg0KdmVQc3JLeFZnUXdOQ0RzUVJucGdJSXNQQmtVMWhWMkt1eFYyS3A1NUZpYVh6dDVlaVQ3Y21wV2FyWHhNNkFaVm1Qb2w3aTJZZnJIdg0KRDNUempaR3k4ejZsQVJRZXUwaWovSmwvZUwrRFo1THJjZkJta1BQNzkzMmZzN0x4NmVCOHErV3lUWml1YTdGVVRiVDArQnVuN0p5VQ0KUzFUajFSUFBKdGRPNTRyVHVlSzA1WkNyQmxKREExQkd4QkdOb0llbCtTdlBTM3ZEVGRVY0xkMEN3WEIyRW4rUzMrWCt2NTllZzBIYQ0KUEg2Si9WMFBlOHYybjJWd1hQSDlQVWQzN1B1WnZtNGRFN0ZYWXFnZFoxU0xUYkY3aDkzK3pFbjh6bnAvYm1EMmhyWTZmRVpubjBIZQ0KWEkwMm5PV2ZDUGk4MW1sa21sZVdWaTBraExPeDdrN25QTjhtUXprWlMzSmVyakVSRkRrRUpQRHgrSmVuY2VHR01tMk1ubXY1a2ZsKw0KTDVKTlowcU1DOVFGcnUzVWYzd0g3YWdmdGp3L2ErZlh0L1pyMmg4RWpCbVA3cy9TZjV2a2Y2UDNlN2x6SGIzWXZpQTVzUTlmVWQvbg0KNy92OTd5SFBTWGhYWVZkaXJzVlozK1IybHRxUDVwYUZHQlZMZVY3cVErQWdqYVFIL2dsQXpGMWtxeEZ5ZEhHOG9lKy9tN3BSajFDMA0KMVJGK0M0VDBaU1A1NDkxcjgxUDRaNTMyMWhxUW4zN1BwbnM5cUxoTEdlbS96L0gydlBzMGowYnNWZGlxSmltNUNoNmpKQXRjb3IrVw0KRmpUdVdLMDdsaXRPNVlyVDByeUw1N0UvcDZWcXNuNy9BR1cxdW1QMi9CSFA4M2dlL3dBK3UvN083UnYwVDU5Qzh2MnIyVncza3hqYg0KcVAwaG4yYnQ1NVR1TGlHM2hlYVp3a1NDck1jcnk1WTQ0bVVqVVF5aEF5TkRtODYxeldKZFV2REthckNsVmhqOEY4VDdudm5uZmFYYQ0KRXRUazR2NFJ5SDQ2dlVhVFRERkd1dlZMczF6bE5kY1ZRczBmQnR2c25wbHNUYmJFMjhrL05IeU9JR2sxN1RZNlF1YTM4Sy9zc1QvZQ0KZ2VCUDJ2ZmZQUmZaWHQzaXJUWlR2L0FmOTcrcjVkenhmdEQyUnczbnhqYitJZnAvVzgwenUza1hZcTdGWHZYL0FEaW41ZWFYVnRaOA0Kd1NMKzd0b1Vzb0dQZDVtRWtsUGRWalgvQUlMTlgybmsyRVhaOW13M01udkhtN1JCck9nM05tb0JuQTlTMko3U3B1djMvWituT2QxdQ0KbjhYR1k5ZW52ZWo3UDFYZ1poTHB5UHVlQ3NyS3hWZ1ZaVFJsT3hCSFk1eGhENkVEYldLdXhWdnBpcXFrbFJ2MXcyd0lYY3NLS2R5eA0KV25jc1ZwM0xGYWVtK1J2UDBVc0g2TzFlWGpQQ3BNRjA1L3ZGWDlsdjhzZHZINTlkL29lMUlpTlpUWENPZjQ2L2U4dDJwMlNRZVBFTg0KanpIZCt6N2xubUR6RExxY3ZweDFqczBQd0ozWS93QXpaelhhdmFzdFRLaHRqSElkL21XelI2SVloWjNraythZHpuWXE3RlZycUdVZw0KNFFhU0RTQm1oUjBlR1ZRNk9Dcm93cUdVaWhCQjdFWmZDWkJCQm9oc0lFaFI1RjRGNTg4cXQ1ZTFwb29nVFlYRlpMTmpVMFd1NkUrSw0KSDhLWjdCMkIydCtjd1dmN3lPMHYxL0g3N2ZNKzJlenZ5dWFoOUV0NC9xK0RHODNycVhZcSsxdnljOG90NVcvTC9UYkdaT0Y5Y0tieQ0KK0JGQ0pwNkhpd1BkRUNvZmxuT2F2THg1Q2VqME9seGNHTURxelhNWnlIa2Y1bitXVFk2aitscmRQOUV2Vy9mVTZKUDFQL0I5Zm5YTw0KWjdXMG5CUGpIMHkrL3dEYTlqMkhydkVoNGN2cWp5OTM3R0Q1cUhmT3hWMkt0NHF2VjYvUEN4SWJyaWgxY1ZkWEZXd3hCQkJvUnVDTQ0KVlpKcE9wQzZqOU9RL3YwRy93RGxEeHpWNmpEd0d4eWRmbnhjSnNja3d6R2NkMkt1eFYyS3FGeWxSekhVZGNuRXM0RmlYNWhlWDExbg0KeTNjSWkxdXJVRzR0aU92SkI4Uy83SmFqNTB6b1BaN3REOHJxb2svUlAweStQWDRGMW5iZWkvTWFjZ2ZWSGNmRDlid0xQWW56SjZIKw0KUnZrYi9GZm5pMytzUjg5SzB1bDVmVkZWYmlmM1VSLzEzNmorVUhNVFdadUNIbVhLMGVIam41QjlrWnp6djNZcWhkVDAyMTFLd21zYg0KdGVjRTY4V0hjZHdSN2c3aks4dUtPU0pqTGtXM0JtbGltSng1aDRQNWcwSzgwVFU1TEc1RmVPOFVvRkZrUS9aWVp4dXAwOHNVekV2bw0KT2sxVWMrTVRqL1lVdHloeVhZcTdGWFlxdUJ4UlRkY1VPcmlycTRxcVFUeVFTckxHYU1wcU1qS0lrS0tKUkJGRmxscGRKYzI2ekowYg0KcVBBOXhtb3lRTVRSZFZrZ1ltbGZJTUhZcTdGV2lBUlE5RGlxQ1plTEVlR1hCdUQ1dTEzVGphK1liN1Q0RkwrbmRTUlFvZ3FTT1pDQQ0KQWQrbTJlNWRuWnpsMDJPWjV5Z0NmbHUrVGEzRDRlZWNCMGtSOXI3QS9KMzh2MDhsK1Q0YlNkUitscjBpNTFOeDFFakQ0WXErRVMvRA0KODZudm1wMWVmeEozMEhKMjJsd2VIQ3VyT2N4bkpkaXJzVlNQemI1V3RmTUduR0ZxUjNjVld0WjZmWmJ3UCtTM2ZNUFc2UVpvVi9FTw0KUmMvcy9YeTA4Ny9oUE1QRDc2eHVyQzdsdEx1TXhYRUxjWkVQWS8wemtjbU9VSkdNdVllOHhaWTVJaVVUWUtIeURZN0ZYWXE3Rlc2NA0Kb2JyaXJxNHE2dUtwbm9ONzZOejZMSDkzTnNQWnUzMzlNeGRWanVOOVE0MnB4M0crNWt1YTExenNWZGlyc1ZRdHlLU1Y4UmxrT1RiRA0Ka3gvOG1meThYVy9PdXArZU5SaTVhYmFYa3k2T3Jpb2xuUnl2ckR4V0tudy81WCtybnJ1REljV2t4NHY0dUNOL0xrK2VUeGpKcVo1Tw0KbkdhK2I2RXpIY3AyS3V4VjJLdXhWalBuWHliYjYvYWVwQ0ZqMU9FZnVKanNHSFgwMzl2QTlzMSt2MEl6UnNmV1B4VHRPek8wanA1VQ0KZDRIbVAwaDR2ZFd0emFYTWx0Y3h0RlBFeFdTTmhRZ2pPVW5BeE5IWWg3akhrak9JbEUyQ281Rm03RlhZcTdGVzhWZGlyc1ZiQklJSQ0KTkNOd2NVTXhzcmdYRnJGTjNkZmkrWTJQNDVwc2tlR1JEcU1rZUdSQ3ZrR0RzVmRpcXRZNkRjYTFjL1ZJcEdnUXFmV3VWNnhxZHFyWA0KYmwvTG0wN0owWno1UUs5STNQdS9hNCtyMUF4WXliM093ZWw2WnBsanBlblcrbldFS3dXZHBHc1VFSzlGVlJRZlAzSjY1Nk1TVHVYaw0KZ0FCUVJPQkxzVmRpcnNWZGlyc1ZZMTV5OGwybXYyeGxqQ3c2bkdQM00vWmdQMkpLZHZmdG12MTJoam1GamFmZit0Mm5admFjdFBLag0KdkE4eCtrUEdMNnh1N0M2a3RMdUpvYmlJMGVOdW8vc3psY21PVUpjTWhSZTN4Wlk1SWlVVFlLSHlEWTdGWFlxN0ZXOFZkaXJzVlpGNQ0KY2w1V2J4bjloOXZrd3pYYXVQcXQxK3JqNnJUZk1SeEhZcXEybHBQZDNDVzhDOHBaRFJSL0UrMlhZTUVzc3hDSXVSWVpNZ2hFeVBJUA0KU05JMHFEVGJOWUk5M084c25kbS96Nlo2Tm9OREhUWXhBYytwN3k4dHFkUWNzcktOek5jZDJLdXhWMkt1eFYyS3V4VjJLcEI1dDhvVw0KUG1DMCtJQ0svalVpM3VmRHZ4ZW5WVCtIYk1MV2FLT2FQZExvWFk5bjlvejA4dStCNWo4ZFhpdXE2VmZhWGV5V1Y3R1k1NCtvN0VkbQ0KVTl3YzVQTmhsamx3eUc3M0dEUERMQVNnYkJRbVZ0enNWZGlyc1ZkaXJzVlQzeXdUUzVIYjRQOEFqYk1IV2RIQjFuUlBjd1hDWFJ4dg0KSklzY2FsbmNnS28zSkp5VUlHUkFHNUtKU0FGbDZCNWQwRk5NdCtjbEd2SlIrOGIrVWZ5aitPZWdkazltRFRRcy93QjVMbjVlVHpXdA0KMVp5eW9mU0U0emJ1QzdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZVbTh6K1Y3RFg3RXdUZ0pjSUsyOXlCOFNOL0ZUM0dZdXIwa2MwYVBQbw0KWE4wT3VucDUyUHA2anZlSmF4bzJvYVJmUFozMFpqbFhkVy9aZGV6SWU0T2NqbndTeFM0WkI3dlRhbUdhSEZBN0lIS205Mkt1eFYySw0KdXhWa1hsdU1pMWxrL25lZy93QmlQN2MxMnNQcUFkZnF6NmdFNFZXWmdxZ2xpYUFEY2tuTVVBazBIRUpwbmZsbnk2dGpHTHE1RmJ4eA0Kc3AvM1dEMitmam5jOWpka2pBUEVuL2VIL1kvdC9zZWQxK3Q4UThNZnArOVA4MzdyWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcQ0Kay9tanl4WStZTEQ2dmNmQlBIVTIxd0JWbzJQNjFQY1ppNnZTUnpSbzgraGMzUTY2ZW5ueERsMUhlOE8xWFM3M1M3K1d4dkU0VHhHaA0KOENPektlNE9jaG13eXh5TVpjdzk1Z3p4eXdFNG5Zb1RLMjUyS3V4VnZyaXJNZE90VEJhUXcwK09ueEFmek51Znh6VVpKY2NpUTZqTg0KUGlrUzlBOHNlV3hhS0x5N1d0MDI4YUgvQUhXRC93QWJaMlhZM1pIaER4TWc5ZlFmemYydk42L1hjZnBqOVAzc2l6b25WdXhWMkt1eA0KVjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFZqUG5yeW5IcnVtbVNGUU5TdGxMVzdkM0hVeG41OXZmTmYyaG94bWhZK3NjdjFPMTdLNw0KUU9DZEg2SmMvd0JieEpsWldLc0NyS2FNcDJJSTdIT1NJZTZCdHJGWFlxbWVnNmUxMWRxNVVza1pCQUcvSi8yUm1QcUprRGhITXVOcQ0KY3ZERjYvNWI4c3BacXQzZUtHdXp1aUhjUi84QU4yZEgyUjJNTVFHVElQM25RZnpmMnZHNjdYbWZwajlQM3Npem9uVnV4VjJLdXhWMg0KS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFY0LythR2hMWWEydDlDdkdEVUFYWUFiQ1pmdC84RlVOODY1eS9hMm40TW5FT1V2dg0KZTA3QzFYaVl1QTg0ZmQwWVhtcWQycjJWbmNYbHpIYlc2RjVaQ0ZWUUs5Y0h1NXNNbVFRaVpIa0h0SGxIeWZCbzlzanpnUGQwcjRoQw0KZXY4QXN2Zk9pN003SzhJK0prM3lmN245dm04UjJqMmtjMGlJL1Q5N0pjM2pxbllxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWA0KWXE3RlhZcTdGV0dmbXZiSko1WlNZL2JndUVLbjJZTXBINDVxdTJJM2h2dUx2T3dKa1o2NzRsNUpaMmR6ZVhNZHJheHRMUEtlTWNhNw0Ka2s1ek1JR1JvYmt2WVpNa1lSTXBHZ0hzL2t6eVhiYURiQ2Fha3VwU3FQVmw2aFA4aFA0bnZuVWFEczRZZlZMZWYzZTU0anRMdE9Xbw0KbFEyZ1B0OTdKODJicW5ZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXF4UDhBTVcxdmRRMHkxMHF4aU10MQ0KZDNBTk95eHhnbG1ZOWdDVnpXZHFSbE9BaEVXWkYzSFkwNFk4a3NrelVZeCswL2dvdnlqNU9zZkwxdFVVbTFDVVV1TGtqNmVDZUMvcg0KL1Zab3RESEFPK1o1bjlUVDJqMmxQVVM3b0RrUDBuelpEbWU2MTJLdXhWMkt1eFYyS3V4VjJLdXhWMkt2LzlrPTwveG1wR0ltZzppbWFnZT4NCgkJCQkJPC9yZGY6bGk+DQoJCQkJPC9yZGY6QWx0Pg0KCQkJPC94bXA6VGh1bWJuYWlscz4NCgkJCTx4bXBNTTpJbnN0YW5jZUlEPnhtcC5paWQ6M2UxZjUyNzItYmVkMy00MjcyLTk0YzktY2QzZmM3MjMwMzRlPC94bXBNTTpJbnN0YW5jZUlEPg0KCQkJPHhtcE1NOkRvY3VtZW50SUQ+eG1wLmRpZDozZTFmNTI3Mi1iZWQzLTQyNzItOTRjOS1jZDNmYzcyMzAzNGU8L3htcE1NOkRvY3VtZW50SUQ+DQoJCQk8eG1wTU06T3JpZ2luYWxEb2N1bWVudElEPnV1aWQ6NUQyMDg5MjQ5M0JGREIxMTkxNEE4NTkwRDMxNTA4Qzg8L3htcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD4NCgkJCTx4bXBNTTpSZW5kaXRpb25DbGFzcz5wcm9vZjpwZGY8L3htcE1NOlJlbmRpdGlvbkNsYXNzPg0KCQkJPHhtcE1NOkRlcml2ZWRGcm9tIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4NCgkJCQk8c3RSZWY6aW5zdGFuY2VJRD54bXAuaWlkOmVhZmUxYTJkLWNiMTctNGMwYy05ZWVjLWE5ZGM0OTc1MWU3Nzwvc3RSZWY6aW5zdGFuY2VJRD4NCgkJCQk8c3RSZWY6ZG9jdW1lbnRJRD54bXAuZGlkOmVhZmUxYTJkLWNiMTctNGMwYy05ZWVjLWE5ZGM0OTc1MWU3Nzwvc3RSZWY6ZG9jdW1lbnRJRD4NCgkJCQk8c3RSZWY6b3JpZ2luYWxEb2N1bWVudElEPnV1aWQ6NUQyMDg5MjQ5M0JGREIxMTkxNEE4NTkwRDMxNTA4Qzg8L3N0UmVmOm9yaWdpbmFsRG9jdW1lbnRJRD4NCgkJCQk8c3RSZWY6cmVuZGl0aW9uQ2xhc3M+cHJvb2Y6cGRmPC9zdFJlZjpyZW5kaXRpb25DbGFzcz4NCgkJCTwveG1wTU06RGVyaXZlZEZyb20+DQoJCQk8eG1wTU06SGlzdG9yeT4NCgkJCQk8cmRmOlNlcT4NCgkJCQkJPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+DQoJCQkJCQk8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+DQoJCQkJCQk8c3RFdnQ6aW5zdGFuY2VJRD54bXAuaWlkOjJkOTVkNTU1LTg0OTItNDVmOC1hMmFjLTUyMzg4ZmUxMTI1Yzwvc3RFdnQ6aW5zdGFuY2VJRD4NCgkJCQkJCTxzdEV2dDp3aGVuPjIwMTYtMDktMDJUMTE6MDY6NTItMDQ6MDA8L3N0RXZ0OndoZW4+DQoJCQkJCQk8c3RFdnQ6c29mdHdhcmVBZ2VudD5BZG9iZSBJbGx1c3RyYXRvciBDQyAyMDE1IChNYWNpbnRvc2gpPC9zdEV2dDpzb2Z0d2FyZUFnZW50Pg0KCQkJCQkJPHN0RXZ0OmNoYW5nZWQ+Lzwvc3RFdnQ6Y2hhbmdlZD4NCgkJCQkJPC9yZGY6bGk+DQoJCQkJCTxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPg0KCQkJCQkJPHN0RXZ0OmFjdGlvbj5zYXZlZDwvc3RFdnQ6YWN0aW9uPg0KCQkJCQkJPHN0RXZ0Omluc3RhbmNlSUQ+eG1wLmlpZDozZTFmNTI3Mi1iZWQzLTQyNzItOTRjOS1jZDNmYzcyMzAzNGU8L3N0RXZ0Omluc3RhbmNlSUQ+DQoJCQkJCQk8c3RFdnQ6d2hlbj4yMDE4LTAyLTE0VDE2OjU5OjU5LTA1OjAwPC9zdEV2dDp3aGVuPg0KCQkJCQkJPHN0RXZ0OnNvZnR3YXJlQWdlbnQ+QWRvYmUgSWxsdXN0cmF0b3IgQ0MgMjIuMCAoTWFjaW50b3NoKTwvc3RFdnQ6c29mdHdhcmVBZ2VudD4NCgkJCQkJCTxzdEV2dDpjaGFuZ2VkPi88L3N0RXZ0OmNoYW5nZWQ+DQoJCQkJCTwvcmRmOmxpPg0KCQkJCTwvcmRmOlNlcT4NCgkJCTwveG1wTU06SGlzdG9yeT4NCgkJCTxpbGx1c3RyYXRvcjpTdGFydHVwUHJvZmlsZT5QcmludDwvaWxsdXN0cmF0b3I6U3RhcnR1cFByb2ZpbGU+DQoJCQk8cGRmOlByb2R1Y2VyPkFkb2JlIFBERiBsaWJyYXJ5IDE1LjAwPC9wZGY6UHJvZHVjZXI+DQoJCTwvcmRmOkRlc2NyaXB0aW9uPg0KCTwvcmRmOlJERj4NCjwveDp4bXBtZXRhPg0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICA8P3hwYWNrZXQgZW5kPSd3Jz8+/9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwECAgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgAEAAQAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAAwDAQACEQMRAD8A/Y79qjxDrPxt+FXxU+F/wt8aW3hL4tQ6AV02+dwjWdxNGWTaRl0yuwGZFYw/aEdQzKFr4r/4I9f8FnvF3xt+P3/CgfjVosmg+NLa0uF0y81G1lsNUmuLdfNewu7ZkwJhbiaUSnywUt8FWdgzfSH7fX7EuvfEnxLD8RPhzM1r4us4l+3W0Ny1rcah5QAilt5VI2XCLleSNyhBuXb81H9lD9ipvHvxf8M/HT4oeGbWL4jaLaTrpt1dKyamjTxNC7zBSBs8mSRVhlDeX5rbVjwM/NYHjTF4fMZZBmGWyqRqSk4V4NWhFJuMpXSur8sZR5lZ3cU20pfSZxwjhKmX0M9yzHxi4qMalGXxOWnMkk3o/elGVrWsm97f/9k='
$iconBytes = [Convert]::FromBase64String($iconBase64)
$stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }

# Create main form
$main_form = New-Object System.Windows.Forms.Form

$main_form.Text ='IP Checker'

$main_form.Width = 1200

$main_form.Height = 900

$main_form.BackColor = 'DarkBlue'

$main_form.StartPosition = 'CenterScreen'

$main_form.AutoScale = $true

$main_form.AutoScaleMode = "Font"

$main_form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

$main_form.MaximizeBox = $false

$main_form.FormBorderStyle='FixedDialog' # This line makes the box unable to be scaled (easier to work with when designing; make it scaleable later)


# Creating a browse button for the form 

$BrowseButton1 = New-Object System.Windows.Forms.Button
$BrowseButton1.Location = New-Object System.Drawing.Size(50,50)
$BrowseButton1.Size = New-Object System.Drawing.Size(120,40)
$BrowseButton1.Text = "Browse"
$BrowseButton1.BackColor = "White"
$BrowseButton1.Font = "Verdana,11"
$BrowseButton1.Add_Click({FileDialogInput})
$BrowseButton1.Visible = $false
$main_form.Controls.Add($BrowseButton1)

# Creating a second browse button for the form 

$BrowseButton2 = New-Object System.Windows.Forms.Button
$BrowseButton2.Location = New-Object System.Drawing.Size(50,150)
$BrowseButton2.Size = New-Object System.Drawing.Size(120,40)
$BrowseButton2.Text = "Browse"
$BrowseButton2.BackColor = "White"
$BrowseButton2.Font = "Verdana,11"
$BrowseButton2.Visible = $false
$BrowseButton2.Add_Click({FileDialogOutput})
$main_form.Controls.Add($BrowseButton2)

# Creating label for input path 

$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "Input path:"
$Label1.BackColor = "Orange"
$Label1.Font = "Verdana,11"
$Label1.Location = New-Object System.Drawing.Point(200,10)
$Label1.AutoSize = $true
$Label1.Visible = $false
$main_form.Controls.Add($Label1)

# Creating label for output path 

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Output path:"
$Label2.BackColor = "Orange"
$Label2.Font = "Verdana,11"
$Label2.Location = New-Object System.Drawing.Point(200,110)
$Label2.AutoSize = $true
$Label2.Visible = $false
$main_form.Controls.Add($Label2)

# Creating a text file button for the form 

$TxtFileBtn = New-Object System.Windows.Forms.Button
$TxtFileBtn.Location = New-Object System.Drawing.Size(550,300)
$TxtFileBtn.Size = New-Object System.Drawing.Size(120,70)
$TxtFileBtn.Text = "Check Text File"
$TxtFileBtn.BackColor = "White"
$TxtFileBtn.Font = "Verdana,11"
$TxtFileBtn.Add_Click({TextFileOnClick})
$main_form.Controls.Add($TxtFileBtn)

# Creating a single ip button for the form 

$SingleIPBtn = New-Object System.Windows.Forms.Button
$SingleIPBtn.Location = New-Object System.Drawing.Size(250,300)
$SingleIPBtn.Size = New-Object System.Drawing.Size(120,70)
$SingleIPBtn.Text = "Check Single IP"
$SingleIPBtn.BackColor = "White"
$SingleIPBtn.Font = "Verdana,11"
$SingleIPBtn.Add_Click({SingleIPOnClick})
$main_form.Controls.Add($SingleIPBtn)

# Creating a bulk manual button for the form 

$ManualIPBtn = New-Object System.Windows.Forms.Button
$ManualIPBtn.Location = New-Object System.Drawing.Size(850,300)
$ManualIPBtn.Size = New-Object System.Drawing.Size(120,70)
$ManualIPBtn.Text = "Check Bulk Manually"
$ManualIPBtn.BackColor = "White"
$ManualIPBtn.Font = "Verdana,11"
$ManualIPBtn.Add_Click({ManualIPOnClick})
$main_form.Controls.Add($ManualIPBtn)

# Creating a run button for the form 

$RunBtn = New-Object System.Windows.Forms.Button
$RunBtn.Location = New-Object System.Drawing.Size(50,400)
$RunBtn.Size = New-Object System.Drawing.Size(70,70)
$RunBtn.Text = "Run"
$RunBtn.BackColor = "White"
$RunBtn.Font = "Verdana,11"
$RunBtn.Add_Click({Run})
$main_form.Controls.Add($RunBtn)

# Creating a rich text box for file path input

$inputPathBox = New-Object System.Windows.Forms.RichTextBox 
$inputPathBox.Location = New-Object System.Drawing.Size(200,50) 
$inputPathBox.Size = New-Object System.Drawing.Size(900,40)
$inputPathBox.Font = "Verdana, 11"
$inputPathBox.ReadOnly = $True
$inputPathBox.MultiLine = $True
$inputPathBox.ScrollBars = "Vertical"
$inputPathBox.Anchor = 'top, left'
$inputPathBox.Visible = $false
$main_form.Controls.Add($inputPathBox)

# Creating a rich text box for file path output

$outputPathBox = New-Object System.Windows.Forms.RichTextBox 
$outputPathBox.Location = New-Object System.Drawing.Size(200,150) 
$outputPathBox.Size = New-Object System.Drawing.Size(900,40)
$outputPathBox.Font = "Verdana, 11"
$outputPathBox.ReadOnly = $True
$outputPathBox.MultiLine = $True
$outputPathBox.ScrollBars = "Vertical"
$outputPathBox.Anchor = 'top, left'
$outputPathBox.Visible = $false
$main_form.Controls.Add($outputPathBox)

# Creating label for output path (SINGLE IP)

$Label2Single = New-Object System.Windows.Forms.Label
$Label2Single.Text = "Output path:"
$Label2Single.BackColor = "Orange"
$Label2Single.Font = "Verdana,11"
$Label2Single.Location = New-Object System.Drawing.Point(200,10)
$Label2Single.AutoSize = $true
$Label2Single.Visible = $false
$main_form.Controls.Add($Label2Single)

# Creating label for Single IP Input path 

$SingleIPInputLabel = New-Object System.Windows.Forms.Label
$SingleIPInputLabel.Text = "Single IP Input:"
$SingleIPInputLabel.BackColor = "Orange"
$SingleIPInputLabel.Font = "Verdana,11"
$SingleIPInputLabel.Location = New-Object System.Drawing.Point(200,110)
$SingleIPInputLabel.AutoSize = $true
$SingleIPInputLabel.Visible = $false
$main_form.Controls.Add($SingleIPInputLabel)

# Creating a rich text box for file path output (FOR SINGLE IP)

$outputPathBoxSingle = New-Object System.Windows.Forms.RichTextBox 
$outputPathBoxSingle.Location = New-Object System.Drawing.Size(200,50) 
$outputPathBoxSingle.Size = New-Object System.Drawing.Size(900,40)
$outputPathBoxSingle.Font = "Verdana, 11"
$outputPathBoxSingle.ReadOnly = $True
$outputPathBoxSingle.MultiLine = $True
$outputPathBoxSingle.ScrollBars = "Vertical"
$outputPathBoxSingle.Anchor = 'top, left'
$outputPathBoxSingle.Visible = $false
$main_form.Controls.Add($outputPathBoxSingle)

# Creating a rich text box for single IP input (FOR SINGLE IP)

$SingleIPInputBox = New-Object System.Windows.Forms.RichTextBox 
$SingleIPInputBox.Location = New-Object System.Drawing.Size(200,150) 
$SingleIPInputBox.Size = New-Object System.Drawing.Size(900,40)
$SingleIPInputBox.Font = "Verdana, 11"
$SingleIPInputBox.ReadOnly = $False
$SingleIPInputBox.MultiLine = $False
$SingleIPInputBox.ScrollBars = "Vertical"
$SingleIPInputBox.Anchor = 'top, left'
$SingleIPInputBox.Visible = $false
$main_form.Controls.Add($SingleIPInputBox)

# Creating a second browse button for the form (SINGLE IP) 

$BrowseButton2Single = New-Object System.Windows.Forms.Button
$BrowseButton2Single.Location = New-Object System.Drawing.Size(50,50)
$BrowseButton2Single.Size = New-Object System.Drawing.Size(120,40)
$BrowseButton2Single.Text = "Browse"
$BrowseButton2Single.BackColor = "White"
$BrowseButton2Single.Font = "Verdana,11"
$BrowseButton2Single.Visible = $false
$BrowseButton2Single.Add_Click({FileDialogOutput})
$main_form.Controls.Add($BrowseButton2Single)

# Creating a rich text box for manual IP input (FOR manual IP)

$manualInputBox = New-Object System.Windows.Forms.RichTextBox 
$manualInputBox.Location = New-Object System.Drawing.Size(200,150) 
$manualInputBox.Size = New-Object System.Drawing.Size(900,40)
$manualInputBox.Font = "Verdana, 11"
$manualInputBox.ReadOnly = $False
$manualInputBox.MultiLine = $False
$manualInputBox.ScrollBars = "Vertical"
$manualInputBox.Anchor = 'top, left'
$manualInputBox.Visible = $false
$main_form.Controls.Add($manualInputBox)

# Creating an add button for the form (for manual ip)

$AddButton = New-Object System.Windows.Forms.Button
$AddButton.Location = New-Object System.Drawing.Size(110,150)
$AddButton.Size = New-Object System.Drawing.Size(60,40)
$AddButton.Text = "Add"
$AddButton.BackColor = "White"
$AddButton.Font = "Verdana,10"
$AddButton.Visible = $false
$AddButton.Add_Click({AddClick})
$main_form.Controls.Add($AddButton)

# Creating an clear button for the form (for manual ip)

$ClearButton = New-Object System.Windows.Forms.Button
$ClearButton.Location = New-Object System.Drawing.Size(50,150)
$ClearButton.Size = New-Object System.Drawing.Size(60,40)
$ClearButton.Text = "Clear"
$ClearButton.BackColor = "White"
$ClearButton.Font = "Verdana,10"
$ClearButton.Visible = $false
$ClearButton.Add_Click({ClearClick})
$main_form.Controls.Add($ClearButton)

# Creating label for Manual IP Input box 

$ManualIPInputLabel = New-Object System.Windows.Forms.Label
$ManualIPInputLabel.Text = "IP Input:"
$ManualIPInputLabel.BackColor = "Orange"
$ManualIPInputLabel.Font = "Verdana,11"
$ManualIPInputLabel.Location = New-Object System.Drawing.Point(200,110)
$ManualIPInputLabel.AutoSize = $true
$ManualIPInputLabel.Visible = $false
$main_form.Controls.Add($ManualIPInputLabel)

# Creating a rich text box for overall output

$outputBox = New-Object System.Windows.Forms.RichTextBox 
$outputBox.Location = New-Object System.Drawing.Size(50,500) 
$outputBox.Size = New-Object System.Drawing.Size(1100,300)
$outputBox.Font = "Verdana, 11"
$outputBox.ReadOnly = $True
$outputBox.MultiLine = $True
$outputBox.ScrollBars = "Vertical"
$outputBox.Anchor = 'top, left'
$main_form.Controls.Add($outputBox)

# Displays the form on screen
$main_form.ShowDialog()

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Regex for checking if an IP Address is valid 
$regex = '^((25[0-5]|(2[0-4]|1\d|[1-9]|)\d)\.?\b){4}$'

# Function for Bulk Lookup with a txt file
function TxtBulkLookupIPAddresses($filePath, $outputFile) {
    try {
        # Read ip addresses and then get rid of whitespaces from input file 
        $cleanedContent = (Get-Content -Path $filePath) -replace '\s+', ''

        $cleanedContent | Set-Content -Path $filePath

        # Assign ipAddresses variable here instead of doing it at the start, as now the file is cleansed of all whitespaces 
        $ipAddresses = Get-Content -Path $filePath

        # Initialize an array to store results
        $results = @()

        foreach ($ipAddress in $ipAddresses) {
            try {
                $ipValid = [String][bool]($ipAddress -match $regex) 

                # Resolve the hostname
                $hostname = [System.Net.Dns]::GetHostEntry($ipAddress).HostName

                # Create an object with the result
                $resultObject = [PSCustomObject]@{
                    IPAddress = $ipAddress
                    Hostname  = $hostname
                }

                # Add the result to the array
                $results += $resultObject

                # Display the result
                Write-Output "`nPing to $ipAddress successful. Hostname: $hostname"
            } catch {

                if ( $ipValid -eq "False" ) { 
                Write-Output "`nPing to $ipAddress failed. Error: Invalid IP Format" 
                $hostname = "Invalid IP Format"
                }

                else { 
                Write-Output "`nPing to $ipAddress failed. Error: $_" 
                $hostname = $_ 
                }

                $resultObject = [PSCustomObject]@{
                    IPAddress = $ipAddress
                    Hostname  = $hostname
                }

                $results += $resultObject
            }
        }

        # Export results to CSV
        $results | Export-Csv -Path $outputFile -NoTypeInformation

        Write-Output "`nResults exported to $outputFile"
    } catch {
        Write-Error "Error: $_"
    }
}

# Function for Bulk Lookup with a manual input
function ManualBulkLookupIPAddresses($outputFile, $inputArray) {
    try {
        $results = @()

        foreach ($ipAddress in $inputArray) {
            try {
                # Resolve the hostname
                $hostname = [System.Net.Dns]::GetHostEntry($ipAddress).HostName

                # Create an object with the result
                $resultObject = [PSCustomObject]@{
                    IPAddress = $ipAddress
                    Hostname  = $hostname
                }

                # Add the result to the array
                $results += $resultObject

                # Display the result
                Write-Output "`nPing to $ipAddress successful. Hostname: $hostname"
            } catch {
                 $ipValid = [String][bool]($ipAddress -match $regex) 

                if ( $ipValid -eq "False" ) { 
                Write-Output "`nPing to $ipAddress failed. Error: Invalid IP Format" 
                $hostname = "Invalid IP Format"
                }

                else { 
                Write-Output "`nPing to $ipAddress failed. Error: $_" 
                $hostname = $_ 
                }

                $resultObject = [PSCustomObject]@{
                    IPAddress = $ipAddress
                    Hostname  = $hostname
                }

                $results += $resultObject
            }
        }

        # Export results to CSV
        $results | Export-Csv -Path $outputFile -NoTypeInformation

        Write-Output "`nResults exported to $outputFile"
    } catch {
        Write-Error "Error: $_"
    }
}



# Function for one singular lookup (NOTE: this one doesn't have whitespace deletion since it is unlikely for a user to need to put in a whitespace as part of their input)
function SingleLookupIPAddress($outputFile, $ipAddress) { 
       try {
           # Resolve the hostname
           $hostname = [System.Net.Dns]::GetHostEntry($ipAddress).HostName # more attributes can be found using this

           # Create an object with the result
           $resultObject = [PSCustomObject]@{
                 IPAddress = $ipAddress
                 Hostname  = $hostname
           }

           # Display the result
           Write-Output "`nPing to $ipAddress successful. Hostname: $hostname"
       } catch {
           $ipValid = [String][bool]($ipAddress -match $regex) 

           if ($ipValid -eq $false) {
           Write-Output "`nPing to $ipAddress failed. Error: Invalid IP Format"
                
           $resultObject = [PSCustomObject]@{
               IPAddress = $ipAddress
               Hostname  = "Invalid IP Format"
           }
           }
           else {

           Write-Output "`nPing to $ipAddress failed. Error: $_"
                
           $resultObject = [PSCustomObject]@{
               IPAddress = $ipAddress
               Hostname  = $_
           }

           }

           $results += $resultObject
       }

       try {
       # Export result to CSV
       $resultObject | Export-Csv -Path $outputFile -NoTypeInformation

       Write-Output "`nResults exported to $outputFile"
       } catch {
            Write-Error "`nError: $_"
       }
}
