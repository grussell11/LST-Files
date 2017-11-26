' Update History:
'
' 08-22-15: remove Multisensor and (Motion) from motion & sensor logs
' 09-05-15: fix door lock log; strip Multisensor (Motion) from motion summary
' 09-06-15: added 4 contact sensors
' 09-13-15: set date last changed virtual devices for all devices (if present)
' 09-13-15: temporarily disable lock log while locks not working
' 09-13-15: start new At Home log
' 09-14-15: format At Home log message
' 09-22-15: look at logs for the past week for At Home logs
' 10-30-15: added two light to Lights and one sensor to Sensors
' 10-31-15: add new appliance & detector groups; re-enable lock log
' 11-01-15: add new appliance & detector summary and event logs
' 11-14-15: flip LOCKED/UNLOCKED logic for door lock
' 11-21-15: fix logic determining if detector is OK
' 11-21-15: add alarm trigger logic
' 11-28-15: delay setting Alarm Triggered virtual device
' 02-28-16: reversed LOCKED/UNLOCKED logic for door lock
' 03-11-16: fixed Lock script logic and also changed STANDARD to GARY for code 3
' 04-02-16: added new stairwell light
' 04-08-16: updated alarm logic
' 08-07-16: commented out print statements on weather and corrected door unlock logic
' 10-01-16: update lock log translations for supporting different lock types
' 10-08-16: update for alarm "insurance"
' 10-09-16: add alarm activity log message
' 12-22-16: add thermostat event log
' 12-26-16: assign virtual thermostat log device, set temp adjustments, and add date/time
' 03-25-17: re-work alarm to add Night & Away modes, and use new Alarm State virtual device
' 05-06-17: re-work logic to get logs due to Homeseer function crashes for certain versions
' 07-29-17: add LogPinAttempt() and pinTable for pin validation and logging
' 07-29-17: add support for PinValidatedVirtualDevice
' 11-19-17: adjust lock message parsing

Imports System.Net
Imports System.Text
Imports System.IO

Public Sub Main(ByVal params As Object)
End Sub


'===================================================================
' BEGIN Home-specific device names defined here
'===================================================================
Public lightDeviceNames() As String = { _
    "Upstairs Bonus Room Left Light", _
    "Upstairs Bonus Room Desk Lamp", _
    "Upstairs Bonus Room Right Light", _
    "Upstairs Hallway Light - Master", _
    "Downstairs Den Den Light", _
    "Downstairs Dining Room Dining Room Light", _
    "Downstairs Hallway Hallway Light", _
    "Downstairs Living Room Front Porch Light", _
    "Downstairs Garage Outside Lights", _
    "Downstairs Kitchen Kitchen Light"}

Public motionDeviceNames() As String = { _
    "Upstairs Bonus Room Multisensor (Motion)", _
    "Downstairs Den Motion Sensor (Motion)", _
    "Downstairs Living Room Motion Sensor (Motion)"}

Public sensorDeviceNames() As String = { _
    "Downstairs Den Rear Door (Motion)", _
    "Downstairs Back Door (Motion)", _
    "Downstairs Dining Room Left Window (Motion)", _
    "Downstairs Dining Room Right Window (Motion)", _
    "Downstairs Den North Window (Motion)", _
    "Downstairs Den Left Window (Motion)", _
    "Downstairs Den Slider (Motion)", _
    "Downstairs Outside Mailbox - Motion", _
    "Downstairs Living Room French Door (Motion)", _
    "Downstairs Living Room Front Door (Motion)", _
    "Upstairs Bonus Room Printer Window (Motion)", _
    "Upstairs Bonus Room Left Window (Motion)", _
    "Upstairs Bonus Room Right Window (Motion)"}

Public applianceDeviceNames() As String = {}

Public detectorDeviceNames() As String = { _
    "Downstairs Dining Room Smoke Detector (Notification)", _
    "Upstairs Hallway Smoke Alarm (Notification)", _
    "Downstairs Hallway Smoke Detector (Notification)", _
    "Upstairs Bedroom Water Sensor (Alarm)", _
    "Downstairs Garage Water Detector (Alarm)"}

Public thermostatDeviceNames() As String = { _
    "Downstairs Hallway"}

Public lockDeviceNames() As String = { _
    "Downstairs Living Room Front Door Lock", _
    "Downstairs Garage Back Door"}

Public lockCodeTable() As String = { _
    "0,UNLOCKED", _
    "5,UNLOCKED by Claire", _
    "4,UNLOCKED by Linda", _
    "3,UNLOCKED by Gary", _
    "255,LOCKED"}

Public pinTable() As String = { _
    "2256,Jenny", _
    "8412,Alex", _
    "7113,Frank", _
    "5562,Lisa"}

' PIN log message type
Public PinSuccessLogMessageType As String = "PIN Success"
Public PinFailureLogMessageType As String = "PIN Failure"
'===================================================================
' END
'===================================================================

'===================================================================
' BEGIN Virtual device names defined here
'===================================================================
Public lightSummaryVirtualDeviceName As String = "Light Summary Status"
Public motionSummaryDeviceVirtualDeviceName As String = "Motion Summary Device"
Public motionSummaryTimeVirtualDeviceName As String = "Motion Summary Time"
Public sensorSummaryVirtualDeviceName As String = "Sensor Summary Status"
Public applianceSummaryVirtualDeviceName As String = "Appliance Summary Status"
Public detectorSummaryVirtualDeviceName As String = "Detector Summary Status"
Public logVirtualDeviceName As String = "Event Log"
Public lightLogVirtualDeviceName As String = "Light Event Log"
Public motionLogVirtualDeviceName As String = "Motion Event Log"
Public sensorLogVirtualDeviceName As String = "Sensor Event Log"
Public lockLogVirtualDeviceName As String = "Lock Event Log"
Public applianceLogVirtualDeviceName As String = "Appliance Event Log"
Public detectorLogVirtualDeviceName As String = "Detector Event Log"
Public thermostatLogVirtualDeviceName As String = "Temperature Event Log"
Public atHomeLogVirtualDeviceName As String = "At Home"
Public pinSuccessLogVirtualDeviceName As String = "Pin Success Log"
Public pinFailureLogVirtualDeviceName As String = "Pin Failure Log"
Public pinValidatedVirtualDeviceName As String = "Pin Validated"
'===================================================================
' END
'===================================================================

'===================================================================
' BEGIN Alarm virtual device names, constants, and options
'===================================================================

' Alarm virtual device name and valid constant values 
Public alarmVirtualDeviceName As String = "Alarm"
Public ALARM_OFF As Integer = 0
Public ALARM_NIGHT_MODE As Integer = 50
Public ALARM_AWAY_MODE As Integer = 100

' Alarm State virtual device name and valid constant values
Public alarmStateVirtualDeviceName As String = "Alarm State"
Public ALARM_STATE_READY As Integer = 0
Public ALARM_STATE_PAUSED As Integer = 25
Public ALARM_STATE_PENDING As Integer = 50
Public ALARM_STATE_TRIGGERED As Integer = 100

' Alarm buffer time (seconds)
'   - Allowed time after arming before noticing device changes
'   - Allowed time after pending to shut off alarm before triggering
Public ALARM_BUFFER_TIME_SEC As Integer = 10

' Optional Alarm Night mode Pause support
'   Pause detection motion device name (set it to "" if pause not desired)
Public alarmNightModePauseMotionDeviceName = "Upstairs Bonus Room Multisensor (Motion)"
'   Pause time (minutes)
Public ALARM_PAUSED_TIME_MIN As Integer = 10

'===================================================================
' END
'===================================================================


Public CRLF As String = Chr(13) & Chr(10)

'===================================================================
'===================================================================
' Call on "IF Any Device has a value that just changed" event
'===================================================================
'===================================================================
Public Sub OnDeviceValueChange(ByVal params As Object)
    CheckAlarm()

    SetLightSummaryStatus()
    SetMotionSummaryStatus()
    SetSensorSummaryStatus()
    SetApplianceSummaryStatus()
    SetDetectorSummaryStatus()

    UpdateLogHistory()
    UpdateAtHomeLogHistory()
End Sub

'===================================================================
' OnWeatherUpdate()

' Get current temperature from OpenWeatherMap.org web service for 
' specified zipcode and set "Current Temperature" virtual device
'===================================================================
Public Sub OnWeatherUpdate(ByVal params As Object)
    Dim url As String = "http://api.openweathermap.org/data/2.5/weather?zip=92807,us&appid=5848c157fbd05e7820b5328c35274105"
    Try
        Dim request As WebRequest = WebRequest.Create(url)
        request.Credentials = CredentialCache.DefaultCredentials
        Dim response As WebResponse = request.GetResponse()
        Dim streamReader As StreamReader = New StreamReader(response.GetResponseStream())
        Dim responseText As String = streamReader.ReadToEnd()
        hs.WriteLog("LiveSmart", url)
        hs.WriteLog("LiveSmart", responseText)

        Dim responseSplit() As String = responseText.Split(New Char() {","c, "{"c})
        Dim n As Integer
        For n = 0 To responseSplit.Length - 1
        Dim search As String = """temp"":"
        Dim index As Integer = responseSplit(n).IndexOf(search, StringComparison.OrdinalIgnoreCase)
        If index >= 0 Then
            Dim currentTemp As Double = Val(responseSplit(n).Substring(index + search.Length))
            currentTemp = currentTemp * 9 / 5 - 459.67 
            hs.SetDeviceStringByName("Current Temperature", currentTemp.ToString("N1") + " F", True)
            hs.SetDeviceValueByName("Current Temperature", currentTemp)
            Exit For
        End If
        Next n
    Catch e As Exception
        hs.WriteLog("LiveSmart", e.Message)
    End Try
End Sub

'===================================================================
' LogPinAttempt(pin, action)
'
'   pin: from PIN text box to validate
'   action: action (Check In, Check Out, etc.) to log
'===================================================================
Public Sub LogPinAttempt(ByVal params() As String)
    Dim pin As String
    pin = params(0)
    Dim action As String
    action = params(1)

    Dim pinTableParts() As String
    Dim person As String = ""

    Dim n As Integer

    Try
        ' Look for pin in pinTable
        For n = 0 To pinTable.Length - 1
            pinTableParts = pinTable(n).Split(New Char() {","c})
            If pinTableParts.Length >= 2 Then
                If pinTableParts(0) = pin Then
                    person = pinTableParts(1)
                    Exit For
                End If
            End If
        Next n

        ' Log in log file successful or unsuccessful PIN entry
        If person = "" Then
            hs.WriteLog(PinFailureLogMessageType, "Invalid PIN Entered for " + action)
            hs.SetDeviceStringByName(pinValidatedVirtualDeviceName, "Invalid Entry", True)
        Else
            hs.WriteLog(PinSuccessLogMessageType, person + " " + action + " Success")
            hs.SetDeviceStringByName(pinValidatedVirtualDeviceName, "Valid Entry - Thank You", True)
        End If
        UpdateLogHistory()
    Catch e As Exception
        hs.WriteLog("LiveSmart", e.Message)
    End Try
End Sub

'===================================================================
Private Function GetDevicesWithActivityCount(ByVal deviceNames() As String, ByVal alarmSetTime As Date) As Integer
    Dim devicesWithActivityCount As Integer
    devicesWithActivityCount = 0

    Try
        Dim deviceCount As Integer
        deviceCount = deviceNames.Length

        Dim n As Integer
        For n = 0 To deviceCount - 1
            Dim deviceRef As Integer
            deviceRef = hs.GetDeviceRefByName(deviceNames(n))
            If deviceRef > 0 Then
                Dim lastChanged As Date
                lastChanged = hs.DeviceLastChangeRef(deviceRef)
                If lastChanged > alarmSetTime Then
                    devicesWithActivityCount = devicesWithActivityCount + 1
                    hs.WriteLog("LiveSmart", String.Format("Activity reported from {0} at {1}", deviceNames(n), lastChanged.ToString("hh:mm:ss tt")))
                End If
            End If
        Next n    
    Catch e As Exception
        hs.WriteLog("LiveSmart",  "GetDevicesWithActivityCount(): " + e.Message)
    End Try
    
    Return devicesWithActivityCount
End Function

'===================================================================
Private Sub CheckAlarm()
    Try    
        Dim alarmDeviceRef As Integer
        alarmDeviceRef = hs.GetDeviceRefByName(alarmVirtualDeviceName)

        Dim alarmStateDeviceRef As Integer
        alarmStateDeviceRef = hs.GetDeviceRefByName(alarmStateVirtualDeviceName)

        ' Both Alarm and Alarm State virtual devices must exist for alarm to work
        If alarmDeviceRef > 0 And alarmStateDeviceRef > 0 Then
            Dim alarmValue As Integer
            alarmValue = hs.DeviceValue(alarmDeviceRef)

            Select Case alarmValue
                Case ALARM_OFF

                Case ALARM_AWAY_MODE, ALARM_NIGHT_MODE
                    Dim alarmStateValue As Integer
                    alarmStateValue = hs.DeviceValue(alarmStateDeviceRef)

                    Select Case alarmStateValue
                        Case ALARM_STATE_READY
                            ' Check if there has been a device change since Alarm was set + ALARM_BUFFER_TIME_SEC
                            Dim alarmSetTime As Date
                            alarmSetTime = hs.DeviceLastChangeRef(alarmDeviceRef)
                            alarmSetTime = alarmSetTime.AddSeconds(ALARM_BUFFER_TIME_SEC)
                            
                            ' Check sensor devices for both ALARM_AWAY_MODE and ALARM_NIGHT_MODE
                            Dim activityDetectedDeviceCount As Integer
                            activityDetectedDeviceCount = GetDevicesWithActivityCount(sensorDeviceNames, alarmSetTime)
                            
                            ' Check motion devices only for ALARM_AWAY_MODE
                            If alarmValue = ALARM_AWAY_MODE Then
                                activityDetectedDeviceCount = activityDetectedDeviceCount + GetDevicesWithActivityCount(motionDeviceNames, alarmSetTime)
                            End If
                        
                            ' If change has occurred, set state to ALARM_STATE_PENDING
                            If activityDetectedDeviceCount > 0 Then
                                hs.WriteLog("LiveSmart", String.Format("Alarm is PENDING set by {0} devices showing activity", activityDetectedDeviceCount))
                                hs.SetDeviceValueByName(alarmStateVirtualDeviceName, ALARM_STATE_PENDING)
                            
                            ' Else check if ALARM_NIGHT_MODE pause motion device exists and has detected motion for PAUSED
                            ElseIf alarmValue = ALARM_NIGHT_MODE And alarmNightModePauseMotionDeviceName <> "" Then
                                Dim deviceRef As Integer
                                deviceRef = hs.GetDeviceRefByName(alarmNightModePauseMotionDeviceName)
                                If deviceRef > 0 Then
                                    Dim deviceValue As Integer
                                    deviceValue = hs.DeviceValue(deviceRef)
                                    Dim lastChanged As Date
                                    lastChanged = hs.DeviceLastChangeRef(deviceRef)
                                    If deviceValue <> 0 And lastChanged > alarmSetTime Then
                                        ' Motion detected with ALARM_NIGHT_MODE pause motion device, set state to ALARM_STATE_PAUSED
                                        hs.WriteLog("LiveSmart", String.Format("Alarm is PAUSED"))
                                        hs.SetDeviceValueByName(alarmStateVirtualDeviceName, ALARM_STATE_PAUSED)
                                    End If
                                End If
                            End If

                        Case ALARM_STATE_PENDING
                            ' Alarm is pending, check if it has been on for more than ALARM_BUFFER_TIME_SEC seconds
                            Dim alarmPendingOverTime As Date
                            alarmPendingOverTime = hs.DeviceLastChangeRef(alarmStateDeviceRef)
                            alarmPendingOverTime = alarmPendingOverTime.AddSeconds(ALARM_BUFFER_TIME_SEC)
                            If Now > alarmPendingOverTime Then
                                ' AlarmPending time is up, set alarm state to TRIGGERED
                                hs.WriteLog("LiveSmart", String.Format("Alarm is TRIGGERED"))
                                hs.SetDeviceValueByName(alarmStateVirtualDeviceName, ALARM_STATE_TRIGGERED)
                            End If

                        Case ALARM_STATE_PAUSED
                            ' Alarm is paused, check if it has been on for more than ALARM_PAUSED_TIME_MIN minutes
                            Dim alarmPausedOverTime As Date
                            alarmPausedOverTime = hs.DeviceLastChangeRef(alarmStateDeviceRef)
                            alarmPausedOverTime = alarmPausedOverTime.AddMinutes(ALARM_PAUSED_TIME_MIN)
                            If Now > alarmPausedOverTime Then
                                ' Touch Alarm device to reset set time
                                hs.SetDeviceValueByName(alarmVirtualDeviceName, ALARM_OFF)
                                hs.SetDeviceValueByName(alarmVirtualDeviceName, ALARM_NIGHT_MODE)

                                ' AlarmPaused time is up, set alarm state to READY 
                                hs.WriteLog("LiveSmart", String.Format("Alarm is READY"))
                                hs.SetDeviceValueByName(alarmStateVirtualDeviceName, ALARM_STATE_READY)
                            End If

                        Case ALARM_STATE_TRIGGERED

                        Case Else
                            hs.WriteLog("LiveSmart", String.Format("CheckAlarm(): invalid Alarm State value ({0}) -- ignoring", alarmStateValue))
                    End Select

                Case Else
                    hs.WriteLog("LiveSmart", String.Format("CheckAlarm(): invalid Alarm value ({0}) -- ignoring", alarmValue))
            End Select
        End If
    Catch e As Exception
        hs.WriteLog("LiveSmart", "CheckAlarm(): " + e.Message)
    End Try
End Sub

'===================================================================
' Updates the last changed virtual device for a given device to a
' date & time the device was last updated.
Private Sub UpdateLastChangedVirtualDevice(ByVal deviceName As String)
    Dim virtualDeviceName As String = ""
    Dim virtualDeviceRef As Integer = 0

    virtualDeviceName = deviceName + " Last Changed"
    virtualDeviceRef = hs.GetDeviceRefByName(virtualDeviceName)

    If virtualDeviceRef > 0 Then
        Dim valueTimeLastChanged As Integer
        valueTimeLastChanged = hs.DeviceTimeByName(deviceName)

        Dim dateTimeLastChanged As DateTime
        dateTimeLastChanged = DateTime.Now.AddMinutes(-valueTimeLastChanged)

        Dim timeLastChanged As String
        timeLastChanged = dateTimeLastChanged.ToString("ddd MM/dd hh:mm tt")

        hs.SetDeviceValueByName(virtualDeviceName, valueTimeLastChanged)
        hs.SetDeviceStringByName(virtualDeviceName, timeLastChanged, True)
    End If
End Sub

'===================================================================
' Sets the status summary virtual device to a string reflecting the
' current status of the specified light devices.
Private Sub SetLightSummaryStatus()
    Dim summary As String = ""
    Dim summaryValue As Integer = 0
    Dim devicesOn As Integer = 0
    Dim deviceCount As Integer
    deviceCount = lightDeviceNames.Length

    Dim n As Integer
    For n = 0 To deviceCount - 1
        Dim deviceRef As Integer
        deviceRef = hs.GetDeviceRefByName(lightDeviceNames(n))
        If deviceRef > 0 Then
            Dim value As Integer
            value = hs.DeviceValue(deviceRef)
            If value > 0 Then
                If devicesOn > 0 Then
                    summary = summary + ", "
                End If
                summary = summary + lightDeviceNames(n)
                devicesOn = devicesOn + 1
            End If
            UpdateLastChangedVirtualDevice(lightDeviceNames(n))
        Else
             hs.WriteLog("LiveSmart", String.Format("Warning: '{0}' device not found", lightDeviceNames(n)))
        End If
    Next n

    If devicesOn = 0 Then
        summary = "All lamps are off"
        summaryValue = 0
    ElseIf devicesOn = deviceCount Then
        summary = "All lamps are on"
        summaryValue = 99
    Else
        summary = summary + " on, all others off"
        summaryValue = 99
    End If

    hs.SetDeviceValueByName(lightSummaryVirtualDeviceName, summaryValue)
    hs.SetDeviceStringByName(lightSummaryVirtualDeviceName, summary, True)
End Sub

'===================================================================
' Sets the status summary virtual device to a string reflecting the
' current status of the specified motion sensor devices.
Private Sub SetMotionSummaryStatus()
    Dim deviceLastChanged As String = ""
    Dim minLastChanged As Integer
    minLastChanged = Integer.MaxValue
    Dim devices As Integer = 0
    Dim deviceCount As Integer
    deviceCount = motionDeviceNames.Length

    Dim n As Integer
    For n = 0 To deviceCount - 1
        Dim deviceRef As Integer
        deviceRef = hs.GetDeviceRefByName(motionDeviceNames(n))
        If deviceRef > 0 Then
            Dim value As Integer
            value = hs.DeviceValue(deviceRef)
            If value <> 0 Then
                Dim lastChanged As Integer
                lastChanged = hs.DeviceTimeByName(motionDeviceNames(n))
                If lastChanged < minLastChanged Then
                    deviceLastChanged = motionDeviceNames(n)
                    minLastChanged = lastChanged
                End If
                devices = devices + 1
            End If
            UpdateLastChangedVirtualDevice(motionDeviceNames(n))
        Else
            hs.WriteLog("LiveSmart", String.Format("Warning: '{0}' device not found", motionDeviceNames(n)))
        End If
    Next n

    If devices > 0 Then
        Dim dateTimeLastChanged As DateTime
        dateTimeLastChanged = DateTime.Now.AddMinutes(-minLastChanged)
        Dim timeLastChanged As String
        timeLastChanged = dateTimeLastChanged.ToString("ddd MM/dd hh:mm tt")

        hs.SetDeviceStringByName(motionSummaryDeviceVirtualDeviceName, StripParenthesisFromString(deviceLastChanged), True)
        hs.SetDeviceStringByName(motionSummaryTimeVirtualDeviceName, timeLastChanged, True)
    End If
End Sub

'===================================================================
' Sets the status summary virtual device To a String reflecting the
' current status of the specified sensor devices.
Private Sub SetSensorSummaryStatus()
    Dim summary As String = ""
    Dim devicesOn As Integer = 0
    Dim minLastChanged As Integer
    minLastChanged = Integer.MaxValue
    Dim deviceCount As Integer
    deviceCount = sensorDeviceNames.Length

    Dim n As Integer
    For n = 0 To deviceCount - 1
        Dim deviceRef As Integer
        deviceRef = hs.GetDeviceRefByName(sensorDeviceNames(n))
        If deviceRef > 0 Then
            Dim value As Integer
            value = hs.DeviceValue(deviceRef)
            If value > 0 Then
                devicesOn = devicesOn + 1
            End If
            Dim lastChanged As Integer
            lastChanged = hs.DeviceTimeByName(sensorDeviceNames(n))
            If lastChanged < minLastChanged Then
                minLastChanged = lastChanged
            End If
            UpdateLastChangedVirtualDevice(sensorDeviceNames(n))
        Else
            hs.WriteLog("LiveSmart", String.Format("Warning: '{0}' device not found", sensorDeviceNames(n)))
        End If
    Next n

    If devicesOn = 0 Then
        summary = "All Closed"
    ElseIf devicesOn = deviceCount Then
        summary = "All Open"
    Else
        summary = devicesOn.ToString() + " Open"
    End If

    hs.SetDeviceStringByName(sensorSummaryVirtualDeviceName, summary, True)
End Sub

'===================================================================
' Sets the status summary virtual device to a string reflecting the
' current status of the specified appliance devices.
Private Sub SetApplianceSummaryStatus()
    Dim summary As String = ""
    Dim summaryValue As Integer = 0
    Dim devicesOn As Integer = 0
    Dim deviceCount As Integer
    deviceCount = applianceDeviceNames.Length

    Dim n As Integer
    For n = 0 To deviceCount - 1
        Dim deviceRef As Integer
        deviceRef = hs.GetDeviceRefByName(applianceDeviceNames(n))
        If deviceRef > 0 Then
            Dim value As Integer
            value = hs.DeviceValue(deviceRef)
            If value > 0 Then
                If devicesOn > 0 Then
                    summary = summary + ", "
                End If
                summary = summary + applianceDeviceNames(n)
                devicesOn = devicesOn + 1
            End If
            UpdateLastChangedVirtualDevice(applianceDeviceNames(n))
        Else
            hs.WriteLog("LiveSmart", String.Format("Warning: '{0}' device not found", applianceDeviceNames(n)))
        End If
    Next n

    If devicesOn = 0 Then
        summary = "All appliances are off"
        summaryValue = 0
    ElseIf devicesOn = deviceCount Then
        summary = "All appliances are on"
        summaryValue = 99
    Else
        summary = summary + " on, all others off"
        summaryValue = 99
    End If

    hs.SetDeviceValueByName(applianceSummaryVirtualDeviceName, summaryValue)
    hs.SetDeviceStringByName(applianceSummaryVirtualDeviceName, summary, True)
End Sub

'===================================================================
' Sets the status summary virtual device to a string reflecting the
' current status of the specified detectors.
Private Sub SetDetectorSummaryStatus()
    Dim summary As String = ""
    Dim summaryValue As Integer = 0
    Dim devicesOn As Integer = 0
    Dim deviceCount As Integer
    deviceCount = detectorDeviceNames.Length

    Dim n As Integer
    For n = 0 To deviceCount - 1
        Dim deviceRef As Integer
        deviceRef = hs.GetDeviceRefByName(detectorDeviceNames(n))
        If deviceRef > 0 Then
            Dim value As Integer
            value = hs.DeviceValue(deviceRef)
            ' Smoke NOT OK values: 1.255, 2.255 -- Water NOT OK value: 255
            If value = 255 Or value = 1.255 Or value = 2.255 Then
                If devicesOn > 0 Then
                    summary = summary + ", "
                End If
                summary = summary + StripParenthesisFromString(detectorDeviceNames(n))
                devicesOn = devicesOn + 1
            End If
            UpdateLastChangedVirtualDevice(detectorDeviceNames(n))
        Else
            hs.WriteLog("LiveSmart", String.Format("Warning: '{0}' device not found", detectorDeviceNames(n)))
        End If
    Next n

    If devicesOn = 0 Then
        summary = "All detectors are OK"
        summaryValue = 0
    Else
        summary = summary + " set to ALARM, all others OK"
        summaryValue = 99
    End If

    hs.SetDeviceValueByName(detectorSummaryVirtualDeviceName, summaryValue)
    hs.SetDeviceStringByName(detectorSummaryVirtualDeviceName, summary, True)
End Sub

'===================================================================
' Logging functions
'===================================================================
Private Function CheckLogMessageForDeviceName(ByVal logMessage As String, ByVal names() As String, ByRef index As Integer) As Boolean
    Dim n As Integer
    For n = 0 To names.Length - 1
        If logMessage.Contains(names(n)) Then
            index = n
            Return True
        End If
    Next n

    index = -1
    Return False
End Function

'===================================================================
Private Function StripEnclosedCharactersFromString(ByVal s As String, ByVal cOpen As String, ByVal cClose As String) As String
    Dim stripped As String = ""
    Dim inTag As Boolean = False

    Dim n As Integer
    For n = 0 To s.Length - 1
        Dim c As String
        c = s.Substring(n, 1)
        If inTag = True Then
            If c = cClose Then
                inTag = False
            End If
        Else
            If c = cOpen Then
                inTag = True
            Else
                stripped = stripped & c
            End If
        End If
    Next n

    Return stripped
End Function

'===================================================================
Private Function StripHtmlFromString(ByVal s As String) As String
    Return StripEnclosedCharactersFromString(s, "<", ">")
End Function

'===================================================================
Private Function StripParenthesisFromString(ByVal s As String) As String
    Return StripEnclosedCharactersFromString(s, "(", ")")
End Function

'===================================================================
Private Function FormatLightLogMessage(ByVal log As HomeSeerAPI.LogEntry, ByVal name As String) As String
    Dim message As String
    Dim onOff As String
    Dim logText As String
    logText = StripHtmlFromString(log.LogText)
    If logText.Contains("Set to 0") Then
        onOff = "OFF"
    Else
        onOff = "ON"
    End If
    message = String.Format("{0} was turned {1} at {2} on {3}", name.ToUpper(), onOff, log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
    Return message
End Function

'===================================================================
Private Function FormatMotionLogMessage(ByVal log As HomeSeerAPI.LogEntry, ByVal name As String) As String
    Dim message As String
    Dim logText As String
    logText = StripHtmlFromString(log.LogText)
    If logText.Contains("Set to On/Open/Motion") Then
        message = String.Format("Motion detected in {0} at {1} on {2}", StripParenthesisFromString(name.ToUpper()).Replace(" MULTISENSOR",""), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
    Else
        message = ""
    End If
    Return message
End Function

'===================================================================
Private Function FormatSensorLogMessage(ByVal log As HomeSeerAPI.LogEntry, ByVal name As String) As String
    Dim message As String
    Dim state As String
    Dim logText As String
    logText = StripHtmlFromString(log.LogText)
    If logText.Contains("Set to On/Open/Motion") Then
        state = "OPEN"
    Else
        state = "CLOSED"
    End If
    message = String.Format("{0} was {1} at {2} on {3}", StripParenthesisFromString(name.ToUpper()).Replace(" MULTISENSOR",""), state, log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
    Return message
End Function

'===================================================================
Private Function TranslateLockCode(ByVal s) As String
    Dim translation As String
    translation = ""

    Try
        Dim search As String
        Dim index As Integer
        Dim codeStartIndex As Integer
        Dim n As Integer
        Dim lockCodeTableParts() As String
        search = "set to "
        index = s.IndexOf(search, StringComparison.OrdinalIgnoreCase)
        codeStartIndex = index + search.Length
        If index >= 0 And s.Length > codeStartIndex Then
            For n = 0 To lockCodeTable.Length - 1
                lockCodeTableParts = lockCodeTable(n).Split(New Char() {","c})
                If lockCodeTableParts.Length >= 2 Then
                    If s.Substring(codeStartIndex).StartsWith(lockCodeTableParts(0)) Then
                        translation = lockCodeTableParts(1)
                        Exit For
                    End If
                End If
            Next n
            If translation = "" Then
                If s.Substring(codeStartIndex).StartsWith("locked", StringComparison.OrdinalIgnoreCase) Then
                    translation = "LOCKED"
                ElseIf s.Substring(codeStartIndex).StartsWith("unlocked", StringComparison.OrdinalIgnoreCase) Then
                    translation = "UNLOCKED"
		        End If
	        End If
        End If
    Catch e As Exception
        hs.WriteLog("LiveSmart", e.Message)
    End Try

    Return translation
End Function

'===================================================================
Private Function FormatLockLogMessage(ByVal log As HomeSeerAPI.LogEntry, ByVal name As String) As String
    Dim message As String
    Dim logText As String
    Dim translation As String
    message = ""
    logText = StripHtmlFromString(log.LogText)
    translation = TranslateLockCode(logText)
    If translation <> "" Then
        message = String.Format("{0} was {1} at {2} on {3}", name.ToUpper(), translation, log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
    End If
    Return message
End Function

'===================================================================
Private Function FormatApplianceLogMessage(ByVal log As HomeSeerAPI.LogEntry, ByVal name As String) As String
    Dim message As String
    Dim onOff As String
    Dim logText As String
    logText = StripHtmlFromString(log.LogText)
    If logText.Contains("Set to 0") Then
        onOff = "OFF"
    Else
        onOff = "ON"
    End If
    message = String.Format("{0} was turned {1} at {2} on {3}", name.ToUpper(), onOff, log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
    Return message
End Function

'===================================================================
Private Function FormatDetectorLogMessage(ByVal log As HomeSeerAPI.LogEntry, ByVal name As String) As String
    Dim message As String
    Dim onOff As String
    Dim logText As String
    logText = StripHtmlFromString(log.LogText)
    If logText.Contains("255") Then
        onOff = "ALARM"
    Else
        onOff = "OK"
    End If
    message = String.Format("{0} set to {1} at {2} on {3}", StripParenthesisFromString(name.ToUpper()), onOff, log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
    Return message
End Function

'===================================================================
Private Function FormatThermostatLogMessage(ByVal log As HomeSeerAPI.LogEntry, ByVal name As String) As String
    Dim message As String
    message = ""

    Try
        Dim logText As String
        logText = StripHtmlFromString(log.LogText)
        Dim logTextSplit() As String = logText.Split(New Char() {" "c})

        If logTextSplit.Length > 2 Then
            If logText.Contains(name + " Heating") And logText.Contains("Setpoint Set to Setpoint") Then
                message = String.Format("{0} HEAT set to {1} F at {2} on {3}", StripParenthesisFromString(name.ToUpper()), logTextSplit(logTextSplit.Length - 2), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))

            ElseIf logText.Contains(name + " Heating") And logText.Contains("Setpoint Set to") Then
                message = String.Format("{0} HEAT set to {1} F at {2} on {3}", StripParenthesisFromString(name.ToUpper()), logTextSplit(logTextSplit.Length - 1), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))

            ElseIf logText.Contains(name + " Cooling") And logText.Contains("Setpoint Set to Setpoint") Then
                message = String.Format("{0} COOL set to {1} F at {2} on {3}", StripParenthesisFromString(name.ToUpper()), logTextSplit(logTextSplit.Length - 2), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))

            ElseIf logText.Contains(name + " Cooling") And logText.Contains("Setpoint Set to ") Then
                message = String.Format("{0} COOL set to {1} F at {2} on {3}", StripParenthesisFromString(name.ToUpper()), logTextSplit(logTextSplit.Length - 1), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))

            ElseIf logText.Contains(name + " Mode Set to") Then
                Dim modes() As String = { "OFF", "HEAT", "COOL", "AUTO" }
                message = String.Format("{0} set to {1} at {2} on {3}", StripParenthesisFromString(name.ToUpper()), modes(Val(logTextSplit(logTextSplit.Length - 1))), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))

            ElseIf logText.Contains(name + " Operating State Set to") Then
                Dim states() As String = { "automatically turned OFF", "HEAT automatically turned ON", "COOL automatically turned ON"}
                message = String.Format("{0} {1} at {2} on {3}", StripParenthesisFromString(name.ToUpper()), states(Val(logTextSplit(logTextSplit.Length - 1))), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))

            ElseIf logText.Contains(name + " Fan Mode Set to") Then
                Dim fanModes() As String = { "OFF", "ON"}
                message = String.Format("{0} fan turned {1} at {2} on {3}", StripParenthesisFromString(name.ToUpper()), fanModes(Val(logTextSplit(logTextSplit.Length - 1))), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
            End If
        End If
    Catch e As Exception
        hs.WriteLog("LiveSmart",  "FormatThermostatLogMessage(): " + e.Message)
    End Try

    Return message
End Function

'===================================================================
Private Function FormatPinLogMessage(ByVal log As HomeSeerAPI.LogEntry) As String
    Dim message As String
    message = String.Format("{0} at {1} on {2}", log.logText, log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
    Return message
End Function

'===================================================================
' Roll our own GetLog retrieve/parse functions because HomeSeer's
' GetLog_* functions are crashing on certain devices...
Private Function GetCurrentLogs() As String()
    Dim logs() As String

    Try
        Dim logsString As String
        logsString = hs.LogGet()
        logs = logsString.Split(new String() {Environment.NewLine}, StringSplitOptions.None)
    Catch e As Exception
        hs.WriteLog("LiveSmart", e.Message)
    End Try

    Return logs
End Function

' We'll use the LogEntry structure HomeSeer has defined already...
'
'Public Structure LogEntry
'    Public LogTime As Date		    ' The date and time the log entry was recorded.
'    Public LogType As String		' The 'type' string as logged.
'    Public LogText As String		' The main message text of the log entry.
' ... more fields we don't use ...
'End Structure

Private Function ParseLog(ByVal logString As String, ByRef log As HomeSeerAPI.LogEntry) As Boolean
    Try     
        Dim logFields() As String
        logFields = logString.Split(new String() {"~!~"}, StringSplitOptions.None)

        If logFields.Length < 3 Then
            Return False
        End If

        log.LogTime = Date.Parse(logFields(0))
        log.LogType = logFields(1)
        log.LogText = logFields(2)
    Catch e As Exception
        hs.WriteLog("LiveSmart", e.Message)
        Return False
    End Try

    Return True
 End Function

'===================================================================
Private Sub UpdateLogHistory()
    Dim logMessages As String
    Dim lightMessages As String
    Dim motionMessages As String
    Dim sensorMessages As String
    Dim lockMessages As String
    Dim applianceMessages As String
    Dim detectorMessages As String
    Dim thermostatMessages As String
    Dim pinSuccessMessages As String
    Dim pinFailureMessages As String

    Dim logs() As String
    Dim logCount As Integer = 0

    logs = GetCurrentLogs()
    If logs IsNot Nothing Then
        logCount = logs.Length
    End If

    Dim n As Integer
    For n = 0 To logCount - 1
        Dim nameIndex As Integer
        Dim message As String

        Dim log As HomeSeerAPI.LogEntry

        If ParseLog(logs(n), log) Then
            If log.LogType = "Z-Wave" Then
                If CheckLogMessageForDeviceName(log.LogText, lightDeviceNames, nameIndex) Then
                    message = FormatLightLogMessage(log, lightDeviceNames(nameIndex))
                    logMessages = message & CRLF & logMessages & CRLF
                    lightMessages = message & CRLF & lightMessages & CRLF
                ElseIf CheckLogMessageForDeviceName(log.LogText, motionDeviceNames, nameIndex) Then
                    message = FormatMotionLogMessage(log, motionDeviceNames(nameIndex))
                    If message <> "" Then
                        logMessages = message & CRLF & logMessages & CRLF
                        motionMessages = message & CRLF & motionMessages & CRLF
                    End If
                ElseIf CheckLogMessageForDeviceName(log.LogText, sensorDeviceNames, nameIndex) Then
                    message = FormatSensorLogMessage(log, sensorDeviceNames(nameIndex))
                    logMessages = message & CRLF & logMessages & CRLF
                    sensorMessages = message & CRLF & sensorMessages & CRLF
                ElseIf CheckLogMessageForDeviceName(log.LogText, lockDeviceNames, nameIndex) Then
                    message = FormatLockLogMessage(log, lockDeviceNames(nameIndex))
                    If message <> "" Then
                        logMessages = message & CRLF & logMessages & CRLF
                        lockMessages = message & CRLF & lockMessages & CRLF
                    End If
                ElseIf CheckLogMessageForDeviceName(log.LogText, applianceDeviceNames, nameIndex) Then
                    message = FormatApplianceLogMessage(log, applianceDeviceNames(nameIndex))
                    logMessages = message & CRLF & logMessages & CRLF
                    applianceMessages = message & CRLF & applianceMessages & CRLF
                ElseIf CheckLogMessageForDeviceName(log.LogText, detectorDeviceNames, nameIndex) Then
                    message = FormatDetectorLogMessage(log, detectorDeviceNames(nameIndex))
                    logMessages = message & CRLF & logMessages & CRLF
                    detectorMessages = message & CRLF & detectorMessages & CRLF
                ElseIf CheckLogMessageForDeviceName(log.LogText, thermostatDeviceNames, nameIndex) Then
                    message = FormatThermostatLogMessage(log, thermostatDeviceNames(nameIndex))
                    If message <> "" Then
                        logMessages = message & CRLF & logMessages & CRLF
                        thermostatMessages = message & CRLF & thermostatMessages & CRLF
                    End If
                End If
            ElseIf log.LogType = PinSuccessLogMessageType Then
                message = FormatPinLogMessage(log)
                logMessages = message & CRLF & logMessages & CRLF
                pinSuccessMessages = message & CRLF & pinSuccessMessages & CRLF
            ElseIf log.LogType = PinFailureLogMessageType Then
                message = FormatPinLogMessage(log)
                logMessages = message & CRLF & logMessages & CRLF
                pinFailureMessages = message & CRLF & pinFailureMessages & CRLF
            End If
        End If

    Next n

    hs.SetDeviceStringByName(logVirtualDeviceName, logMessages, True)
    hs.SetDeviceStringByName(lightLogVirtualDeviceName, lightMessages, True)
    hs.SetDeviceStringByName(motionLogVirtualDeviceName, motionMessages, True)
    hs.SetDeviceStringByName(sensorLogVirtualDeviceName, sensorMessages, True)
    hs.SetDeviceStringByName(lockLogVirtualDeviceName, lockMessages, True)
    hs.SetDeviceStringByName(applianceLogVirtualDeviceName, applianceMessages, True)
    hs.SetDeviceStringByName(detectorLogVirtualDeviceName, detectorMessages, True)
    hs.SetDeviceStringByName(thermostatLogVirtualDeviceName, thermostatMessages, True)
    hs.SetDeviceStringByName(pinSuccessLogVirtualDeviceName, pinSuccessMessages, True)
    hs.SetDeviceStringByName(pinFailureLogVirtualDeviceName, pinFailureMessages, True)
End Sub

'===================================================================
Private Sub UpdateAtHomeLogHistory()
    Dim atHomeMessages As String

    Dim logs() As String
    Dim logCount As Integer = 0

    logs = GetCurrentLogs()
    If logs IsNot Nothing Then
        logCount = logs.Length
    End If

    Dim n As Integer
    For n = 0 To logCount - 1
        Dim log As HomeSeerAPI.LogEntry

        If ParseLog(logs(n), log) Then
            If log.LogType = "Device Control" And log.LogText.Contains("Arrive Home") Then
                Dim message As String
                message = String.Format("{0} at {1} on {2}", StripParenthesisFromString(log.LogText.Replace("Device: Virtual Virtual ","")), log.LogTime.ToString("hh:mm tt"), log.LogTime.ToString("MMM-dd"))
                message = message.Replace("Arrive Home to Off ", "has left home ")
                message = message.Replace("Arrive Home to On ", "has arrived home ")
                atHomeMessages = message & CRLF & atHomeMessages
            End If
        End If
    Next n

    hs.SetDeviceStringByName(atHomeLogVirtualDeviceName, atHomeMessages, True)
End Sub