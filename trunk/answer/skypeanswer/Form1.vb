Public Class Form1

    Dim oSkype As SKYPE4COMLib.Skype

    Private Sub Form_Load()
        '// Create a Skype4COM object:
        'Set oSkype = WScript.CreateObject("Skype4COM.Skype", "Skype_")
        oSkype = New SKYPE4COMLib.Skype 'CreateObject("Skype4COM.Skype", "Skype_")

        '// Start the Skype client:
        'If Not oSkype.Client.IsRunning Then oSkype.Client.Start() End If
        If Not oSkype.Client.IsRunning Then
            Call oSkype.Client.Start()
        End If

        '// Connect to the Skype API:
        oSkype.Attach()

        '// Run an infinite loop to ensure the connection remains active:
        Do While True
            'WScript.Sleep (60000)
            'DoEvents()
            System.Threading.Thread.Sleep(60000)

        Loop

    End Sub

    '// The AttachmentStatus event handler monitors attachment status and automatically attempts to reattach to the API following loss of connection:
    Public Sub Skype_AttachmentStatus(ByVal aStatus)
        'WScript.Echo ">Attachment status " & oSkype.Convert.AttachmentStatusToText(aStatus)
        'If aStatus = oSkype.Convert.TextToAttachmentStatus("AVAILABLE") Then oSkype.Attach() End If
        If aStatus = oSkype.Convert.TextToAttachmentStatus("AVAILABLE") Then
            Call oSkype.Attach()
        End If
    End Sub

    '// The CallStatus event handler monitors call status and if the status is "ringing" and it is an incoming call, it attempts to answer the call:
    Public Sub Skype_CallStatus(ByRef aCall, ByVal aStatus)
        'WScript.Echo ">Call " & aCall.Id & " status " & aStatus & " " & oSkype.Convert.CallStatusToText(aStatus)
        If oSkype.Convert.TextToCallStatus("RINGING") = aStatus And _
          (oSkype.Convert.TextToCallType("INCOMING_P2P") = aCall.Type Or _
           oSkype.Convert.TextToCallType("INCOMING_PSTN") = aCall.Type) Then
            'WScript.Echo("Answering call from " & aCall.PartnerHandle)
            ListView1.Items.Add("Call target identity: " & aCall.TargetIdentity)
            If aCall.TargetIdentity <> "" Then
                'WScript.Echo "Call target identity: " & aCall.TargetIdentity
                ListView1.Items.Add("Call target identity: " & aCall.TargetIdentity)
            End If
            aCall.Answer()
        End If
    End Sub


End Class
