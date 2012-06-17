
'This script provides a response to a message received in a chat. 

'http://stackoverflow.com/questions/2806584/how-do-i-run-a-vbscript-in-32-bit-mode-on-a-64-bit-machine
'How do I run a VBScript in 32-bit mode on a 64-bit machine?
'* Click Start, click Run, type %windir%\SysWoW64\cmd.exe, and then click OK.
'cscript vbscriptfile.vbs

'// Create a Skype4COM object:
Set oSkype = WScript.CreateObject("Skype4COM.Skype", "Skype_")

'// Start the Skype client:
If Not oSkype.Client.IsRunning Then oSkype.Client.Start() End If

'// Declare the following Skype constants:
cAttachmentStatus_Available = oSkype.Convert.TextToAttachmentStatus("AVAILABLE")
cMessageStatus_Sending = oSkype.Convert.TextToChatMessageStatus("SENDING")
cMessageStatus_Received = oSkype.Convert.TextToChatMessageStatus("RECEIVED")
cMessageType_Said = oSkype.Convert.TextToChatMessageType("SAID")
cMessageType_Left = oSkype.Convert.TextToChatMessageType("LEFT")

'// The SendMessage command will fail if the user is offline. To avoid failure, check user status and change to online if necessary:
If cUserStatus_Offline = oSkype.CurrentUserStatus Then oSkype.ChangeUserStatus(cUserStatus_Online) End If  

'// Sleep 
Do While True 
  WScript.Sleep(60000)
Loop

'// The AttachmentStatus event handler monitors attachment status and attempts to connect to the Skype API:
Public Sub Skype_AttachmentStatus(ByVal aStatus)
  WScript.Echo  ">Attachment status " & oSkype.Convert.AttachmentStatusToText(aStatus)
  If aStatus = cAttachmentStatus_Available Then oSkype.Attach() End If
End Sub

'// The MessageStatus event handler monitors message status, decodes received messages and, for those of type "Said", sends an autoresponse quoting the original message:
Public Sub Skype_MessageStatus(ByRef aMsg, ByVal aStatus)
  WScript.Echo ">Message " & aMsg.Id & " status " & oSkype.Convert.ChatMessageStatusToText(aStatus)
  If aStatus = cMessageStatus_Received Then 
    DecodeMsg aMsg       
    If aMsg.Type = cMessageType_Said Then 
     'oSkype.SendMessage aMsg.FromHandle, "<Reposta automatica GAutoAnswer: > desculpe estou ausente deixe seu recado."
      'aMsg.Chat.SendMessage "You said [" & aMsg.Body & "]"
	  aMsg.Chat.SendMessage "<Reposta automatica GAutoAnswer: > desculpe estou ausente deixe seu recado."
    End If
  End If    
End Sub

'// The DecodeMsg event handler decodes messages in a chat and converts leave reasons to text for messages of type "Left":
Public Sub DecodeMsg(ByRef oMsg)       
  sText = oMsg.FromHandle & " " & oSkype.Convert.ChatMessageTypeToText(oMsg.Type) & ":"
  If len(oMsg.Body) Then 
    sText = sText & " " & oMsg.Body
  End If
  Dim oUser
  For Each oUser In oMsg.Users
    sText = sText & " " & oUser.Handle
  Next
  If oMsg.Type = cMessageType_Left Then 
    sText = sText & " " & oSkype.Convert.ChatLeaveReasonToText(oMsg.LeaveReason)
  End If
  WScript.Echo ">" & sText  
End Sub
