'// Create a Skype4COM object:
Set oSkype = WScript.CreateObject("Skype4COM.Skype","Skype_")

'// Start the Skype client:
If Not oSkype.Client.IsRunning Then oSkype.Client.Start() End If

'// Create a custom group:
Set oMyGroup = oSkype.CreateGroup("MyGroup")

'// Add user "echo123" to the group:
oMyGroup.AddUser "echo123"

'// List the users in MyGroup and return the Skypename of each user:
WScript.Echo "Group ("& oMyGroup.Id &") labelled ["& oMyGroup.DisplayName &"] has " & oMyGroup.Users.Count & " users."
For Each oUser In oMyGroup.Users
  WScript.Echo oUser.Handle & " ("& oUser.FullName &")"
Next

'// Remove a user from a group:
oMyGroup.RemoveUser "echo123"
  
'// Delete a custom group:
oSkype.DeleteGroup oMyGroup.Id

'// List all groups:
WScript.Echo "There are total " & oSkype.Groups.Count & " groups (" &_
oSkype.CustomGroups.Count & " custom and " & oSkype.HardwiredGroups.Count & " hardwired)" & vbCrLf 

'// List all custom groups:
WScript.Echo "Custom groups are:"& vbCrLf 
For Each oGroup In oSkype.CustomGroups
  WScript.Echo "Group ("& oGroup.Id &") labelled ["& oGroup.DisplayName &"] has " & oGroup.Users.Count & " users."
  For Each oUser In oGroup.Users
    WScript.Echo oUser.Handle & " ("& oUser.FullName &")"
  Next
  WScript.Echo ""
Next

'// List all hardwired groups and return the Skypename and full name of each member of each group:
WScript.Echo "Hardwired groups are:"& vbCrLf 
For Each oGroup In oSkype.HardwiredGroups
  WScript.Echo "Group ("& oGroup.Id &") type of ["& oSkype.Convert.GroupTypeToText(oGroup.Type) &"] has " & oGroup.Users.Count & " users."
  For Each oUser In oGroup.Users
    WScript.Echo oUser.Handle & " ("& oUser.FullName &")"
  Next
  WScript.Echo ""  
Next

'// Keep the script running for 60 seconds:
WScript.Sleep(60000)

'// The GroupVisible event handler returns information about whether a group is visible or hidden:
Public Sub Skype_GroupVisible(ByRef aGroup, ByVal aVisible)     
  WScript.StdOut.Write "Group ("& aGroup.Id &") type of ["& oSkype.Convert.GroupTypeToText(aGroup.Type) &"] labelled ["& aGroup.DisplayName &"]"
  If aVisible Then
    WScript.Echo " is visible."
  Else
    WScript.Echo " is hidden."
  End If
End Sub

'// The GroupExpanded event handler returns information about whether a group is expanded or collapsed:
Public Sub Skype_GroupExpanded(ByRef aGroup, ByVal aExpanded)     
  WScript.StdOut.Write "Group ("& aGroup.Id &") type of ["& oSkype.Convert.GroupTypeToText(aGroup.Type) &"] labelled ["& aGroup.DisplayName &"]"
  If aExpanded Then
    WScript.Echo " is expanded."
  Else
    WScript.Echo " is collapsed."
  End If
End Sub

'// The GroupUsers event handler gets the Skypenames of members of a group:
Public Sub Skype_GroupUsers(ByRef aGroup, ByRef aUsers)     
  WScript.StdOut.Write "Group ("& aGroup.Id &") type of ["& oSkype.Convert.GroupTypeToText(aGroup.Type) &"] labelled ["& aGroup.DisplayName &"] users"
  Dim oUser
  For Each oUser In aUsers
    WScript.Stdout.Write " " & oUser.Handle
  Next
  WScript.Echo ""
End Sub

'// Delete a group:
Public Sub Skype_GroupDeleted(ByRef aGroupId)     
  WScript.Echo "Group " & aGroupId & " deleted."
End Sub

'// Bring a contact into focus:
Public Sub Skype_ContactsFocused(ByVal aHandle)     
  If Len(aHandle) Then
    WScript.Echo "Contact " & aHandle & " focused."
  Else
    WScript.Echo "No contact focused."
  End If  
End Sub

'// The AttachmentStatus event handler monitors attachment status and automatically attempts to reattach to the API following loss of connection:
Public Sub Skype_AttachmentStatus(ByVal aStatus)
  WScript.Echo ">Attachment status " & oSkype.Convert.AttachmentStatusToText(aStatus)
  If aStatus = oSkype.Convert.TextToAttachmentStatus("AVAILABLE") Then oSkype.Attach() End If
End Sub
