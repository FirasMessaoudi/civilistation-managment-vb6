Attribute VB_Name = "Module2"
'FONCTION Pour tester si L'admin existe dans le fichier ou non
Public Function exist(user As logger) As Integer
Dim user1 As logger
'user.login = txtlogin.Text
'user.pass = txtpass.Text
Dim j As Integer
Dim trouve As Boolean
Open "pass.txt" For Random As #3 Len = Len(user1)
j = 1
trouve = False
While Not EOF(3) And Not trouve
Get #3, j, user1
If user.login = user1.login And user1.pass = user.pass Then
exist = j
trouve = True
End If
j = j + 1
Wend
If Not trouve Then
exist = -1
End If
Close #3
End Function

'La meme fonction mais pour les utilisateurs

Public Function exist2(user As logger) As Integer

Dim user2 As logger
Dim j As Integer
Dim k As Integer
Dim trouve As Boolean
k = FreeFile
Open "users.txt" For Random As #k Len = Len(user2)
j = 1
trouve = False
While Not EOF(k) And Not trouve
Get #k, j, user2
If user.login = user2.login And user2.pass = user.pass Then
exist2 = j
trouve = True
End If
j = j + 1
Wend
If Not trouve Then
exist2 = -1
End If
Close #k
End Function


Public Function exist_nom(user As logger) As Integer

Dim user2 As logger
Dim j As Integer
Dim k As Integer
Dim trouve As Boolean
k = FreeFile
Open "users.txt" For Random As #k Len = Len(user2)
j = 1
trouve = False
While Not EOF(k) And Not trouve
Get #k, j, user2
If user.login = user2.login Then
exist_nom = j
trouve = True
End If
j = j + 1
Wend
If Not trouve Then
exist_nom = -1
End If
Close #k
End Function
