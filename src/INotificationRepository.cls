Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function EnqueueEmail(ByVal destinatarios As String, _
                           ByVal asunto As String, _
                           ByVal cuerpoHTML As String, _
                           Optional ByVal conCopia As String = "", _
                           Optional ByVal conCopiaOculta As String = "", _
                           Optional ByVal adjuntoPath As String = "") As Boolean
End Function