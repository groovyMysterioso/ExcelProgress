Attribute VB_Name = "TestModule"
Sub ProgressTest()
    
    Dim x As Integer
    Dim MyTimer As Double
    Dim Max
    Dim MagicNumber
    Max = 200
         
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo HandleCancel:
    ProgressBar xlInitMeter, , Max
         
    For x = 0 To Max
         
        MyTimer = Timer
        Do: Loop While Timer - MyTimer < 0.03
        DoEvents
         
        ProgressBar xlUpdateMeter, "Working, please hold... ", x
        
    Next x

HandleCancel:
    ProgressBar xlRemoveMeter
     
End Sub


