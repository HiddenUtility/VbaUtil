Attribute VB_Name = "HyspeedMode"
Private Function TrunOnHyspeed()

    With Application
        .ScreenUpdating = False
        .Calculation = xlManual
        .EnableEvents = False
    End With

End Function


Private Function TrunOffHyspeed()

    With Application
        .ScreenUpdating = True
        .Calculation = xlAutomatic
        .EnableEvents = True
    End With

End Function

