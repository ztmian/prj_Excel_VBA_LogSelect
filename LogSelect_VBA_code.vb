Sub Test()
Dim LogStr, cmd_key, cmd_expe, Key, Expe, vi, Logi


Range("C2:E500").ClearContents

For Logi = 2 To 500 Step 1
    LogStr = Range("F" & Logi).Value
    Range("E" & Logi).Value = ""
    
    For vi = 2 To 31 Step 1
        
        Key = Range("K" & vi).Value
        Expe = Range("J" & vi).Value
        cmd_key = Range("I" & vi).Value
        cmd_expe = Range("H" & vi).Value
        If LogStr Like cmd_key Then
            
            Range("C" & Logi).Value = cmd_expe
            'MsgBox "cmd_key" & Logi & ":" + cmd_key + "_vi" & vi
            
            If LogStr Like Key Then
                
                Range("E" & Logi).Value = Expe
                Range("D" & Logi).Value = vi
                 'MsgBox "no match" & Logi & ":" + cmd_key + "_vi" & vi
                Exit For
                
            End If
            
        Else
            'MsgBox "no match" & Logi & ":" + cmd_key + "_vi" & vi
            
        End If
     
    Next vi
        
    If vi >= 31 Then
        'MsgBox "Not Found[Key" & vi & ":]" + Key
    Else
        'MsgBox vi & "found"
    End If
        
Next Logi
    
End Sub
