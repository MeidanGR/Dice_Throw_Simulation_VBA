' Developed by Meidan Greenberg

Sub DiceSimulation()
    ' This subprocedure simulates throwing fair-sided dice
    ' until all of the dice numbers are fetched regardless of the order.
    
    ' VARIABLE DEFINITION
    Dim n_simulations As Integer
    Dim n_throws As Byte
    Dim throws_count As Long
    
    Dim temp As Integer
    Dim firsts_array(5) As Integer
    Dim array_check As Integer
    
    ' PARAMETER SETUP
    n_simulations = Range("B2").Value
    n_throws = Range("B3").Value ' Maximum throws for each simulation.
    throws_count = 0
    
    For simulation = 1 To n_simulations
        Erase firsts_array
        array_check = 0
        
            For throw = 1 To n_throws
                temp = Int((6 - 1 + 1) * Rnd + 1)  ' create a random number between 1 and 6.
                
                If firsts_array(temp - 1) <> temp Then ' If the random number hasn't yet been assigned
                    firsts_array(temp - 1) = temp      ' Then It is the first time: keep it in an array.
                   
                    array_check = array_check + 1 ' Count the total numbers that has been assigned to the array.
                End If
                
                If array_check < 6 Then
                    GoTo NextThrow
                Else        ' When the array is full = all of the numbers have been assigned: finish the throws.
                    Exit For
                End If
NextThrow:
            Next throw
                
            throws_count = throws_count + throw 'Sum all of the relevant throws.
         
        Next simulation
    
Range("D3").Value = throws_count / n_simulations ' Calculate the average.
End Sub


