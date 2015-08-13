# vbscript
VbScripts
Just a sample respository for the Vb script I am Testing

Please change the cell Row and Column to the exact cell where your TO address and CC is located.

            .to = Cells("2", "A").Value
            .CC = Cells("3", "A").Value
2 and 3 are Row and A is the Column.

Also you can change the Subject and Body Content by changing the below part

            .Subject = "Contest Form approval request"
            .Body = "Hi" & vbNewLine & _
            "Please help to approve the contest form."

