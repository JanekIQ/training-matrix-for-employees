# Training matrix for employees
Here I present training matrix for employees made in excel with some automatic option and short VBA script. Project can be the starting point to make your own matrix or just an inspiration. (Excel spreadsheet is in polish)
### **Training matrix for employees**
**Here I present training matrix for employees made in excel with some automatic option and short VBA script.**
The purpose of the spreadsheet is to organise and maintain trainings in procedures for all employees taking into account their position, employment day and all vacations in real time. 
Spreadsheet helps to meet all training deadlines.

There are five sheets:
`Training roadmap`, `Employees`, `Training matrix`, `Metadata`, `Information`


![image](https://github.com/JanekIQ/training-matrix-for-employees/assets/129321529/d6af4d23-1e96-4dae-9cb9-64633178516a)

### `Training roadmap` 

![image](https://github.com/JanekIQ/training-matrix-for-employees/assets/129321529/89128874-9d7d-45f2-92f4-71ef25b6ca2c)


In the top left corner there is a today's date to which all formulas refer.
Going from the left - there are couple info titles and panels in which (_training package, document code, document title, version, training form, implementation date, days to implement_) than every employee has their own section (_position, training date, training deadline_)

_Version_ column is crucial. After updating version of a document there is a need to retake the training so that after changing _version_  in column D, all training dates and deadlines of this particular document are deleted. This includes all employees whom position require this training.
This function is made using `VBA code`.

    Dim originalValues As New Collection
    Dim originalFormulas As New Collection
    Dim undoChanges As Boolean 
    Private Sub Worksheet_Change(ByVal Target As Range)
    Dim AffectedRange As Range
    Dim Cell As Range
    Dim ChangeRange As Range
    Dim ColDRange As Range
    Dim i As Integer
    
    ' Determine which cells have changed in a columne D (D8:D80)
    Set ChangeRange = Intersect(Target, Me.Range("D8:D80"))
    
    ' If the changes affect a column D
    If Not ChangeRange Is Nothing Then
        Application.EnableEvents = False ' Disable event handling to avoid incorrectly calling this procedure again
        
        ' Clearing the set of original values
        Set originalValues = New Collection
        Set originalFormulas = New Collection
        undoChanges = False 
        
        ' Storing the original cell values ​​in column D
        For Each Cell In ChangeRange
            originalValues.Add Cell.Value
            originalFormulas.Add Cell.Formula
        Next Cell
        
        ' Displaying a message to the user with the option "Continue?
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox("NOTE: Changing the version will delete the training date for a given document. Information specifying the form of training, implementation date, impact on GxP and time (days) for implementation will also be deleted. Do you want to CONTINUE?", vbExclamation + vbYesNo, "Announcement")
        
        ' If the user selects "No" (i.e. does not want to continue)
        If userResponse = vbNo Then
            ' Iterate through the cells that have been changed
            For Each Cell In ChangeRange
                ' Finding the corresponding cell in column D (same row)
                Application.Undo 
                originalValues.Remove 1
                originalFormulas.Remove 1
            Next Cell
            undoChanges = False 
        Else ' If the user selects "Yes" (i.e. wants to continue), continue deleting and delete the contents of columns E, F, G, H
            ' Specifying the range of column D (D8:D80)
            Set ColDRange = Me.Range("D8:D80")
            
            ' Iterate through the cells that have been changed
            For Each Cell In ChangeRange
                ' Finding the corresponding cell in column I (in the same row)
                Me.Cells(Cell.Row, "I").ClearContents
                
                ' Deleting content in columns that are 50 multiples of 7 away from column I
                For i = 7 To 50 * 7 Step 7
                    Me.Cells(Cell.Row, "I").Offset(0, i).ClearContents
                Next i
                
                ' Delete the contents in columns E, F, G, H
                For i = -1 To -4 Step -1
                    Me.Cells(Cell.Row, "I").Offset(0, i).ClearContents
                Next i
            Next Cell
        End If
        
        Application.EnableEvents = True ' Enable event handling
    End If
    End Sub




In _training deadline_ in column J, spreadsheet shows if employee has completed the training, or if not, how much time does their have to do so (in days). It also shows employee status - _out of office_ or _after long leave_. To visualize employee status, 
conditional formatting and different cell colours were used. Cell may stays black if employee doesn't need to take some particular training.
### `Employees`
` 
![image](https://github.com/JanekIQ/training-matrix-for-employees/assets/129321529/679dbf9e-a9ce-4acc-aace-d0afb49fc336)

Here manager of the spreadsheet can input some information about all employees.
_Employment date_, check if employee is on vacation, when does employee come back to work after longer leave. 
All of that affect _training deadline_ in `Training roadmap`.

### `Training Matrix`

![image](https://github.com/JanekIQ/training-matrix-for-employees/assets/129321529/76562cc0-60da-4cb6-9db6-1cba79e95fef)

This sheet determines which training packages have to be completed by employees at a specific position.
This shows or covers space to add _training date_ and _training deadline_ in `Training roadmap`. 
Shows - when an employee have to complete specific training in a given position
or
Covers - when an employee does not need to complete specific training in a given position
