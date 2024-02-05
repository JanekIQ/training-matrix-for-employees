# Training-matrix-for-employees
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

In _training deadline_ in column J, spreadsheet shows if employee has completed the training, or if not, how much time does their have to do so (in days). It also shows employee status - _out of office_ or _after long leave_. To visualize employee status, 
conditional formatting and different cell colours were used. Cell may stays black if employee doesn't need to take some particular training.
