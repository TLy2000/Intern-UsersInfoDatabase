#### Author
----------
## Thinh Ly
#### Date: 11/8/2020
#### Objective
	Take the user's input information and write it to an excel file, then read from the excel file.
#### Product
	Console App UsersInfoDatabase
### Console App
#### Functions
	UsersInfoDatabase project, Person (public static string GetFirstName(), public string GetLastName(), public static string GetPrompt()), Excel(public Excel(string path, int Sheet), public void ReadRange(int row, int col), public int CountColumns(int Sheet), public int CountRows(int Sheet), public void WriteRange(int count, List<string> list), public void Save(), public void SaveAs(string path), public void Close()).
-Person(public static string GetFirstName()): this functions to get the user's input, namely the first name
-Person(public string GetLastName()): this functions to get the user's input, namely the last name
-Person(public static string GetPrompt()): this functions as a way to run a do while loop, continously looping until the user inputs an "n" or an "N"
-Excel(public Excel(string path, int Sheet)): this creates a new Excel object
Excel(public void ReadRange(int row, int col)): this method takes two parameters, the number of rows and the number of columns that contain data, and reads all the cells contained within the parameters
-Excel(public int CountColumns(int Sheet)): this method counts the columns that contains data and returns an int value
-Excel(public int CountRows(int Sheet)): this method counts the rows that contains data and returns an int value
-Excel(public void WriteRange(int count, List<string> list)): this method writes the user's inputs into the excel file in two columns marked first name and last name
-Excel(public void Save()): this saves the excel file
-Excel(public void SaveAs(string path)): this assigns a designated file name to the excel file and saves it
-Excel(public void Close()): this closes the excel file
### References
-https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects
-https://stackoverflow.com/questions/43353073/c-sharp-excel-correct-way-to-get-rows-and-columns-count
-https://stackoverflow.com/questions/40369074/c-sharp-reading-excel-cell-values-using-microsoft-office-interop-excel
### My Experiences
	This would be my second project in C# and has elements based on a previous project, so in the beginning things were relatively easy. However, I had encountered several issues when trying to write to the excel file, namely writing multiple values. I had first tried to combine the first name list and last name list into a tuple and pass it to the WriteRange method, however I had some serious headaches in doing this. In the end, I had decided to pass two separate lists and had them written in separately. 
----------------
### New American Business Association
