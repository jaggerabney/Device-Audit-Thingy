# Device Audit Thingy

This is a script I've been plinking away at during work hours to help the district's inventory guy. It pulls info from various spreadsheets to create a master inventory list that'll be inputted into the district's new ticketing software. I chose Java because it's the language I'm most comfortable with - sue me, C# nerds - and because I hate working with VBA.

## Usage
DAT works on a per-school basis. In other words, it is made to collate all device information for a given school and output it to one workbook. To use DAT, drop it in a directory with the following three files:

- inventory.xlsx : the old master inventory list
- audit.xlsx : the device audit workbook for the given school
- target.xlsx : the sheet that everything will be written to

To not disclose *everything* that the district has, I've chosen to keep these files out of the repo. If you're the one person who will actually use this script, however, you know what these files are. 

The actual school in question (which gets inputted into target.xlsx's Location column) is passed as an argument when the program is ran. This means that the program has to be run from the command line. I think you can also add the arguments to the .jar itself by editing its Properties in File Explorer, but that would be tedious and annoying, so just use the command line. I've been debugging the program with the following command:

    java -jar dat-v[version number].jar "[location]"
    
If we were using v1.0.0 - the latest release, as of writing this - and the location were Bow Lake, then the command would be as follows:

    java -jar dat-v1.0.0.jar "Bow Lake"
    
I personally like to put the command in a .bat file with some extra stuff to make it look pretty, but you don't have to. 
