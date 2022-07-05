# Device Audit Thingy

This is a script I've been plinking away at during work hours to help the district's inventory guy. It pulls info from various spreadsheets to create a master inventory list that'll be inputted into the district's new ticketing software. I chose Java because it's the language I'm most comfortable with - sue me, C# nerds - and because I hate working with VBA.

## Usage

DAT works on a per-school basis. In other words, it is made to collate all device information for a given school and output it to one workbook. To use DAT, drop it in a directory with the following two files:

- inventory.xlsx : the old master inventory list
- audit.xlsx : the device audit workbook for the given school

To not disclose _everything_ that the district has, I've chosen to keep these files out of the repo. If you're the one person who will actually use this script, however, you know what these files are.

Because DAT is ran from a .jar file, it's most convenient to run it from the command line. I plan on making it capable of launching on its own in the future, but for now, it's meant to be run from the command line. This is the command I've been using to run it:

    java -jar dat-v[version number].jar

For example, the release I plan on putting out right after I finish editing this readme file would be ran as follows:

    java -jar dat-v1.0.2.jar

Please note that this release must have the config.properties file in the same directory as the .jar and the two .xlsx files! Otherwise, it won't run. There are several things in this file that can be edited for flexibility: workbook names, column names, column indices, console messages, etc. Just make sure that it's accurate, otherwise things will break.

v1.0.2 includes a .bat file that makes running DAT a little less awful.
