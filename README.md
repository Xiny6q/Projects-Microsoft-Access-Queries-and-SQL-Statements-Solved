Download link :https://programming.engineering/product/projects-microsoft-access-queries-and-sql-statements-solved/

# Projects-Microsoft-Access-Queries-and-SQL-Statements-Solved
Projects: Microsoft Access Queries and SQL Statements Solved
Project 1: Written SQL Statements

Your company has a number of employees that are MANAGERS. Each MANAGER may ‘manage’ one or more PRODUCT. The database stores the date the MANAGER was assigned to each PRODUCT, the number of employees each manages per PRODUCT and the location of where each PRODUCT is manufactured. The database also stores the price of each PRODUCT.

Using the database shown above and create a Microsoft Word document containing the SQL statements (as shown in the lecture materials) to do the following:

Find the Manager ID, first name, last name, phone number and province for managers who live in Peterborough.

Modify the employee information for Manager ID equal to 3137 to your first name, last name and full address. (you may use a fake address, but must be your actual name).

Find only the Manager’s first and last name and product managed ID for all assigned dates that took place in March 2018.

Find the Product Name, Product Price, Number of Employees Managed, Assigned Date, and Manager’s Last Name for the products that have less than 50 employees being managed.

Find the number of products that you manage and save the result in a field labeled: Products Manages. (You MUST use your name in the query, NOT your Manager ID).

Find the Managers first and last name, manager’s phone number, Assigned Date, Location Name and Product Name for all products that cost less than $100.00.

Save the SQL statements in the Word document named “youraccountname_sql.doc(x)” .

HINT 1:

it is very easy to tell if a value is stored as a number or as text in an MS Access table.

-if the value is right justified (pushed to the right) then it is stored as a number.

-if the value is left justified (pushed to the left) then it is stored as a text.

HINT 2:

remember how field names with spaces are treated differently from field names without spaces.

HINT 3:

remember, round brackets are required to be used with all INNER JOIN statements.

The format (how the SQL statements are written) MUST match the style shown in the notes.

Each SQL reserved word MUST appear on their own line and in capital letters in the document. Each SQL Statement MUST be indented as shown below. You will lose major marks otherwise. example:

SELECT

something

FROM

( somewhere

INNER JOIN

somewhere else ON some condition )

WHERE

some condition is true;

This is non-optional. You MUST use this standard.

You will be graded on adhering to this standard.

You MUST write the SQL without the use or aid of any electronic method.

For example: You can NOT use MS Access Query Builder to create the SQL graphically and then copy or type in the result to your Word document. You will be graded on adhering to this standard.

You must identify yourself on the document. You will lose marks if this is missing.

Somewhere visible on the beginning (top) of the Word file you must include:

your first and last name

your Western ID (see below for a description of your Western ID)

your student number

Save your SQL statements in the Word file named “youraccountname_sql.doc(x)” and attach the file to your submission.

Project 2: Queries using Query Designer in Microsoft Access

Create a brand new, blank database and name it “youraccountname_Trains.accdb (.mdb)”. (substitute youraccountname with your actual account name.)

Using the XML Import utility in MS Access import the data in the file Assign5.xml.

This will import five (5) tables with data into your database. Everything is set by this XML and the schema. No further action or intervention is required by the students on the MS Access file.

This database represents a listing of PASSENGER(s), the destination(s), and the CONDUCTOR(s) for the Ampax Train Line.

It also contains information on which CONDUCTORS drive (trains) to which destination and for which destination the PASSENGERS have purchased tickets.

Create a query using the Graphical Query Design Tool in the Create tab in MS Access for each of the following and name each query as follows: QueryA, QueryB, etc. so they match:

Create a query that changes the conductor’s name with the CID of 001289348 to your first and last name.

Find the Province, TrainStation, Gate and Class for all of the DESTINATIONS whose TrainStations end with the designation of ‘1032’ and have a ‘B’ Gate.

Find the Destination Identification number (DID), Province, TrainStation and Gate for all of the trains taken by the Passenger with the PassengerID is equal to “0056348”.

Find the passenger number (PassengerID), first name and last name for all passengers who have bought tickets for a train to any TrainStation in Alberta. In addition to showing the passenger number, passenger first name and passenger last name, show the TrainStation and Gate as part of the resulting dataset. If a passenger has taken more than one train to Alberta that passenger will appear more than once in the resulting dataset.

Find the total number of tickets with the Destination of Alberta. Passengers who have taken more than one train to Alberta will be included in the answer for each train (ticket) that they have taken. Label this result field as: Alberta Trains

Find the passenger first name, passenger last name, Passenger Identification number (PassengerID), TrainStation, Gate and date of the train for all of the trains that you are the conductor. (use your last name and first name in the query not the CID). List the passengers in ascending order by the passenger’s last name.

Save your database in the file you named “youraccountname_Trains.accdb (.mdb)” and include the file to your submission to OWL.

Submission Instructions:

You must upload, attach and submit, via the CS1032 OWL/Sakai course site, the following 2 files:

youraccountname_sql.docx (for later versions) (.doc – for Word 2003 or earlier)

youraccountname_Trains.accdb (for later versions) (.mdb – for Access 2003)

NOTE: This description of the assignment contains instructions that tell you to create files with names that have a specific format. In these file names; the “youraccountname” is your university username.

AND AGAIN PLEASE: Do not cheat or copy.

When I ask students what I can do to limit the cheating, they recommend I remind the students that we do check the work and that there is a high probably they will be caught. So, here is that reminder.
