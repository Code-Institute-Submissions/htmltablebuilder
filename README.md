# HTML Table Builder <img align="right" width="75" height="75" src="favicon.ico">

LLC denotes Local Link Cork
-
---------------------------

## Site Goals

HTML Table Builder is an application developed in Python for LLC staff to auto generate html table code for Timetables populated and edited on Google Sheets.
This application is an assist for LLC staff engaged in Timetable Management to easily publish timetable revisions to the LLC website.   

HTML Table Builder is useful by allowing LLC staff to manipulate and design timetables to best display the data to the public.

Within some pre-defined parameters, they can design the timetable layout to accommodate peculiarities on some routes.
Excel is a familiar environment for LLC staff to work with. Google Sheets has a lot of similarities and of course LLC staff can copy and paste their Excel creations to Google Sheets.

Important Note: All references to Local Link are for Local Link Cork only.

Google Sheets              |  Python Utility
:-----------------:|:-----------------:
![](assets/images/sheets.webp)  |  ![](assets/images/python.webp)

# UX/UI & Features

## Design Choices

This is a non GUI application.

The only interaction with the user is a simple yes/no input like this: user_input = input('Do you wish to proceed (yes/no): ')
This gets verification from the user whether to proceed or not in running the application.

The user is advised of the consequences of running the application like this:
print("Running this Automation Program will overwrite, your previous results")
print("But be assured your Timetable Sheets will remain untouched.")

While the application is running the user is kept informed of progress with messages like this:
print("Preparing data from worksheets...")
print("Data ready to start creating HTML Table code!\n")

When the application has finished running the user is advised of this:
print("The program is finished executing!")

The user is also advised to look at the completed work on Google Sheets.
print("Take a look at Google Sheets to view your HTML Table code.")

The user is instructed on how to proceed with Google sheets and Wordpress.
print("Just copy the contents of Cell A1 in the HTML worksheet.")
print("Then Paste into matching Wordpress Schedule Post.\n")

## User Stories

- Users who have no web programming skills need to be able to publish LLC timtables to the website.
    - This application allows the user to do exactly that.
- Users prefer to work in a familiar environment when manipulating timetables.
    - By using Excel and or Google Sheets the user can work with familiar interfaces.
- Users need to be able to change Timetable data on a regular basis and get the results published without delay.
    - The level of automation here minimises the time delay.
- Users need to be able to redesign timetables like adding an extra column, or putting header text into 2 lines.
    - This application can cater for that.
- Users need to be able to highlight certain columns to draw attention to detail.
    - This application allows for that.
- Users wish to avoid steep learning curves by being introduced to new systems.
    - The Copy and Paste method to publish to the web is a concept they are already familiar with.


## Site Navigation
 
As there is only one option no navigation is required

---- 

===============================================================

### How it works

===============================================================

The user sets up Google Sheets prior to running the application

![](assets/images/sheets.webp)

The user can setup as many worksheets as they wish. But typically there would be only a few timetables requiring update at any given time.

The user renames the worksheets to match the Route Number.

Where the timetable is an amalgamation of several routes then the user labels the worksheet with all of the route numbers seperated by a hyphen.

![](assets/images/multipleexcel.webp)

The Basic Rules for Excel Template are provided in the Rules tab. 

![](assets/images/rulesexcel.webp)

These same rules apply to Google Sheets other first column width is 400 not 40 and other column widths are 100 not 10.

When the user is happy that the necessary Google Sheets are in place then they can run the HTML Table Builder Application.

---------------------------------

## Application


![](assets/images/python.webp)

----------------------------------

### Overview

HTML Table Builder Application runs through all the worksheets in Google sheets other than HTML prefixed sheets.

Looping through each sheet the table html code is created and written to a local txt file.

The table elements are gathered in 4 distinct groupings:

1. Table definition html code
2. Table header html code
3. Table Rows html code
4. Table footer html code

When all elements of the table html code have been compiled the txt file is then read.

The files contents are written back to a newly created worksheet with a HTML prefix and the Route number.

----------------------------------



Clonakilty Pickup                 |  Kinsale Pickup
:-----------------:|:-----------------:
![](assets/images/destinationclonakilty.webp)  |  ![](assets/images/destinationkinsale.webp)

Important Note: The first stop is not included, because a passenger cannot be Picked Up and have as the Destination the same stop.

Clonakilty Pickup Stop                |  Clonakilty available Destination Stops
:-----------------:|:-----------------:
![](assets/images/pickstop.webp)  |  ![](assets/images/destinationstops.webp)

Important Note: A Passenger cannot have a Destination that is earlier than or the same as the Pick Up Point on the route.
The Destination Select responds by only making available Stops that are later than the Pickup stop.

The user can select what passenger type they are like for example "Adult".

The user cans elect what ticket they are looking to purchase like for example "Weekly Ticket".

Passenger                |  Ticket
:-----------------:|:-----------------:
![](assets/images/passenger.webp)  |  ![](assets/images/ticket.webp)

Then by clicking the Calculate Fares button the correct fare is calculated based on the user selections and the fare is presented on the form.

![](assets/images/correctfare.webp)

----
----

# Testing

## Tests carried out by me.

- HTML Prefix title on worksheets are ignored.
- Worksheet is not a blank new worksheet (Title:Sheet 1)
- Worksheet title begins with 3 numeric digits.
- Worksheet has no header row.
- Header Row
    - Test with only 1 column
    - Test with Last column not having HEADEND aas text value
    - Rules in Last Colum
        - : rule marker but invalid rules
        - : rule marker but no rules
        - : rule marker with column counter descending
- Only a Header Row, No other rows.

## Validator Testing

- HTML:             All pages were passed through the official https://validator.w3.org/ and no errors were found.
- CSS:              All pages were passed through the official https://jigsaw.w3.org/css-validator/ and no errors were found.
- Accessibility:    By running the site pages through Lighthouse in Inspect on Chrome I got the following results:

index desktop                |  index mobile
:-----------------:|:-----------------:
![](assets/images/desktopindex.webp)  |  ![](assets/images/mobileindex.webp)

single desktop               |  single mobile
:-----------------:|:-----------------:
![](assets/images/desktopsingle.webp)  |  ![](assets/images/mobilesingle.webp)

timetable desktop               |  timetable mobile
:-----------------:|:-----------------:
![](assets/images/desktoptimetable.webp)  |  ![](assets/images/mobiletimetable.webp)

- Javscript:        All javascript was passed through the official https://jshint.com/ and no errors were found.
    - jshint did provide these warnings.

    ![](assets/images/jshint.webp) 

Note: As all async functions are operating correctly I assume that I can ignore these warnings.

----

## Development Transition

### Initial Wireframe Concept
 

## Bug Fixes
 
### Solved Bugs

- August 28th Problem as index page had a capital letter Index.html.
    - Fix e96c69b2fcd38294c0075cdd66a0dbcddcca3e5a changed to index.html.

- September 1st Problem with json populating Drop Select on Fares Calculation on DOM Load.
    - Fix 8c56c5f7d2a5d92af624c5578bd6dc4aab10ef4f fixed DOM load to populate Drop.

- September 1st Problem when fare is not found in fares json file for given parameters.
    - Fix c63f1f924202060d3a47a30489ee205802aa5629 fixed by adding a Try and Err trap.  

- September 1st Change direction event listener not populating Pick and Drop selects properly.
    - Fix 7cabaa36155e59fed16446a49c6fd3ab3cf9212f fixed event listener code. 

- September 12th Problem with some event listeners firing on pages other than single.
    - Fix 8c687082c717c84a70b63f1eabd4173f4bafb202 fixed with conditional if. 

- September 12th Problem with closing div on pu-container.
    - Fix 07f5f0efd84c0c1947c0dba8dbbcd9bd211a6d5f add in closing div.  

- September 12th Problem with stray tag on New Search button on timetable.html.
    - Fix c6f4f81fdde1768055a3d3ce3221e17a56d076ec remove tag.  

- September 13th Problem Dom loaded not firing select route population by json.
    - Fix 9b331f9e6cf1cad47f8c6608567760b5a5260fdb add in a window load listener. 

- September 13th Select route population by json needs a conditional so it only fires on index.html.
    - Fix 015683d84575d2fe6c69a21553ff0076a890b9b4 add in a conditional if. 

- September 20th Fix error in title iframe on index.html.
    - Fix 4c893f408c664917484bdb347ca265eff02b8729 removed typo. 

### Unfixed Bugs

- No known unfixed bugs

----

## Deployment

### Deployment to GitHub

The site is deployed to GitHub pages. 

- Git status check in Gitpod to ensure all is pushed to GitHub.
- In the GitHub repository under Recent Repositories select [TMartin88/farescalculator](https://github.com/TMartin88/farescalculator)
- Then Settings.
- Then Pages.
- Under Build and Deployment select Deploy from a branch from the Source dropdown list.
- Pick main and then root from the Branch selection area.
- Now GitHub Pages displays a link to the live site.

The live link to the site on GitHub is: [Fares Calculator](https://tmartin88.github.io/farescalculator/)

### Deployment to work domain

This project is also deployed on a live domain for the company I work for part-time. The website on this domain is developed by me in Wordpress using Elementor Pro site builder and Custom Posts (basically each route is a custom post).

The goal of this site is to provide an online search platform which allows interested parties such as passengers to search through all of the Local Link Cork routes.

Fares Calculation is simple for most routes but for some high frequency routes like the 253 it is very complicated and difficult to present in a user friendly way. To make an easy to understand interface for passengers really required something like Javascript so that was the inspiration for me to do this project 2 in the way I did it.

This Project is very useful for Local Link Cork for certain Routes like the 253. The code for fares calculator is the same apart from a few tweaks that are wordpress/elementor specific.

This Fares Calculation element will also be applied to other routes similar to the 253 in the Wordpress implementation.

The link to the site live on a domain is: [253-monday-to-sunday](https://locallinkcork.com/schedule/253-monday-to-sunday/)

----

## Future Features

- To include all routes run by Local Link Cork.
- To allow the user to free type place names in Search and present a matching Search results page.
- To migrate all Excel data to a database.
- To enhance the VBA json generating code to pick up data from database and convert the data to meaningful json.
- To link into Irish Rail and Bus Eireann Fares Calculators.
- To incorporate a QR Code creator to link directly to route urls.
- To put in Extensive Error trapping. 

----

## Performance Improvements

- To get all lighthouse results close to 100%.

----
 
## Credits

### Inspiration

- Working part time with Local Link Cork and in association with Steve Ellis we identified that passengers could not get easy access to our route details.
- Also that the fares rate tables were largly incomprehensible to the regular person.

The challenge is to take data stored in Excel files and bring them to the public.

And so the Fares Calculator was conceived. This project is hopefully just the beginning.

### Content

- Credit for Fetch API to [javascripttutorial](https://www.javascripttutorial.net/javascript-fetch-api/)
- Credit for Filter Array ideas to [w3schools](https://www.w3schools.com/jsref/jsref_filter.asp)

### JSON data

- Thanks to Steve Ellis Operations Co-Ordinator of Local Link Cork for permission to use logo, video, data and pdf.

----




Credits:

https://thispointer.com/how-to-append-text-or-lines-to-a-file-in-python/

https://pynative.com/python-delete-files-and-directories/#h-example-remove-file-in-python

https://www.pythonpool.com/python-loop-through-files-in-directory/

love sandwiches


