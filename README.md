# Worksheet-Maker-App-Thing

I built this application with tkinter to help automate some work I did for my old Job, it takes downloaded Excel files
that contain work order data and reorganizes that data into something I needed.

## The work that this project automated

I used to work in supply chain, and our warehouse management system wasn't always the best. There were things
it just simply didn't do and that created a lot of extra work for us. There were maybe different types
of items we would need for a work order, and these different items needed to be acquired at different times and in 
different ways by different people. The system did not categorize different types of items in the way that mattered to
us, it also didn't recognize our work flow process. Some items were needed for the pre-assembly team, other items on
the list were actually cable that needed to be cut. There were also items that were too large to be put together 
with the rest of the site kit. When the work orders would be kitted by the picker, the picker would have to 
ignore many of the items on his list and pre-assembly and the wire-cutter would have to find the work order in 
our system and manually look for the items they knew they needed to get.

## How this project automates that work

The work order data could be downloaded as an Excel file from our WMS, this application can read those files and
categorize items in whatever category the user sets. It can then take those specific items and write them on an
Excel file in a much more readable format than what our WMS would give us.

## How to demo this project

1. Clone this repo.
2. Run main_gui.py.
3. Click Material Finder.
4. Click Import (located in top left of GUI).
5. In the file browser select the demo_data.xlsx file which is located inside this projects directory.
6. Click Select Category button select a category(located in next the import button).
- The panel on the left has a tree view containing Job# and Site ID columns.
- When a row is selected from the left side tree view, the right side tree view will display the Materials and their
quantities for that site.
7. Click on a row from the left side tree view.
8. On the right side tree view, double click on a row and then select Add to Category
- Now all sites and rows containing the material that was added to the current category will be highlighted..
- To only the sites and rows that contain the material selected for your category, click on the switch that says Show Only Selected.
9. Click on the Review and print button.
- This page shows the available options for how the data will be layed out on the excel sheet.
- The drop down box on the top gives two options: By Material and By Site.
    - By Material will group all the same material together and display the site info to each row on the left columns.
    - By Site will group material with site it goes to.
10. Select By Material from the drop down box (located at the top of the window).
11. Select the Automatic radio button.
12. Click the load button (located on the bottom right).
- An excel sheet should be created and pop up on your screen.
- That covers typical use of this app. Thanks for checking this out!





