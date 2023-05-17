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







