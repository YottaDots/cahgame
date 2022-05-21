# Cards Against Humanity
This project aims to collect all the cards, versions and sets (card decks) of the game Cards Against Humanity. We use the following source:
## Source code used
- https://docs.google.com/spreadsheets/d/1lsy7lIwBe-DWOi2PALZPf5DgXHx9MEvKfRw1GaWQkzg/edit#gid=13
- Spout https://opensource.box.com/spout/
- Last used and updated content is from 20-05-2022

## Installation
Just download the files and you have a full working example. If needed you can remove the spout code and use the installer from their side.

## Usage
``` text
- Go to the Google docs and download the file in the Excel format
- Open the excel and remove all the worksheets which are hidden
- Also remove de worksheets which are only informative (like index etc). Every deck starts with SET which is also the value the porgram searches for.
```

## How the script works
``` text
- First the script reads the file per sheet. Only visible sheets are taken into account.
- Second it walkthrough the sheets and searches for the word SET. That is the start of each deck. It calculates the mar columns per set and the max rows per set. Collects the deck name and so on.
- Third it goes through the sheet again and now collects all the cards and extra invformation like version or comments. Taking into account the start cell and the end cell (maxcolumn and maxrow).
-- Fourth it saves the information in the database
```
