# Revision2Highlight

Convert revisions to highlight in Word document.

## What does this project do

1. open a .docx file and copy its content to a new document.
2. for each revision, if the type of revision is Insert/Delete/Moved, close the revision and set its color to green/red. If not, set it to yellow.
3. save the new document.

## Issues

1. if an exception occurs, Microsoft Word may not close properly.
2. styles library in the document will not be copied.
