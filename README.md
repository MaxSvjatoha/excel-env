# excel-env
A private repo for managing excel data.

## 💡 General idea
There is a need to automate climate reporting at Danir and its subsidiaries. While the long-term solution would be to have an online user interface where each subsidiary could enter the required data, this is an intermediary step that acts as a proof of concept.

The idea is to read in climate reporting sheets from each subsidiary and input the data into a summary file, all in a non-hardcoded way so that subsidiaries can be added or removed at will.

Folder structure:
- excel-env
- - this repo
- Working folder (Arbetsmapp datainsamling)
- - Subsidiary 1
- - - Climate excel file, Subsidiary 1 (Klimatbokslut Subsidiary 1 - Datainsamling.xlsx)
- - Subsidiary 2
- - - Climate excel file, Subsidiary 2 (Klimatbokslut Subsidiary 2 - Datainsamling.xlsx)

...

- - Subsidiary n
- - - Climate excel file, Subsidiary n (Klimatbokslut Subsidiary n - Datainsamling.xlsx)
- - Summary sheet (Aktivitetsdata Klimatbokslut.xlsx)

## 🤖 Algorithm
0. Setup - load all folders, excel files and sheets
1. You read the single excel file inside each subsidiary folder

TODO: finish documentation
