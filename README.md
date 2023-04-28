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
- - Summary excel file (Aktivitetsdata Klimatbokslut.xlsx) <--- contains 1 sheet per subsidiary + summary sheet for all of them

## 🤖 Algorithm
**1. Setup - load all folders, excel files and sheets**
- 1.1 Load settings from settings json file
- 1.2 Load scope 2 special case dictionary
- 1.3 Load paths to each subsidiary folder containing the climate excel file (separately extract just the folder names)
- 1.4 Load climate excel files based folder paths
- 1.5 Load summary excel file and extract its sheets
- 1.6 Add mismatches sheet to summary excel file

\
**2. Match sheet names (step 1.5) to input folder names (step 1.3)**
- 2.1 For each input folder name, compare it to every sheet name and get a match score. Register the best match
- 2.2 Potentially filter doubles, which is handler to make sure that all matches are unique. If an item is already matched, let its match score decide whether to overwrite the previous match or not
- 2.3 Return dictionary in the format:

```
{
    'input folder name 1': {
        'match': 'best matching summary sheet name'
        'score': 'ratio of similarity here'
    },
    'input folder name 2': {
        'match': 'best matching summary sheet name'
        'score': 'ratio of similarity here'
    }
}
```

\
**3. Read input data from each input excel file and store it in a dictionary using the input folder names as keys**
- 3.1. This is done by scanning each excel sheet up to a max_row and max_column parameter, which are automatically set by finding the highest cell values which contain any information. The scanning is done in triplets, using the previous, current and next cell parameters to determine whether to read in the data at the cell and what key to use to register it. (see _get_scope_data() docstring)
- 3.2 Return dictionary with the format:

```
{
    'input folder name 1': {
        'scope 1 & 2':{
            'cell name 1' : data,
            'cell name 2' : data,
            ...
            'cell name n' : data        
        },
        'scope 3':{
            'cell name 1' : data,
            'cell name 2' : data,
            ...
            'cell name n' : data        
        }
    },
    'input folder name 2': {
        'scope 1 & 2':{
            'cell name 1' : data,
            'cell name 2' : data,
            ...
            'cell name n' : data
        },
        'scope 3':{
            'cell name 1' : data,
            'cell name 2' : data,
            ...
            'cell name n' : data        
        }    
    },

    ...
    
    'input folder name x': {
        'scope 1 & 2':{
            'cell name 1' : data,
            'cell name 2' : data,
            ...
            'cell name n' : data
        },
        'scope 3':{
            'cell name 1' : data,
            'cell name 2' : data,
            ...
            'cell name n' : data        
        }    
    }
}
```

\
**4. Write data to summary sheet by using matches from step 2 and the scope 2 special cases dict**
*NOTE: this part of the code needs some serious factoring AND a finalized scope 2 handling. See 'future work' section below.*
- 4.1 Initialize relevant variables
- 4.2 Go two layers deep into output dictionary from step 3, so 'input folder name' -> 'scope sheet'
- 4.3 Find the relevant sheet in summary excel file based on matches made in step 2'
- 4.4 For each 'cell name' in 4.2, scan the entire summart excel file sheet and find the matching cell name while keeping track of the current row and column.
- - NOTE: special rules for 0 or more than 1 matching cell names
- 4.5 Write the data that 'cell name' points to in the current row and (column + 1)
- - NOTE: special rules for handling certain entries in scope 2. Hardcoded write locations due to very differing names
- 4.6 Save the sheet as a new file. Name it based on "Output file name" settings.json and save to "Output file folder name"

## 🛠 Future work
Currently, the code **cannot** handle multiple offices per scope 2 sheet. This is due to limitations of the openpyxl library. Essentially, one needs to check whether a cell already contains a value and, if true, add to the value rather than overwrite it. This also needs special cases for strings and integers, since strings would require a ", " or similar in-between the additions.

A temporary workaround would be to sum up all offices at the very bottom of the scope 2 sheet *using the exact same keys* (cells to the left of the input values). That way, this would be the last values that the function reads and would overwrite anything previously written down with them.
