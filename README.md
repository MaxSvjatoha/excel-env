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
1. Setup - load all folders, excel files and sheets
- 1.1 Load settings from settings json file
- 1.2 Load scope 2 special case dictionary
- 1.3 Load paths to each subsidiary folder containing the climate excel file (separately extract just the folder names)
- 1.4 Load climate excel files based folder paths
- 1.5 Load summary excel file and extract its sheets
- 1.6 Add mismatches sheet to summary excel file

2. Match sheet names (step 1.5) to input folder names (step 1.3)
- 2.1 For each input folder name, compare it to every sheet name and get a match score. Register the best match
- 2.2 Potentially filter doubles, which is handler to make sure that all matches are unique. If an item is already matched, let its match score decide whether to overwrite the previous match or not
- 2.3 Return dictionary in the format:

```
{
    'input folder name 1': {
    'match': 'best matching summary sheet name'
    'score': 'ratio_of_similarity_here'
    },
    'input folder name 2': {
    'match': 'best matching summary sheet name'
    'score': 'ratio_of_similarity_here'
    }
}
```

3. Read input data from each input excel file and store it in a dictionary using the input folder names as keys
- 3.1. This is done by scanning each excel sheet up to a max_row and max_column parameter, which are automatically set by finding the highest cell values which contain any information. The scanning is done in triplets, using the previous, current and next cell parameters to determine whether to read in the data at the cell and what key to use to register it. (see _get_scope_data() docstring)

4. Write data to summary sheet by using matches from step 2 and the scope 2 special cases dict
