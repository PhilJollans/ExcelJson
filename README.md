This project is an Excel add-in with some specific features to read and write json files.  

The project now has two specific features.

# Read and write an array of simple objects

The json file must contain a an array of objects, and no nested objects.
The property names will be imported to the first row in the excel sheet.
After that, each row in the excel file represents a single object from the json array.

The worksheet can be exported to json with the same logic.
The first row conatins the property names.
The following rows each define one json object.

# Read and write json files used by Angular i18n

If you are localizing an Angular project using the built in Angular i18n support, you can choose to use json instead of the default XLIFF format.

For example, the command
```cmd
ng extract-i18n --output-path src/locale --format json
```
generates a file `messages.json` in the directory `src/locale`, from which you can generate language specific files, for example `messages.de.json.

That is easy to do once, but hopeless when you regenerate the file `messages.json` with new (or removed) texts and need to merge the changes into specific language files.

This feature has two commands:
 - Read i18n
 - Write i18n

## Read i18n

The read command starts with a file open dialog, where you select the file `messages.json`.  
It reads this file, and any language specific files with a name like `messages.???.json`, into a worksheet, with
 - the language tag in the first row
 - the property name or key in the first column
 - each language in a separate colum

This is approximately what it looks like.

| key        | de        | en         |
|------------|-----------|------------|
| examples   | Beispiele | Examples   |
| statistics | Statistik | Statistics |
| play       | Spiel     | Game       |
| reset      | Reset     | Reset      |
| yes        | Ja        | Yes        |
| no         | Nein      | No         |

## Write i18n

The write command starts with a file save dialog, where you again select the file `messages.json`.

It does not overwrite the file `messages.json`, but it does generate a file `messages.???.json` for each additional column in the worksheet (starting at column C).

Note:
 - If you enter a new file name (where the file does not yet exist), for example `messages-copy.json`, then it will generate this file, as well as the language specific files.
 - You can add new columns for new languages, which will then be exported as new files.