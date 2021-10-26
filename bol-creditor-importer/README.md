# bol-creditor-importer

Deployment:
- From the Spreadsheet where this script will be run, from  the `Tools > Script editor` menu item.
- Replace the contents of the pre-existing `Code.js` file with that of `Code.js`.
- Add a new script file and name it  `consts.js`. Copy the contents of `consts.js` to this file.
- Save all files and close the script editor.

Usage:
- Open the Spreadsheet to which the script has been added.
- Run the importer by following the `Add-ons > BOL Creditor Import txt Generator > BOL Creditor Import txt Generator` menu item.
  - If this is the first time running the script, an authorization window will appear. You must authorise this script to run on your behalf.
  - It will also be necessary to run the script from the menu item a second time if you have just confirmed authorization.
- A dialog window will appear. In the input field, enter the batch number for which to fetch the data. Preceding 0s are required. 
- Press `OK`. Then check the folder that you added its ID to `consts.js`

Customisation:
Manipulatable variables are stored within the `consts.js` file as you want.

