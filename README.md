# LIP Package Builder

LIP Package Builder is a tool for creating LIP packages or new add-ons easier.

Through an easy-to-use GUI you can select your components for the LIP package. Click Generate Package to create a LIP zip file as well as everything you need when you are creating a new add-on.


## Requirements

* Lime Bootstrap framework.


## Features
The output when creating a new package is a folder containing two things:

1. A folder called *add-on* that will contain everything you need to put in the GitHub repository if you are creating a new add-on. This follows the requirements for Community Add-ons. This folder is only generated if you select "Add-on" in the Package Builder GUI on the Details tab.
2. A LIP zip file of the content of the folder *add-on\lip*. This is the zip file that you should use to install the package in another Lime CRM application. The file package.json in the zip file follows the structure described in the [LIP repository](https://github.com/Lundalogik/lip).


### Supported Components
The following objects in a Lime CRM application can be selected in the GUI:

* Tables and fields
* VBA modules
* SQL functions and procedures
* Localizations
* Actionpads

#### Limitations
Some things that are not supported by the LIP installation of a zip file are exported to separate folders (according to the Community Add-ons requirements) for manual installation:

* SQL procedures and functions
* Option queries
* SQL expressions, SQL for update and SQL for new on fields.
* Table descriptive expressions
* Table icons


### Open Existing Files

#### LIP Package zip
It is possible to open a previously created LIP zip file. The Package Builder will show which components that were part of that zip file. You can then add new or remove previously included components to create a new version of the package/add-on.

#### CHANGELOG.md
*Only when working with an add-on:* You can upload an existing CHANGELOG.md and then simply choose how to increment the version number and if you want to use the same author as for the last version. The versioning info written in the form will be inserted into a copy of the uploaded CHANGELOG.md file when the Package Builder generates the add-on files for you.

#### metadata.json
*Only when working with an add-on:* You can upload an existing metadata.json. The form will automatically be filled with data from the uploaded metadata.json.


### Possible Future Features

* Select dependencies to other add-ons.
* Select Lime Bootstrap apps to include.


## Installation

1. Add the SQL procedures to the database.
2. Run `EXEC lsp_setdatabasetimestamp` and `EXEC lsp_refreshldc` on the database.
3. Restart the LDC.
4. Restart the Lime CRM Desktop Client.
5. Enter the VBA editor and run `lip.Install("LIPPackageBuilder")` in the immediate window.
6. Compile and save VBA.
7. Publish Actionpads.
8. Run `LIPPackageBuilder.OpenPackageBuilder` in the immediate window in the VBA editor to start the Package Builder.


## How to use
1. Open the Package Builder GUI by running `LIPPackageBuilder.OpenPackageBuilder` in the immediate window in the VBA editor. Note that the user you are logged in as need to have a coworker card for it to work.
2. If you want to create a new version of an existing package, click the Open Existing Package button and select the relevant LIP zip file.
3. Enter the information on the Details tab.
4. Select the objects that should be included in the package.
5. Click the button "Create Package". Choose where to save the generated files. The folder created will be opened in a Windows explorer.
6. *If working with a new version of an add-on:* Add the content of the *add-on* folder to a GitHub repo.
7. *If using it as a launching tool:* Logon to the target Lime CRM application and use LIP to install your package by pointing out the generated zip file.


## Upgrading
*More info coming...*


## Troubleshooting
See [issues](https://github.com/Lundalogik/addon-lip-package-builder/issues) under the GitHub repository. If you cannot find an answer there, create a new issue and describe your problem there.
