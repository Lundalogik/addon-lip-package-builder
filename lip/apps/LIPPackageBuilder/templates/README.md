# <*displayName*>
<*A brief, selling and non-technical, overview in just a few sentences.*>

<*description*>


## Requirements
<*A list of requirements on servers, client versions, other add-ons or whatever it can be.*>

<*cloudCompatible*>


## Features
<*Describe important features of the add-on using sub headers like below.*>

### Feature 1
<*This is a description of feature 1 and what it does, why it's good etc.*>


## GDPR

### What is done with personal data
<*Display/store/use*>

### Why personal data is needed
<*Answer why the add-on needs the personal data it uses.*>

### What personal data
<*List data used in some way that is considered to be personal data.*>

### Transfer of personal data
<*Is personal data transferred to other systems? If so, how is it done?*>

### Logging of personal data
<*What, where and when is personal data logged? When are logs deleted?*>

### Possible avoidance of using personal data
<*Is it possible through the configuration of the add-on to avoid using personal data? If so, how and are there any drawbacks?*>


## Installation
<*Step-by-step instructions on how to perform a new installation of the add-on. Examples of bullets are shown below. Note that the bullet showing how to initialize the LBS app should contain a complete configuration that works with the latest Lime Core database.*>

1. Run SQL procedure scripts
2. Restart the LDC manually (right-click on it and click "Shut down").
3. Restart the Lime CRM desktop client.
4. Download the LIP zip file from the latest [release](https://github.com/Lundalogik/addon-documentsearch/releases). <*Update the link to releases in the correct repository*>
5. Enter the VBA editor in the desktop client and run ```vba lip.InstallFromZip``` from the immediate window. Select your zip file.
6. Compile and save the VBA project.
7. Add the "DocumentSearch" folder from folder apps to the Actionpads\apps folder.
8. In the main actionpad, where you want the app to be shown: Insert one of the two following ways of showing the app.

```html
<!-- Use the below for usage without expandable header. -->
<div data-app="{app:'DocumentSearch', config:{}}"></div>
```

```html
<!-- Use the below for usage with an expandable header. -->
<ul class="menu expandable">
    <li class="menu-header" data-bind="text: localize.Addon_DocumentSearch.i_menuHeader"></li>
    <li class="divider"></li>
    <li id="documentsearch">
        <div data-app="{app:'DocumentSearch', config:{}}"></div>
    </li>
</ul>
```


9. Publish the actionpads.
10. Restart the Lime CRM desktop client and start using the add-on!
11. Add a customization record in Lime's Lime CRM under the customer. Note the version installed (can be found in the app.json file). Link it to the product card.


## Configuration
<*Explain all the configuration parameters that can be set in the LBS app config or in a VBA module. Limit use of VBA configuration for passwords or other sensitive information that we do not want in files in the Actionpad folder).*>


## Upgrading
<*Step by step instructions on how to upgrade an existing installation of the add-on. How to do it to not leave traces of old code or files, how to keep the configuration of the add-on etc.*>


## Troubleshooting

### One header per type of common pitfalls or errors
<*Description of the problem and a solution for it.*>


## More Information
<*Are there any more places where information that could be of help can be found? Add links and describe what they are useful for.*>
