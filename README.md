## org-chart

SharePoint Framework (SPFx) webpart to display organization hierarchy.

**Big tiles**
![SPFx-org-chart-big-tiles](https://github.com/Vansyork/SPFx-Org-Chart/blob/master/readme-images/Aantekening%202019-10-25%20144725.png?raw=true)

**Small tiles**
![SPFx-org-chart-small-tiles](https://github.com/Vansyork/SPFx-Org-Chart/blob/master/readme-images/Aantekening%202019-10-25%20145206.png?raw=true)

### Building the code

  

```bash
git clone the repo

npm i

npm i -g gulp

gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts

* dist/* - the bundled script, along with other resources

* deploy/* - all resources which should be uploaded to a CDN.

  

### Build options

#### Create Development server

```bash
gulp serve
```
Use this url to test on any sharepoint Online site:
*/_layouts/15/workbench.aspx*

#### Create .sppkg package
```bash
gulp bundle --ship

gulp package-solution --ship
```

###  Configurations

#### Config SharePoint list
 - There is a config list deployed as default for you to configure, add
   items to the list to start building your organizational chart.
 - Start with adding a few items before setting the **My Reportees** field.
 
 **Config list**
 ![SPFx-org-chart-big-tiles](https://github.com/Vansyork/SPFx-Org-Chart/blob/master/readme-images/Aantekening%202019-10-25%20141134.png?raw=true)

 #### Webpart Property Pane configurations
 
|Setting |Description  |
|--|--|
|Use AD data to build the org chart |Use the Microsoft Graph API to generate your organizational tree.|
|Select Org Config List|Select a config list to generate your organizational tree.|
|Select user to start building the Org-Chart from the config list|Select a user from the selected configuration list to use as starting point for your organizational tree.|
|Select user to start building the Org-Chart from AD data|Select a user from the AD to use as starting point for your organizational tree.|
|Use small tiles|Use only pictures/persona to display the nodes|
|Create Configuration List button|Will display a dialog to create a new Configuration list |

**Webpart properties**
![SPFx-org-chart-big-tiles](https://github.com/Vansyork/SPFx-Org-Chart/blob/master/readme-images/Aantekening%202019-10-25%20145442.png?raw=true)
