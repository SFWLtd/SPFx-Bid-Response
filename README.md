## newbid-webpart

SharePoint Framework web part utilising the React Framework to create a folder structure for a bid, using user defined parameters as metadata. Links with the SPFx New Organisation web part.

Done:

* Submit to document library with a folder structure
* Metadata correctly applied to the folders

To be done:

* Create search of existing organisations
* Rework text fields defining people to be people pickers from the Office UI Fabric component

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* commonjs components - this allows this package to be reused from other packages.
* dist/* - a single bundle containing the components used for uploading to a cdn pointing a registered Sharepoint webpart library to.
* example/* a test page that hosts all components in this package.

### Build options

gulp nuke - TODO
gulp test - TODO
gulp watch - TODO
gulp build - TODO
gulp deploy - TODO
