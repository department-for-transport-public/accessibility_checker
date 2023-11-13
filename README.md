# Spreadsheet Accessibility Checker

This is the repository for the R code for the ODS/XLSX spreadsheet accessibility checker. This Shiny app is designed to be hosted to allow users to check their spreadsheets in an interactive web application, against a list of accessibility points. 

Once launched the Shiny app will look like this:

![image](https://github.com/department-for-transport-public/accessibility_checker/assets/84339173/7bf1c444-a602-46fe-8a12-d5a05041ddda)


Users can select either the ODS or XLSX tabs depending on the format of their files. They can then use the file uploader format to import one or more files into the app. It will then check the content of each file against the various accessibility checks; this may take a little while especially if files are large.

The app will then display the results of the checks across the different tabs.

## To set up for hosting 

To use this app in your own infrastructure, fork and clone it from Github. The libraries used are version controlled using renv.

You will need to host it on a system which has a backend server (e.g. shinyapps.io or a Posit connect instance).

Once hosted, the app does not save or copy data to any other location, the data is not stored in the app beyond the end of a session and is not accessible to any other user.
