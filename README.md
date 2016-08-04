
Azure ML Excel Add-in Starter Kit
===================


**Summary**
----------

Azure ML Excel Add-in Starter Kit demonstrates how you can quickly build an Excel Add-in that integrates with your published backend Azure ML based web service to perform predictive analysis of your Excel data.


**Applies To**
----------

 - Office Add-in

**Try it Out**
----------

Azure ML Excel Add-in Starter Kit Visual Studio solution has 2 projects. The first project is an Add-in project that contains the Add-in Manifest XML file. The second project is a pure HTML/JS/CSS web project.

When the solution is run, a sample Excel Workbook that contains sample income data is displayed along with the add-in. 

 1. Enter a valid URL that points to the backend Azure ML Web Service
 2. Enter the API Key for the backend Azure ML Web Service 
 3. Highlight rows and columns from Case # - Income in the table including the table header
 4. Click the "Predict" button

---- 
The Add-in will upload/post the selected data to your Azure ML Web service entered in the URL The Add-in will update the spreadsheet with the Scored Label and the associated Scored Probability results returned from Azure ML Web service. Excel chart is also refreshed to reflect the latest Predicted Score along with the Probability risk.
