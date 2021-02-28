# How to automatically populate PowerPoint presentations with Alteryx

If you find yourself manually updating the same Powerpoint presentations again and again, you may want to automate this process.

This repository contains examples and workflows that you can use to populate your own PowerPoint presentations using [Alteryx](https://www.alteryx.com/products/alteryx-platform/alteryx-designer). Instead of manually replacing values in tables or text boxes, you can use the macro [PPTX_macro_values.yxmc](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/PPTX_macro_values.yxmc).

## Prerequisites

* [Alteryx Designer 2019.3](https://www.alteryx.com/products/alteryx-platform/alteryx-designer)
* [7-Zip](https://www.7-zip.org/)
* You will have to change the file paths in the batch script workflows depending on your folder structure.

## Step-by-step guide for text boxes and tables

1. First, you will need a PowerPoint template. It should look exactly like the finished PowerPoint presentation you are trying to create, but all values will be replaced by unique strings. I usually pick {1}, {2}...etc.
The template for this project can be found [here](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/template.pptx). If you want to assign different colours to negative or positive values (eg red for negative, green for positive) you should highlight them in a colour that isn't used elsewhere in the presentation. In this case I used #FFC000. The macro will ask you to specify this template colour so it can easily pick up these values.

<p>
    <img src="01 Screenshots/slide2.PNG" width="600" height="200" />
</p>

2. Create a copy of your template, change the file extension to .zip and unzip the contents to a separate folder. This project unzips it to 02 Workflows\03_Template. This step can become quite tedious if you make lots of changes to your template and can be handled with a batch script in [Unzip template.yxmd](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/00%20Unzip%20template.yxmd).

3. Create a folder called 02_Replacement and copy and paste the contents from 03_Template. This will be used to generate the final .pptx file.

4. Next, create a mapping table with the columns *Name* and *Value*. *Name* contains your dummy strings whereas *Value* contains your real data. For simplicity's sake I have used a text input, but you will most likely generate this table with a separate workflow that generates your data.

<p>
    <img src="01 Screenshots/Mapping_table.JPG" width="600" height="300" />
</p>

5. Each slide is stored as .xml file in *03_Template\ppt\slides*. Import each slide you want to populate into Alteryx using .csv as data type in the 'File Format' field. 

* Uncheck 'First row contains field names'
* Use a pipe | as delimiter
* Set the string length to something very long such as 2540000
* Change the code page to Unicode UTF-8


6. Connect your mapping table to the bottom input anchor of the *PPTX_macro_values.yxmc macro* and the slide xml file to the top input anchor. Specify the highlighter colour described in step 1 and assign colour values for positive and negative values. Colours need to be entered as hex codes without #. If this is left blank, the font colour will be the same as in the template.

7. Output each slide as .xml file into *02_Replacement\ppt\slides*.

* Use .xml as file extension in the 'Write to File or Database' field, but use .csv as file extension in the 'File Format' field.
* Use \0 as delimiter
* Deselect *First Row Contains Field Names*
* Change the code page to Unicode UTF-8

8. Create a 04 PPTX folder and run the [PPTX_bat.yxmd](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/03%20PPTX%20bat.yxmd) workflow to generate your presentation. You will most likely have to change the file paths in this workflow according to your folder structure. This script zips up 02_Replacement, adds yesterday's date as suffix and renames it to a .pptx file.

## Step-by-step guide for graphs

1. Instead of creating dummy strings in your template you will need to pre-populate each graph with dummy numeric values which should be unique in the presentation. 

<p>
    <img src="01 Screenshots/graphs.PNG" width="600" height="200" />
</p>

2. For each graph you need to edit two files: an Excel file and the accompanying chart.xml. Input the Excel file from *\02 Workflows\03_Template\ppt\embeddings*. Input the chart xml file from *\03_Template\ppt\charts* with the following settings:

* Select .csv in the Field Format box
* Uncheck 'First row contains field names'
* Use a pipe | as delimiter
* Set the string length to something very long such as 2540000
* Change the code page to Unicode UTF-8

3. Connect the xml file to the top input anchor of the [PPTX_macro_values.yxmc] macro and your mapping table to the bottom input anchor. Overwrite the xml file in *\02_Replacement\ppt\charts*. Instead of using the macro, you can also use a *Summarize tool* to concatenate the two rows of your xml file and use the *Find & Replace tool* to replace your dummy data.

* Use .xml as file extension in the 'Write to File or Database' field, but use .csv as file extension in the 'File Format' field.
* Use \0 as delimiter
* Deselect *First Row Contains Field Names*
* Change the code page to Unicode UTF-8

4. Connect your Excel file to a select tool and change the data type of your dummy data from a numeric type to V_String. Connect it to the F anchor of a *Find & Replace tool* and connect your dummy data to the R anchor. Output the Excel file to *\02_Replacement\ppt\embeddings* and select *Overwrite (Drop)* in the *Output tool*.

5. Change the files paths in [PPTX_bat.yxmd](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/03%20PPTX%20bat.yxmd) and run the workflow to generate your presentation. If you want to create the .pptx file manually by zipping up the *02_Replacement* folder, you first need to navigate to *\02_Replacement\ppt\embeddings* and delete the .bak files which Alteryx created when it updated the Excel files.
