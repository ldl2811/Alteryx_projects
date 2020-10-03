## How to automatically populate PowerPoint presentations with Alteryx

If you find yourself manually updating the same Powerpoint presentation again and again, you may want to automate this process.

This repository contains examples and workflows that you can use to populate your own PowerPoint presentations using [Alteryx](https://www.alteryx.com/products/alteryx-platform/alteryx-designer). Instead of manually replacing values in tables or text boxes, you can use the macro [PPTX_macro_values.yxmc](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/PPTX_macro_values.yxmc).

### Prerequisites

* [Alteryx Designer 2019.3](https://www.alteryx.com/products/alteryx-platform/alteryx-designer)
* [7-Zip](https://www.7-zip.org/)
* You may have to change the file paths in the batch script workflows if your folder structure is different.

### Step-by-step guide for text boxes and tables

1. First you will need a PowerPoint template. It should look exactly like the finished PowerPoint presentation you are trying to create, but all values will be replaced by unique strings. I usually pick {1}, {2}...etc.
The template for this project can be found [here](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/template.pptx). If you want to assign different colours to negative or positive values (eg red for negative, green for positive) you should highlight them in a colour that isn't used elsewhere in the presentation. In this case I used #FFC000. The macro will ask you to specify this template colour so it can easily pick up these values.

![Template](https://raw.githubusercontent.com/lb930/Alteryx_projects/main/Populating%20pptx%20files/01%20Screenshots/slide2.PNG)

2. Create a copy of your template, change the file extension to .zip and unzip the contents to a separate folder. This project unzips it to 02 Workflows/03_Template. This step can become quite tedious if you make a lot of changes to your template and can be handled with a batch script in [this workflow](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/00%20Unzip%20template.yxmd).

3. Create a folder called 02_Replacement and copy and paste the contents from 03_Template. This will be used to generate the final .pptx file.

4. Next, create a mapping table with the columns *Name* and *Value*. *Name* contains your dummy strings whereas *Value* contains your real data. For simplicity's sake I have used a text input, but you will most likely generate this table with a separate workflow that feeds into your data.

![mapping table](https://raw.githubusercontent.com/lb930/Alteryx_projects/main/Populating%20pptx%20files/01%20Screenshots/Mapping_table.JPG)

5. Each slide is stored as .xml file in \ppt\slides. Import each slide you want to populate into Alteryx using .csv as data type. 

* Uncheck 'First row contains field names'
* Use  a pipe | as delimiter
* Set the string length to something very long such as 2540000
* Change the code page to Unicode UTF-8

6. Connect your mapping table to the bottom input of the PPTX_macro_values.yxmc macro and the slide xml file to the top input. Specify the highlighter colour described in step 1 and assign colour values for positive and negative values. Colours need to be entered as hex codes without #. If this is left blank, the font colour will be the same as in the template. Run the workflow.

7. Output each slide as .csv file into 02_Replacement/ppt/slides folder.

* Use .xml as file extension in the 'Write to File or Database' field, but use .csv as file extension in the 'File Format' field.
* Use \0 as delimiter
* Change the code page to Unicode UTF-8

8. Create a 04 PPTX folder and run the [PPTX_bat.yxmd](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/03%20PPTX%20bat.yxmd) workflow to generate your presentation. You will most likely have to change the file paths in this workflow according to your folder structure. This script zips up 02_Replacement, adds yesterday's date as suffix and renames it to a .pptx file.

### Step-by-step guide for graphs

COMING SOON!
