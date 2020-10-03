## How to automatically populate PowerPoint presentations with Alteryx

This repository contains examples and workflows that you can use to populate your own PowerPoint presentations using [Alteryx Designer](https://www.alteryx.com/products/alteryx-platform/alteryx-designer). Instead of manually replacing values in tables or text boxes, you can use the macro [PPTX_macro_values.yxmc](https://github.com/lb930/Alteryx_projects/blob/main/Populating%20pptx%20files/02%20Workflows/PPTX_macro_values.yxmc). 

This macro uses xml files to change the values in a PowerPoint presentation

### Step-by-step guide

1. First you will need a PowerPoint template. It should look exactly like the finished PowerPoint presentation you are trying to create, but all values are replaced by dummy data strings which are unique in the presentation. I usually pick {1}, {2}...etc.

![Template](https://raw.githubusercontent.com/lb930/Alteryx_projects/main/Populating%20pptx%20files/01%20Screenshots/slide2.PNG  =250x250)
