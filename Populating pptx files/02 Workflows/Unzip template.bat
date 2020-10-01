
set PATH=%PATH%;C:\Program Files\7-Zip

copy /y template.pptx template_copy.pptx

ren template_copy.pptx template_copy.zip

7z x template_copy.zip -o./03_Template/

del template_copy.zip


