2020-09-30
set PATH=%PATH%;C:\Program Files\7-Zip

cd C:\Users\bezlui\Documents\Projects\Populating pptx files\02 Workflows\02_Replacement\ppt\embeddings
del *.bak

cd C:\Users\bezlui\Documents\Projects\Populating pptx files\02 Workflows\02_Replacement
7z a ..\Company_presentation.zip -tzip -r *.*

cd..
ren Company_presentation.zip Company_presentation_2020-09-30.pptx

move /Y Company_presentation_2020-09-30.pptx ..\"03 PPTX"


