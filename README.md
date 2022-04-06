# Compare BOM
This repository is home to the python script that will compare two BOMs.  This script's original intention is to compare an Engineering BOM against an IFS BOM.   

# Description 
Within the directory that this script resides, there shall be two different excel files.  One should contain **ENG** in the tile, while the other should contain **IFS** in the title.  If the title is slightly off, the script won't be able to _automatically_ distinguish ENG vs. IFS, however, the user will be prompted to define which is which.  

This script will only operate on .xlsx files, and not .xls files. This script will automatically sift through every sheet of each file, but each file shall only contain one sheet with BOM data.  For example, it is common on Engineering BOMs to include a revision or changelog sheet.  This script is intelligent enough to omit sheets that do not contain BOM data.  

Each BOM _shall_ contain headings: __QPN__ | __QTY__ | __DES__ | __REF__ 

In most, if not all cases, IFS BOMs name the _REF_ field _Notes_.  The script will look for the _Notes_ column, and substitute that field for _REF_.  Furthermore, some engineering BOMs will not contain a _REF_ field.  This field can be added, or, so long as there's a _Notes_ field in the engineering BOM, the _Note_ field can be used in place of _REF_.  

Subtle discrepancies will be accepted.  For example, _Des_, _DES_, _Description_, etc. will be accepted as heading __DES__.  Since the application automatically locates the location of various data columns, it needs to seek out this header before starting. Locating the header is what's critical.  This is to say, that the entire column can be blank under a particular header.  For example, the user may wish to add a _REF_ column just to facilitate proper operation, although no reference values exist.  

# Revisions
v1.0 -- Initial release.   

v1.1 --  When searching for the "reference" fields, the "notes" field is no longer included.  Therefore, "notes" is no longer a substitute for "reference".  This means that BOMs shall have a "reference" field.  The reason for this change is that boms containing both "notes" and "reference" fields would produce errors.   

v1.2 -- Made logging more verbose so it is easier to identify BOM formatting issues.  It is no longer required that there be a REF field in the BOM -- which is good for BOMs that include cable drawings, which do not require a reference field.  

V1.3 -- BOM descriptions are no longer fixed at ENG or IFS, but rather, the user can enter a short description based on the filename.  For example, he/she may wish to enter A02 and A03.  