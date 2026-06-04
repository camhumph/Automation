# Automation

## CMS XT Export macro fixes

The repository did not include the original SolidWorks VBA macro source. The file
`CMS_XT_Export_DropIn_Fixes.bas` contains drop-in replacements for the sections
discussed in the task:

- J BLOCK DXF export now uses a temporary native `.sldasm` instead of reopening a
  temporary `.x_t`, preserving `CMS_TOP` / standard-view orientation.
- The DXF 1:1 setting is absolute for every DXF, including BASE and HOLDERS.
- Drawing sheet scale and individual drawing view scale are both forced to 1:1.
