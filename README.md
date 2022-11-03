# KiCad_Large_Part
Develop a KiCad Schematic Symbol for parts with multiple bodies and large number (>100) of pins.

Project: Large Component Builder python program (LCB)

Purpose: Use a spreadsheet to define the blocks and pins of a large (>100 pin) device to generate a schematic symbol broken into blocks. A stretch goal would also define the PCB footprint for same. Examples to operate on would extend from large memory devices to the ST and NXP classes of processors. The spreadsheet should have a separate page (tab) for each functional block of the target device, with an extra “top page” which would have common definitions, including configuration table for parts such as the NXP or ST Processors which have widely varying configurations.

    1. Call spreadsheet name with program invocation
       
    2. Read spreadsheet Top – This sheet does NOT invoke a body, but defines all characteristics of the rest of the component. Will include things like:
        1. Overall Part Number
        2. Manufacturer Data if appropriate
        3. PCB Data if appropriate
        4. ? Textual Data for the part (need to think about this one…)
           
    3. Each tab of spreadsheet thereafter is a BODY – Each BODY has the characteristics:
        1. Pin Name
        2. Pin Number
        3. Direction (L, R, T, D)
        4. Pin Origin 
        5. Pin Type

Files: 
    1. Input file: Pdef_{Component Name}.xlsx – Spreadsheet defining the part.  See Appendix II for format. Can be in OpenDocument form but must be converted to .xlsx for python / openpyxl to recognize.
    2. Output file: {Component Name}.kicad_sym – (ie: strips off Pdef_)  In same (working) directory.

