
# List Just Parameter Names
(get-help %SCRIPT/CMDLETNAME%).Syntax | SELECT-OBJECT –ExpandProperty SyntaxItem | SELECT-OBJECT –ExpandProperty parameter | SELECT-OBJECT name

# List Parameters and their proerties (Description, Mandatory, etc)
(get-help %SCRIPT/CMDLETNAME%).Syntax | SELECT-OBJECT –ExpandProperty SyntaxItem | SELECT-OBJECT –ExpandProperty parameter