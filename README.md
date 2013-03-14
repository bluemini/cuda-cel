cuda-cel
========

A project to investigate the conversion of an excel spreadsheet into a CUDA kernel for parallelising over a grid. Currently uses the Excel 2003 XML format as an intermediate format, which is parsed by SAX.

The SAX parser builds a DAG structure of populated cells by workbook, named cells and named ranges.

Any output location in the DAG can be chosen and evaluated. If the cell contains a formula, PLY is used to parse it and decompose it into appropriate references to other cells. Effectively, the complete underlying calculation of the chosen cell location is unwound to references and functions of known values (or not).

- Experimental
