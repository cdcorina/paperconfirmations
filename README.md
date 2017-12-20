# paperconfirmations
Why Excel2Xslt?

Most clients give their specifications in Excel format. For paper confirmations the structure of the confirmation is already in Excel; why not generate the .xsl stylesheet directly from Excel?  

The advantages of using Excel2Xslt

For each sheet in the workbook the code will generate a corresponding Sheet_name.xsl file, that can be easily uploaded in your chosen back-office tool.

It reads each cell of  the sheet. The code for each row and cell, as it is in Excel, will be put in the stylesheet, where a table will be created. If  the text of the cell starts with ‘MailBody/’, then the code will know that this is a tag. If the first cell in any row starts with choose when or otherwise or end, then the code will treat this as part of a choose command (see choose command syntax in XSLT). Choose is similar to a case structure that allows you to have different values in different cases.

If the cell in Excel spans on multiple columns, then it will also span on the respective number of columns in the xsl stylesheet. (the code recognizes a MergeArea).

It reads the value of  each cells’s background color and put the corresponding hexadecimal value of the color in the .xslt stylesheet. 

It generates the Constant_ExcelFileName.xml file that will contain all the constants needed for the name of the MailConfigurations used. The file can be easily uploaded into your chosen back-office tool.

It generates the MailHeaderCfg_ExcelFileName.xml file that will contain the MailHeaderCfg configuration for each sheet in the workbook. The file can be easily uploaded into your chosen back-office tool.

Example:

For exotic options, you might have as many as 60 different templates for different types of options. (See the attached specifications file Specifications_Options.xls). Using Excel2Xslt, the number of templates used is only 4. See the Excel file created Options.xls. This file was created by replacing the constant values as specified by the client with MailBody/Tag-name. Also, the choose option was used to reduce the number of templates and simply enumerate the different cases and their corresponding values.
