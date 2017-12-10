using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Xsl.Runtime;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Xml;
using System.Diagnostics;
using System.Configuration;

namespace ExcelExample
{

	public class ExcelClass 
	{

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args) 
		{

String[] Color = new String[57];
int file_counter = 0;
			
Color[1] = "#000000";
Color[2] = "#FFFFFF";
Color[3] = "#FF0000";
Color[4] = "#00FF00";
Color[5] = "#0000FF";
Color[6] = "#FFFF00";
Color[7] = "#FF00FF";
Color[8] = "#00FFFF";
Color[9] = "#800000";
Color[10] = "#008000";
Color[11] = "#000080";
Color[12] = "#808000";
Color[13] = "#800080";
Color[14] = "#008080";
Color[15] = "#C0C0C0";
Color[16] = "#808080";
Color[17] = "#9999FF";
Color[18] = "#993366";
Color[19] = "#FFFFCC";
Color[20] = "#CCFFFF";
Color[21] = "#660066";
Color[22] = "#FF8080";
Color[23] = "#0066CC";
Color[24] = "#CCCCFF";
Color[25] = "#000080";
Color[26] = "#FF00FF";
Color[27] = "#FFFF00";
Color[28] = "#00FFFF";
Color[29] = "#800080";
Color[30] = "#800000";
Color[31] = "#008080";
Color[32] = "#0000FF";
Color[33] = "#00CCFF";
Color[34] = "#CCFFFF";
Color[35] = "#CCFFCC";
Color[36] = "#FFFF99";
Color[37] = "#99CCFF";
Color[38] = "#FF99CC";
Color[39] = "#CC99FF";
Color[40] = "#FFCC99";
Color[41] = "#3366FF";
Color[42] = "#33CCCC";
Color[43] = "#99CC00";
Color[44] = "#FFCC00";
Color[45] = "#FF9900";
Color[46] = "#FF6600";
Color[47] = "#666699";
Color[48] = "#969696";
Color[49] = "#003366";
Color[50] = "#339966";
Color[51] = "#003300";
Color[52] = "#333300";
Color[53] = "#993300";
Color[54] = "#993366";
Color[55] = "#333399";
Color[56] = "#333333";

char Context = ' ';
char SubContext = ' '; 

string fileName = "";
string Path = "";
string Entity = "";
string Header = "";
string Footer = ""; 

			StreamReader SR;
    		string S;
   		 	SR=File.OpenText("excel2xslt.config");
    		S=SR.ReadLine();
    		for(int i=1;i<=11;i++)
    		{
    		Console.WriteLine(S);
    		S=SR.ReadLine();
    		if (S.Length != 0 && S.Substring(0,1) != "[" && S.Substring(0,5) == "Files") 
    		{
	    		Path = XsltFunctions.SubstringBefore(XsltFunctions.SubstringAfter(S,"\""),"\"");
    		}
    		if (S.Length != 0 && S.Substring(0,1) != "[" && S.Substring(0,5) == "Excel") 
    		{
	    		fileName = XsltFunctions.SubstringBefore(XsltFunctions.SubstringAfter(S,"\""),"\"");
    		}
    		if (S.Length != 0 && S.Substring(0,1) != "[" && S.Substring(0,6) == "Header") 
    		{
	    		Header = XsltFunctions.SubstringBefore(XsltFunctions.SubstringAfter(S,"\""),"\"");
    		}
    		 if (S.Length != 0 && S.Substring(0,1) != "[" && S.Substring(0,6) == "Footer") 
    		{
	    		Footer = XsltFunctions.SubstringBefore(XsltFunctions.SubstringAfter(S,"\""),"\"");
    		}
    		if (S.Length != 0 && S.Substring(0,1) != "[" && S.Substring(0,7) == "Context") 
    		{
	    		Context = Convert.ToChar(XsltFunctions.SubstringBefore(XsltFunctions.SubstringAfter(S,"\""),"\""));
    		}
    		if (S.Length != 0 && S.Substring(0,1) != "[" && S.Substring(0,3) == "Sub") 
    		{
	    		SubContext = Convert.ToChar(XsltFunctions.SubstringBefore(XsltFunctions.SubstringAfter(S,"\""),"\""));
    		}
    		if (S.Length != 0 && S.Substring(0,1) != "[" && S.Substring(0,6) == "Entity") 
    		{
	    		Entity = XsltFunctions.SubstringBefore(XsltFunctions.SubstringAfter(S,"\""),"\"");
    		}
			}
    		SR.Close();
    		
    		if (Context == ' ') { Console.WriteLine("Context not defined. Please update the value in excel2xslt.config"); 		};
			if (SubContext == ' ') {Console.WriteLine("SubContext not defined. Please update the value in excel2xslt.config"); }; 
			if (fileName == "") {Console.WriteLine("FileName not defined. Please update the value in excel2xslt.config"); } ;
			if (Path == "") {Console.WriteLine("Path not defined. Please update the value in excel2xslt.config"); };
			if (Entity == "") {Console.WriteLine("Entity not defined. Please update the value in excel2xslt.config"); };
			if (Header == "") {Console.WriteLine("Header not defined. Please update the value in excel2xslt.config"); };
			if (Footer == "") {Console.WriteLine("Footer not defined. Please update the value in excel2xslt.config"); }; 
					
			
			Excel.Application excelApp = new Excel.ApplicationClass();  // Creates a new Excel Application
			excelApp.Visible = false;  // visible
			string workbookPath = Path + fileName; 
			Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0,
				false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, 
				false,  0, true, false, false);
			
			Console.WriteLine();
			Console.WriteLine("Output files written to: "+Path);
				
			StreamWriter swXMLCustomCfg = new StreamWriter(Path+"MailCustomCfg_"+XsltFunctions.SubstringBefore(fileName,".")+".xml", false, System.Text.Encoding.GetEncoding("ISO-8859-1"));				
			StreamWriter swXMLHeaderCfg = new StreamWriter(Path+"MailHeaderCfg_"+XsltFunctions.SubstringBefore(fileName,".")+".xml", false, System.Text.Encoding.GetEncoding("ISO-8859-1"));				
			StreamWriter swXMLConstants = new StreamWriter(Path+"Constants_"+XsltFunctions.SubstringBefore(fileName,".")+".xml", false, System.Text.Encoding.GetEncoding("ISO-8859-1"));				
			
			
			swXMLCustomCfg.WriteLine("<?xml version='1.0' encoding='ISO-8859-1' standalone='yes'?>");
			swXMLCustomCfg.WriteLine("<ResultSet xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>");
			
			swXMLHeaderCfg.WriteLine("<?xml version='1.0' encoding='ISO-8859-1'?>");
			swXMLHeaderCfg.WriteLine("<ResultSet>");
			
			swXMLConstants.WriteLine("<?xml version='1.0' encoding='ISO-8859-1'?>");
			swXMLConstants.WriteLine("<root>");
				
			// The following gets the Worksheets collection
			Excel.Sheets excelSheets = excelWorkbook.Worksheets;
			
			//Console.WriteLine("No of sheets"+excelSheets.Count);
		
			
			for(int s=1;s<=excelSheets.Count;s++)
			{
	
			Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(s);
			
			StreamWriter sw = new StreamWriter(Path+excelWorksheet.Name+".xsl", false, System.Text.Encoding.GetEncoding("ISO-8859-1"));

swXMLConstants.WriteLine("    <Constant>");
swXMLConstants.WriteLine("        <DictDetail_Name Type='String'>"+ excelWorksheet.Name +"</DictDetail_Name>");
swXMLConstants.WriteLine("        <DictDetail_ShortName Type='String'>"+ excelWorksheet.Name +"</DictDetail_ShortName>");
swXMLConstants.WriteLine("        <ElementOwner Type='Character'>C</ElementOwner>");
swXMLConstants.WriteLine("        <DictCoherency_Name Type='String'>XSLTFileName</DictCoherency_Name>");
swXMLConstants.WriteLine("        <ReturnType Type='Character'>S</ReturnType>");
swXMLConstants.WriteLine("        <Description Type='String'/>");
swXMLConstants.WriteLine("        <Sorter Type='String'>XSLT</Sorter>");
swXMLConstants.WriteLine("        <StringValue Type='EmptyString'>"+ excelWorksheet.Name +"</StringValue>");
swXMLConstants.WriteLine("        <Context Type='Character'>" + Context + "</Context>");
swXMLConstants.WriteLine("    </Constant>");			

					
swXMLHeaderCfg.WriteLine("	  <MailHeaderCfg>");
swXMLHeaderCfg.WriteLine("        <MailHeaderCfg_Name type='KDB_FIELD_NAME_0032'>"+excelWorksheet.Name+"</MailHeaderCfg_Name>");
swXMLHeaderCfg.WriteLine("        <Description type='string'/>");
swXMLHeaderCfg.WriteLine("        <CFG_Receiver type='string'>EDD_ThirdParty_Id_Cpty</CFG_Receiver>");
swXMLHeaderCfg.WriteLine("        <MailType_Name type='string'>Pdf Confirmation</MailType_Name>");
swXMLHeaderCfg.WriteLine("        <CFG_Master type='string'>EmptyEvent</CFG_Master>");
swXMLHeaderCfg.WriteLine("        <CFG_FaxAddress type='string'>DummyElement</CFG_FaxAddress>");
swXMLHeaderCfg.WriteLine("        <CFG_EmailFrom type='string'>DummyElement</CFG_EmailFrom>");
swXMLHeaderCfg.WriteLine("        <CFG_EmailTo type='string'>DummyElement</CFG_EmailTo>");
swXMLHeaderCfg.WriteLine("        <CFG_EmailCC type='string'>DummyElement</CFG_EmailCC>");
swXMLHeaderCfg.WriteLine("        <CFG_Bic type='string'>EVB_Bic_Id</CFG_Bic>");
swXMLHeaderCfg.WriteLine("        <ToBePrinted type='YesNo_t'>Y</ToBePrinted>");
swXMLHeaderCfg.WriteLine("        <NumberOfCopies type='KDB_FIELD_INTEGER'>1</NumberOfCopies>");
swXMLHeaderCfg.WriteLine("        <CFG_ExternalRef type='string'>NONE</CFG_ExternalRef>");
swXMLHeaderCfg.WriteLine("        <CFG_AddAddress type='string'>DummyElement</CFG_AddAddress>");
swXMLHeaderCfg.WriteLine("        <CFG_TypeOfOper type='string'>TAG22A</CFG_TypeOfOper>");
swXMLHeaderCfg.WriteLine("        <CFG_ReleaseDate type='string'>DummyElement</CFG_ReleaseDate>");
swXMLHeaderCfg.WriteLine("        <CFG_DeadLine type='string'>DummyElement</CFG_DeadLine>");
swXMLHeaderCfg.WriteLine("        <ValidityDateBegin type='KDB_FIELD_DATE'>2007/03/09</ValidityDateBegin>");
swXMLHeaderCfg.WriteLine("        <ValidityDateEnd type='KDB_FIELD_DATE'>2100/03/09</ValidityDateEnd>");
swXMLHeaderCfg.WriteLine("        <Priority type='BOPriority_t'>N</Priority>");
swXMLHeaderCfg.WriteLine("        <CFG_Entity type='string'>EBO_Entity_Id</CFG_Entity>");
swXMLHeaderCfg.WriteLine("        <AutoSending type='YesNo_t'>N</AutoSending>");
swXMLHeaderCfg.WriteLine("        <CFG_XSLT type='string'>"+excelWorksheet.Name+"</CFG_XSLT>");
swXMLHeaderCfg.WriteLine("        <Context type='BOContext_t'>" + Context + "</Context>");
swXMLHeaderCfg.WriteLine("        <SubContext type='BOContext_t'>" + SubContext + "</SubContext>");
swXMLHeaderCfg.WriteLine("        <CFG_Description type='string'>ModelDescripMails</CFG_Description>");
swXMLHeaderCfg.WriteLine("        <Entity_Name type='string'>"+Entity+"</Entity_Name>");
swXMLHeaderCfg.WriteLine("        <Validity type='BOValidity_t'>A</Validity>");
swXMLHeaderCfg.WriteLine("    </MailHeaderCfg>");
				
			Excel.Range cells = (Excel.Range)excelWorksheet.get_Range("A1", excelWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing));
			

			int[] actual_columns = new int[cells.Rows.Count+1];
			
			for(int i=1;i<= cells.Rows.Count;i++)
			{
				actual_columns[i] = 0; 
				for(int j = 1; j<= cells.Columns.Count;j++)
				{
					Excel.Range current = (Excel.Range)cells[i, j];
					if(current.Value2 != null)
					{
						actual_columns[i]=j;
					};
				};
				//Console.WriteLine(actual_columns[i]);
			};
			
			
			
sw.WriteLine("<?xml version='1.0' encoding='ISO-8859-1'?>");
sw.WriteLine("<xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform' xmlns:fo='http://www.w3.org/1999/XSL/Format' xmlns:xs='http://www.w3.org/2001/XMLSchema' xmlns:java='http://xml.apache.org/xslt/java' exclude-result-prefixes='java' xmlns:date='http://exslt.org/dates-and-times'");
sw.WriteLine(">");
sw.WriteLine(" <xsl:include href='"+Header+"'/>");
sw.WriteLine(" <xsl:include href='"+Footer+"'/> ");
sw.WriteLine("");
sw.WriteLine(" <xsl:template match='/'>");
sw.WriteLine("<fo:root xmlns:fo='http://www.w3.org/1999/XSL/Format'>");
sw.WriteLine("        <fo:layout-master-set>");
sw.WriteLine("                <fo:simple-page-master master-name='single_page' page-height='297mm' page-width='210mm' >");
sw.WriteLine("                        <fo:region-body region-name='xsl-region-body' ");
sw.WriteLine("                                margin='20mm' />");
sw.WriteLine("                </fo:simple-page-master>");
sw.WriteLine("                <fo:page-sequence-master master-name='repeatable_master'>");
sw.WriteLine("                        <fo:repeatable-page-master-reference master-reference='single_page' maximum-repeats='10'/>");
sw.WriteLine("                </fo:page-sequence-master>");
sw.WriteLine("        </fo:layout-master-set>");
sw.WriteLine("");
sw.WriteLine("        <fo:page-sequence master-reference='repeatable_master'>");
sw.WriteLine("                <fo:flow flow-name='xsl-region-body'>");
sw.WriteLine("                ");
sw.WriteLine("        <fo:table font-family='Times New Roman' font-size='10.00pt' width='180.00mm'>");
sw.WriteLine("                                <fo:table-column column-width='180mm'/>");
sw.WriteLine("                                                        ");
sw.WriteLine("                                <fo:table-body>");
sw.WriteLine("                                        <fo:table-row height='55mm'>");
sw.WriteLine("                                                <fo:table-cell   padding-before='3pt' padding-after='3pt'   padding-start='3pt' padding-end='3pt'>");
sw.WriteLine("                                                <xsl:call-template name='HEADER'/> ");
sw.WriteLine("                                                <fo:block height='20mm' text-align='left'></fo:block>");
sw.WriteLine("                                                </fo:table-cell>");
sw.WriteLine("                                        </fo:table-row>");
sw.WriteLine("                                </fo:table-body>");
sw.WriteLine("                </fo:table>");

/* second table - body */
sw.WriteLine("	<fo:table font-family='Times New Roman' font-size='10.00pt' width='180.00mm'>"); 

			
			for(int i=1;i<= cells.Columns.Count;i++)
			{
				sw.WriteLine("<fo:table-column column-width='"+180/cells.Columns.Count+"mm'/>");
			}

			sw.WriteLine("								<fo:table-body>");
			
			int[] total_columns = new int[cells.Rows.Count+1];
			string[] choose = new string[cells.Rows.Count+1];
			
			for(int i=1;i<=cells.Rows.Count;i++)
				{
					total_columns[i] = 0;
					Excel.Range current = (Excel.Range)cells[i, 1];
					if (current.Value2 != null)
					{
						if((string)XsltFunctions.Substring(current.Value2.ToString(),1,6) == "choose")
						{
							choose[i] = "<xsl:choose>";
						}
						else
						{
							if((string)XsltFunctions.Substring(current.Value2.ToString(),1,4) == "when")
							{
								choose[i] = "<xsl:when test=\""+(string)XsltFunctions.SubstringAfter(current.Value2.ToString(),"when ")+"\">";
							/*	if (XsltFunctions.Substring((string)XsltFunctions.SubstringAfter(current.Value2.ToString(),"when "),1,8) == "MailBody")
										{	
										swXMLCustomCfg.WriteLine("	<MailCustomCfg>");
										swXMLCustomCfg.WriteLine("<MailHeaderCfg_Name type='string'>"+excelWorksheet.Name+"</MailHeaderCfg_Name>");
										swXMLCustomCfg.WriteLine("		<MailCustomTag type='string'>"+XsltFunctions.SubstringAfter(XsltFunctions.SubstringBefore((string)XsltFunctions.SubstringAfter(current.Value2.ToString(),"when ")," and "),"MailBody/")+"</MailCustomTag>");
										swXMLCustomCfg.WriteLine("		<Comment type='string'>/</Comment>");
										swXMLCustomCfg.WriteLine("		<CFG_Element type='string'>Empty</CFG_Element>");
										swXMLCustomCfg.WriteLine("		<SortingNum type='KDB_FIELD_INTEGER'></SortingNum>");
										swXMLCustomCfg.WriteLine("	</MailCustomCfg>");
										};										
								if (XsltFunctions.Substring((string)XsltFunctions.SubstringAfter(current.Value2.ToString()," and "),1,8) == "MailBody")
										{	
										swXMLCustomCfg.WriteLine("	<MailCustomCfg>");
										swXMLCustomCfg.WriteLine("<MailHeaderCfg_Name type='string'>"+excelWorksheet.Name+"</MailHeaderCfg_Name>");
										swXMLCustomCfg.WriteLine("		<MailCustomTag type='string'>"+XsltFunctions.SubstringAfter((string)XsltFunctions.SubstringAfter(current.Value2.ToString()," and "),"MailBody/")+"</MailCustomTag>");
										swXMLCustomCfg.WriteLine("		<Comment type='string'>/</Comment>");
										swXMLCustomCfg.WriteLine("		<CFG_Element type='string'>Empty</CFG_Element>");
										swXMLCustomCfg.WriteLine("		<SortingNum type='KDB_FIELD_INTEGER'></SortingNum>");
										swXMLCustomCfg.WriteLine("	</MailCustomCfg>");
										}
							*/
							}
						
							else 
							{
								if((string)XsltFunctions.Substring(current.Value2.ToString(),1,7) == "endwhen")
								{
									choose[i] = "</xsl:when>";
								}
								else
								{
									if((string)XsltFunctions.Substring(current.Value2.ToString(),1,9) == "endchoose")
										{
											choose[i] = "</xsl:choose>";
										}
										
									else
									{
										if((string)XsltFunctions.Substring(current.Value2.ToString(),1,9) == "otherwise")
										{
											choose[i] = "<xsl:otherwise>";
										}
											
										else
										{
										 if((string)XsltFunctions.Substring(current.Value2.ToString(),1,12) == "endotherwise")
											{
											choose[i] = "</xsl:otherwise>";
											}	
										else
									{
										choose[i]="";
									};
								};
							};
						}
					};	
		};
	};
};
			
			
				
			for (int i = 1; i <= cells.Rows.Count; i++)
			{
				int span = 1;
				int mergedcellno = 0;
				int t=0;
				
			if (choose[i] != "")
			{
				sw.WriteLine(choose[i]);
			}
		else
		{
				
				sw.WriteLine("<fo:table-row height='5mm'>");
			for(int j = 1; j<= cells.Columns.Count && total_columns[i]<cells.Columns.Count;j++)
					{ 
				if (mergedcellno != 0 && t<mergedcellno && t!= 0)
					{
						t = t+ 1;
					}
				else
				{	
				if (t>=mergedcellno)
					{
						t=0;
						mergedcellno = 0;
					};	
				};	
					
				 if (t == 0)
				 {
					String color;
					Excel.Range current = (Excel.Range)cells[i, j];
					//Console.WriteLine("Color "+current.Interior.ColorIndex);
					if ((int)current.Interior.ColorIndex>=1 && (int)current.Interior.ColorIndex<=56)
					{
					color = Color[(int)current.Interior.ColorIndex];
					}
					else
					{
						color="white";
					};
				
		
					
					if (current.Value2 != null)
					{			
						if((bool)current.MergeCells)
						{
							//Console.WriteLine(current.get_Address("", "", Excel.XlReferenceStyle.xlR1C1,"", "")+"Row "+i+" Column "+j+" merged"+":"+current.MergeArea.get_Address("", "", Excel.XlReferenceStyle.xlR1C1,"", ""));
							span = Convert.ToInt32(XsltFunctions.SubstringAfter(XsltFunctions.SubstringAfter(current.MergeArea.get_Address("", "", Excel.XlReferenceStyle.xlR1C1,"", ""),":"),"C")) - Convert.ToInt32(XsltFunctions.SubstringBefore(XsltFunctions.SubstringAfter(current.MergeArea.get_Address("", "", Excel.XlReferenceStyle.xlR1C1,"", ""),"C"),":")) + 1;
							mergedcellno = span;
						}
						else
						{
							span = 1;
						};
											
					sw.WriteLine("							<fo:table-cell number-columns-spanned='"+ span +"' background-color='"+ color +"' border-color='black' border-width='0.2pt' padding-before='3pt' padding-after='3pt'   padding-start='3pt' padding-end='3pt'>");
					sw.WriteLine("									<fo:block height='5mm' text-align='left'>");
																	
	
							
					if ((bool)XsltFunctions.Contains(current.Value2.ToString(),"MailBody") == true)
					{	
						if ((bool)(XsltFunctions.Contains(current.Value2.ToString(),"Amount")) == true || (bool)(XsltFunctions.Contains(current.Value2.ToString(),"Rate")) == true)
						{						
						sw.WriteLine("												<xsl:call-template name='Amount'>");
						sw.WriteLine("													<xsl:with-param name='Amount' select='"+current.Value2.ToString()+"'/>");
						sw.WriteLine("												</xsl:call-template>");
						}
						else
						{
							if ((bool)(XsltFunctions.Contains(current.Value2.ToString(),"Date")) == true)
								{						
						sw.WriteLine("												<xsl:call-template name='Date'>");
						sw.WriteLine("													<xsl:with-param name='Date' select='"+current.Value2.ToString()+"'/>");
						sw.WriteLine("												</xsl:call-template>");
								}
								else
								{
									sw.WriteLine("									"+"<xsl:value-of select = '"+current.Value2.ToString()+"'/>");
								}
						};
					swXMLCustomCfg.WriteLine("	<MailCustomCfg>");
					swXMLCustomCfg.WriteLine("<MailHeaderCfg_Name type='string'>"+excelWorksheet.Name+"</MailHeaderCfg_Name>");
					swXMLCustomCfg.WriteLine("		<MailCustomTag type='string'>"+XsltFunctions.SubstringAfter(current.Value2.ToString(),"MailBody/")+"</MailCustomTag>");
					swXMLCustomCfg.WriteLine("		<Comment type='string'>/</Comment>");
					swXMLCustomCfg.WriteLine("		<CFG_Element type='string'>Empty</CFG_Element>");
					swXMLCustomCfg.WriteLine("		<SortingNum type='KDB_FIELD_INTEGER'></SortingNum>");
					swXMLCustomCfg.WriteLine("	</MailCustomCfg>");
					}
					else
					{
						sw.WriteLine("									"+current.Value2.ToString());
					};
				
					//Console.WriteLine(current.Value2);
					sw.WriteLine("									</fo:block>");
					sw.WriteLine("							</fo:table-cell>");
										}
					else
					{
					sw.WriteLine("							<fo:table-cell number-columns-spanned='"+ span +"' background-color='"+ color + "' border-color='black' border-width='0.2pt' padding-before='3pt' padding-after='3pt'   padding-start='3pt' padding-end='3pt'>");
					sw.WriteLine("									<fo:block height='5mm' text-align='left'>");
					sw.WriteLine("									</fo:block>");
					sw.WriteLine("							</fo:table-cell>");
					};
									
					total_columns[i] = total_columns[i] + span;
					span = 1;		
				}
				
	
					
					if (mergedcellno != 0 && t<mergedcellno)
					{
						t= t+ 1;
					};
						}	
				sw.WriteLine("						     		</fo:table-row>");
				sw.WriteLine();
		};
			}

sw.WriteLine("								</fo:table-body>");
sw.WriteLine("								</fo:table>");			
						
		
sw.WriteLine("                                                               /* Placeholder for footer */");
sw.WriteLine("                                                                                                ");
sw.WriteLine("                                                        <fo:table font-family='Times New Roman' font-size='10.00pt' width='120.00mm'>");
sw.WriteLine("                                                                        <fo:table-column column-width='120mm'/>");
sw.WriteLine("                                                        ");
sw.WriteLine("                                                                <fo:table-body start-indent='0pt'>");
sw.WriteLine("                                                                ");
sw.WriteLine("                                                                <fo:table-row height='100mm'>");
sw.WriteLine("                                                                        <fo:table-cell   padding-before='3pt' padding-after='3pt'   padding-start='3pt' padding-end='3pt'>");
sw.WriteLine("                                                                        <xsl:call-template name='FOOTER'/>");
sw.WriteLine("                                                                                        <fo:block height='20mm' text-align='left'>  </fo:block>");
sw.WriteLine("                                                                        </fo:table-cell>");
sw.WriteLine("                                                                </fo:table-row>");
sw.WriteLine("                                                                </fo:table-body>");
sw.WriteLine("                                                                </fo:table>                     ");
sw.WriteLine("                </fo:flow>");
sw.WriteLine("        </fo:page-sequence>");
sw.WriteLine("</fo:root>");
sw.WriteLine("");
sw.WriteLine("</xsl:template>");
sw.WriteLine("</xsl:stylesheet>");
	

		
			
			sw.Close();
			file_counter++;
			Console.WriteLine(file_counter+". XSL file successfully written to: " +excelWorksheet.Name+".xsl");
		}
		swXMLCustomCfg.WriteLine("</ResultSet>");
		swXMLCustomCfg.Close();
		file_counter++;
		Console.WriteLine(file_counter+". MailCustomCfg file successfully written to: "+"MailCustomCfg_"+XsltFunctions.SubstringBefore(fileName,".")+".xml");
		
		swXMLHeaderCfg.WriteLine("</ResultSet>");
		swXMLHeaderCfg.Close();
		file_counter++;
		Console.WriteLine(file_counter+". MailHeaderCfg file successfully written to: "+"MailHeaderCfg_"+XsltFunctions.SubstringBefore(fileName,".")+".xml");
		
		swXMLConstants.WriteLine("</root>");
		swXMLConstants.Close();
		file_counter++;
		Console.WriteLine(file_counter+". Constants file successfully written to: "+"Constants_"+XsltFunctions.SubstringBefore(fileName,".")+".xml");
		excelWorkbook.Close(false, 0,0 );
		}
	}
}
