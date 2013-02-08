/**************************************************************************\
PROGRAM INFORMATION
Author  : William Roehl - MarketShare 
Project : SAS Global Forum 2013 Submission
Purpose : Provide a relatively easy interface for importing and exporting Google Spreadsheets into/out of SAS
Inputs  : API->datasets
Outputs : dataset->API
Requires: HTML Tidy for Win32: http://www.paehl.com/open_source/?HTML_Tidy_for_Windows
		      Autocall macros: %array, %do_over

PROGRAM HISTORY
2012-03-13 WR Initial program developed.
2013-01-10 WR Updates to CreateWS to properly handle column headers.
\**************************************************************************/;

/* Set file paths */
%global DRIVE TEMP TINY; %let DRIVE = C; %let TEMP = TEMP; %let TIDY = SAS\TIDY.EXE;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %SAS_SETHEADER
Desc 	    : Pass HTTP headers to PROC HTTP
Parameters: HEADER, HEADER2= through HEADER5=
			      FILE_EXTENSION=TXT (if you're uploading a file other than TXT, change this (e.g. XLS)
\**************************************************************************/;
%macro sas_setHeader(header, header2=NO, header3=NO, header4=NO, header5=NO, file_extension=txt);
	/* Set the header files and output file */
	filename hin  "&DRIVE.:\&TEMP.\header_in.txt";
	filename hout "&DRIVE.:\&TEMP.\header_out.txt";
	filename out "&DRIVE.:\&TEMP.\out.&FILE_EXTENSION.";

	/* Create the header file to pass */
	data _null_;
		file hin lrecl=10000;
		put "&HEADER.";
		%do_over(values=&HEADER2.;&HEADER3.;&HEADER4.;&HEADER5., phrase=if "?" ^= "NO" then do; FIXED = prxchange('s/!/-/',-1,"?"); put FIXED; end;, DELIM=;);
	run;
%mend;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_AUTH
Desc 	    : Authenticate and grab the AUTH/AUTH_S variables for later use (globally)
Parameters: U - Google Account username (username@gmail.com)
			      P - Google Account password
\**************************************************************************/;
%macro gdoc_auth (u,p);
	/* Setup files for PROC HTTP, %arrays(), and global macro AUTH variables */
	%global auth auth_s;
	filename in "&DRIVE.:\&TEMP.\in.txt"; filename out "&DRIVE.:\&TEMP.\out.txt"; filename hout "&DRIVE.:\&TEMP.\header_out.txt"; filename hin "&DRIVE.:\&TEMP.\header_in.txt";
	data _null_; file in; file out; file hout; file hin; run;
	%array(TYPE, values=writely wise); %array(AUTH, values=AUTH AUTH_S);

	/* Do each type of authentication: DOCS and SPREADSHEETS then create AUTH and AUTH_S macro vars */
	%do_over(TYPE AUTH, phrase=
		proc http in=in out=out url="https://www.google.com/accounts/ClientLogin?Email=&u.&Passwd=&p.&service=?TYPE" method="POST"; run;
		data _null_;
			infile out lrecl=10000 dlm='=';
			format var $8. value $10000.;
			input var value;
			if var = 'Auth'; 
			put value=;
			call symputx('?AUTH',trim(value));
		run;);
%mend;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_GETLIST
Desc 	    : Get a list of Google Drive 'files'
Parameters: N/A
\**************************************************************************/;
%macro gdoc_getList();
	/* Set header, do call, make the XML SAS readable */	
	%sas_setHeader(Authorization: GoogleLogin Auth=&AUTH.);
	proc http headerin=hin headerout=hout out=out url="https://docs.google.com/feeds/default/private/full?v=3" method="GET"; run;
	x "&DRIVE.:\&TIDY. -xml -indent -wrap 256 -quiet -m &DRIVE.:\&TEMP.\out.txt";

	/* Read API's XML response into a dataset */
	data gdoc_list;
		format line $256.;
		infile out DLM='0A'x;
		input line;
	run;

	/* Grab URLs for the documents and strip extraneous tags */
	data gdoc_list_new(where=(id ^= ''));
		set gdoc_list;
		format id $100.;
			
		/* New list entry */
		if prxmatch('/<\/entry>/',line) then rank + 1;

		/* Edit link */
		if prxmatch ('/resumable-create-media/',line) then do;
			line = prxchange("s/^.*href=.?//",-1,line);
			line = prxchange("s/'.\/>.*$//",-1,line);
			id = 'create_link';
		end;		

		/* Export link */
		if prxmatch ('/Export/',line) then do;
			line = prxchange("s/^.*src=.?//",-1,line);
			line = prxchange("s/'.\/>.*$//",-1,line);
			id = 'export_link';
		end;		

		/* Get the ETag for editing */
		if prxmatch('/gd:etag/',line) then do;
			line = prxchange('s/^.*etag=.//',-1,line);
			line = prxchange('s/".>.*$/"/',-1,line);
			id = 'e_tag';
		end;		

		/* ID */
		if prxmatch ('/<id>/',line) then do;
			line = prxchange('s/^.*<id>//',-1,line);
			line = prxchange('s/<\/id>.*$//',-1,line);
			id = 'sheet_id';
		end;

		/* Name of document */
		if prxmatch ('/<\/title>/',line) then do;
			line = prxchange('s/^.*<title>//',-1,line);
			line = prxchange('s/<\/title>.*$//',-1,line);
			id = 'sheet_name';
		end;

		/* Sheet URL */
		if prxmatch ('/link rel=.edit-media..type/',line) then do;
			line = prxchange("s/^.*href=.?//",-1,line);
			line = prxchange("s/'.\/>.*$//",-1,line);
			id = 'sheet_url';
		end;

		retain rank;
	run;

	/* Put each Docs listing on a single line */
	proc transpose data=gdoc_list_new let
				   out=gdoc_list_xpose(drop=_NAME_ rank where=(e_tag ne '')); 
				   var line;
				   id id;
				   by rank;
	run;

	/* Force the create link to propagate -- yes this is ugly but I'm lazy today. */
	data gdoc_DocList(drop=l_create_link);
		set gdoc_list_xpose;

		if create_link ^= '' then l_create_link = create_link;
		if create_link = ''  then create_link = l_create_link;

		retain l_create_link;
	run;

	/**************************************************************************\
		This section is for Spreadsheet worksheets
	\**************************************************************************/;

	/* Set header, do call, make the XML SAS readable */	
	%sas_setHeader(Authorization: GoogleLogin Auth=&AUTH_S.);
	proc http headerin=hin headerout=hout out=out url="https://spreadsheets.google.com/feeds/spreadsheets/private/full?v=3" method="GET"; run;
	x "&DRIVE.:\&TIDY. -xml -indent -wrap 256 -quiet -m &DRIVE.:\&TEMP.\out.txt";

	/* Read API's XML response into a dataset */;
	data gdoc_list;
		format line $256.;
		infile out lrecl=256 DLM='0A'x;
		input line;
	run;

	data gdoc_list_new(where=(id ^= ''));
		set gdoc_list;
		format id $100.;

		/* New list entry */
		if prxmatch('/<\/entry>/',line) then rank + 1;
			
		/* New list entry */
		if prxmatch('/worksheets/',line) then do;
			line = prxchange("s/.*src=.//", -1, line);
			line = prxchange("s/' \/>//", -1, line);
			id = 'worksheet_id';
		end;

		/* Name of document */
		if prxmatch ('/<\/title>/',line) then do;
			line = prxchange('s/^.*<title>//',-1,line);
			line = prxchange('s/<\/title>.*$//',-1,line);
			id = 'sheet_name';
		end;

		retain rank;
	run;

	/* Put each spreadsheet listing on a single line */
	proc transpose data=gdoc_list_new let
				   out=gdoc_wslist(drop=_NAME_ rank); 
				   var line;
				   id id;
				   by rank;
	run;

	/* Remove temporary datasets */
	proc datasets library=work nodetails nolist;
		delete gdoc_list gdoc_list_new gdoc_list_xpose;
	run;
	quit;
%mend gdoc_getList;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_GETDATA
Desc 	    : Download entire Spreadsheet. 
Notes	    : Will create one dataset per tab in Spreadsheet named from a slightly altered tab name 
Parameters: URL - Google Docs URL of item to download
\**************************************************************************/;
%macro gdoc_getData(url);
	/* Set header, do call */	
	%sas_setHeader(Authorization: GoogleLogin Auth=&AUTH_S., file_extension=xls);
	proc http headerin=hin headerout=hout out=out url="&URL.&exportFormat=XLS" method="GET"; run;

	/* Import the XLS file into a SAS library */
	LIBNAME GDOCXLS EXCEL "&DRIVE.:\&TEMP.\OUT.XLS"
		DBSASLABEL=COMPAT
		GETNAMES=Yes
		HEADER=Yes
		SCANTEXT=Yes
		SCANTIME=Yes
		USEDATE=Yes; 

	/* Pull sheet names in from XLS and put in appropriate macro variables (no special chars - dataset name) */
	proc sql NOPRINT; 
		SELECT DISTINCT X INTO :XLSNOSPACES separated by '|'
		FROM (SELECT propcase(prxchange('s/[^\w*]//', -1, trim(memname))) as X, memname FROM dictionary.tables WHERE libname = 'GDOCXLS' and memname contains '$') 
		ORDER BY memname;
	quit;

	/* Pull sheet names in from XLS and put in appropriate macro variables (drop the ' only, used for the range/sheet) */
	proc sql NOPRINT; 
		SELECT DISTINCT X INTO :XLSSPACES separated by '|'
		FROM (SELECT trim(prxchange("s/'//", -1, memname)) as X, memname FROM dictionary.tables WHERE libname = 'GDOCXLS' and memname contains '$') 
		ORDER BY memname;
	quit;

	/* We're done with the Excel libname, clear the reference */
	LIBNAME GDOCXLS CLEAR;

	/* Import the XLS into datasets in the work library */
	%array(DELIMSHEETS, values=&XLSNOSPACES.,delim = '|')
	%array(SHEETS, values=&XLSSPACES.,delim = '|')
	%do_over(DELIMSHEETS SHEETS, DELIM=|, phrase=proc import datafile="&DRIVE.:\&TEMP.\OUT.XLS" 
															 out=?DELIMSHEETS
															 dbms=excel 
															 replace;
															 range="?SHEETS"; 
												 run;);

	options noxwait;
	x "del &DRIVE.:\&TEMP.\OUT.XLS";
%mend gdoc_getData;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_CREATEDOC
Desc 	    : Create a Google Doc (new)
Parameters: TITLE - Title of the new Google Document to create
			      DATASET - SAS dataset name to upload
\**************************************************************************/;
%macro gdoc_createDoc(title, dataset);

	/* Create empty XML payload with title */
	data _null_;
		file "&DRIVE.:\&TEMP.\payload.xml";

		put '<?xml version="1.0" encoding="UTF-8"?>';
		put '	<entry xmlns="http://www.w3.org/2005/Atom" xmlns:docs="http://schemas.google.com/docs/2007">';
  		put '	<category scheme="http://schemas.google.com/g/2005#kind"';
		put '		term="http://schemas.google.com/docs/2007#spreadsheet"/>';
		put "	<title>&TITLE.</title>";
		put '</entry>';
	run;

	/* Set header */	
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH. 
				 , header2=X-Upload-Content-Length: 0 
				 , header3=Slug: &TITLE.);

	/* Do API call */
	filename in "&DRIVE.:\&TEMP.\payload.xml";
	proc http headerin=hin headerout=hout in=in out=out ct="application/atom+xml" url="https://docs.google.com/feeds/default/private/full?v=3" method="POST"; run;
	x "&DRIVE.:\&TIDY. -xml -indent -wrap 256 -quiet -m &DRIVE.:\&TEMP.\out.txt";

	/* Read the response and get the URLs */
	filename hout;
	data gdoc_created;
		format line $256.;
		infile out lrecl=256 DLM='0A'x;
		input line;
	run;

	/* Grab URLs for the documents and strip extraenous tags */
	data _null_;
		set gdoc_created;

		/* Edit link */
		if prxmatch ('/revisions/',line) then do;
			line = prxchange("s/^.*href=.?//",-1,line);
			line = prxchange("s/'.\/>.*$//",-1,line);
			line = prxchange('s/\/revisions//',-1,line);
			call symputx('SPREADSHEET',urlencode(trim(line)));
		end;
	run;

	/* Export the dataset to TSV */
	proc export data=&DATASET.
		outfile="&DRIVE.:\&TEMP.\dataset.tsv"
		dbms=tab
		replace;
	run;

	/* Overwrite the newly created blank spreadsheet with data */

	/* Set header and make API call */	
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH. 
				 , header2=%str(If!Match: *));
	filename in "&DRIVE.:\&TEMP.\dataset.tsv";
	proc http headerin=hin headerout=hout in=in out=out ct="text/tab-separated-values" url="https://docs.google.com/feeds/default/media/&&SPREADSHEET.?v=3" method="PUT"; run;

	/* Remove temporary datasets */
	proc datasets library=work nodetails nolist;
		delete gdoc_created;
	run;
	quit;
%mend gdoc_createDoc;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_PUTDATA
Desc 	    : Overwrite preexisting Spreadsheet with updated data (or new data)
Parameters: URL - URL of Spreadsheet to overwrite
			      DATASET - SAS dataset name to upload to Google Docs
			      ETAG - This option intends to avoid people overwriting data when changes were made by others. (* to disable)
\**************************************************************************/;
%macro gdoc_putData(url, dataset, etag);
	/* Export the dataset to TSV */
	proc export data=&DATASET.
		outfile="&DRIVE.:\&TEMP.\dataset.tsv"
		dbms=tab
		replace;
	run;

	/* Set header and make API call */	
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH. 
				 , header2=%str(If!Match: &ETAG.));
	filename in "&DRIVE.:\&TEMP.\dataset.tsv";
	proc http headerin=hin headerout=hout in=in out=out ct="text/tab-separated-values" url="&URL." method="PUT"; run;
%mend gdoc_putData;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_ADDSHEET
Desc 	    : Create a new Google Spreadsheet Worksheet and Upload Data 
Parameters: URL - URL of spreadsheet
			      TITLE - Title of new worksheet
			      ROWS - Number of total rows that will be in the new worksheet 
			      COLS - Number of total columns that will be in the new worksheet
\**************************************************************************/;
%macro gdoc_addSheet(url, title, rows, cols);
	/* Create XML payload with title */
	data _null_;
		file "&DRIVE.:\&TEMP.\payload.xml";

		put '<entry xmlns="http://www.w3.org/2005/Atom"';
	    put '	xmlns:gs="http://schemas.google.com/spreadsheets/2006">';
  		put "	<title>&TITLE.</title>";
		put "	<gs:rowCount>&ROWS.</gs:rowCount>";
		put "	<gs:colCount>&COLS.</gs:colCount>";
		put '</entry>';
	run;

	/* Set header and make create call to API */
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH_S.);
	filename in "&DRIVE.:\&TEMP.\payload.xml";
	proc http headerin=hin headerout=hout in=in out=out ct="application/atom+xml" url="&URL." method="POST"; run;
	x "&DRIVE.:\&TIDY. -xml -indent -wrap 256 -quiet -m &DRIVE.:\&TEMP.\out.txt";

	/* Read the response file and get the URLs */
	data gdoc_worksheets;
		format line $256.;
		infile out lrecl=256 DLM='0A'x;
		input line;
	run;

	/* Grab URLs for the documents and strip extraenous tags */
	data gdoc_worksheets(where=(id ^= ''));
		set gdoc_worksheets;
		format id $100.;

		/* New list entry */
		if prxmatch('/etag/',line) then rank + 1;

		/* ID */
		if prxmatch ('/<id>/',line) then do;
			line = prxchange('s/^.*<id>//',-1,line);
			line = prxchange('s/<\/id>.*$//',-1,line);
			line = prxchange('s/worksheets/list/',-1,line);
			id = 'worksheet_id';
		end;

		/* Edit Link */
		if prxmatch ('/link rel=.edit/',line) then do;
			line = prxchange("s/^.*href=.?//",-1,line);
			line = prxchange("s/'.\/>.*$//",-1,line);
			id = 'edit_url';
		end;

		/* Name of document */
		if prxmatch ('/<\/title>/',line) then do;
			line = prxchange('s/^.*<title>//',-1,line);
			line = prxchange('s/<\/title>.*$//',-1,line);
			id = 'worksheet_name';
		end;
	run;

	/* Put each spreadsheet listing on a single line */
	proc transpose data=gdoc_worksheets
				   out=gdoc_sheetlst(drop=_NAME_ rank); 
				   var line;
				   id id;
				   by rank;
	run;

	/* Remove temporary datasets */
	proc datasets library=work nodetails nolist;
		delete gdoc_worksheets;
	run;
	quit;
%mend;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_UPDATECELL
Desc 	    : Update each cell (ROW/COLUMN) with DATA
Parameters: TITLE - Title of worksheet
			      ROW - Row number (X)
			      COLUMN - Column number (Y)
			      DATA - Contents of created cell
\**************************************************************************/;
%macro gdoc_updateCell(title, row, column, data);
	/* Get worksheet ID for new sheet */
	data _null_;
		set gdoc_sheetlst;
		where worksheet_name = "&TITLE.";
		call symputx('CELLS', prxchange('s/list/cells/',-1,trim(WORKSHEET_ID)));
	run;

	/* Set header and make create call to API to read cell data for row/column cell creation*/
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH_S.);
	proc http headerin=hin headerout=hout out=out ct="application/atom+xml" url="&&CELLS./private/full/R&ROW.C&COLUMN." method="GET"; run;
	x "&DRIVE.:\&TIDY. -xml -indent -wrap 256 -quiet -m &DRIVE.:\&TEMP.\out.txt";

	/* Read the response and get the URLs */
	data gdoc_cell;
		format line $256.;
		infile out lrecl=256 DLM='0A'x;
		input line;
	run;

	/* Grab URLs for the documents and strip extraneous tags */
	data gdoc_cell(where=(id ^= ''));
		set gdoc_cell;
		format id $100.;

		/* New list entry */
		if prxmatch('/etag/',line) then rank + 1;

		/* ID */
		if prxmatch ('/<id>/',line) then do;
			line = prxchange('s/^.*<id>//',-1,line);
			line = prxchange('s/<\/id>.*$//',-1,line);
			id = 'worksheet_id';
		end;

		/* Edit Link */
		if prxmatch ('/link rel=.edit/',line) then do;
			line = prxchange("s/^.*href=.?//",-1,line);
			line = prxchange("s/'.\/>.*$//",-1,line);
			id = 'edit_url';
		end;

		/* Name of document */
		if prxmatch ('/<\/title>/',line) then do;
			line = prxchange('s/^.*<title>//',-1,line);
			line = prxchange('s/<\/title>.*$//',-1,line);
			id = 'worksheet_name';
		end;
	run;

	/* Put each spreadsheet listing on a single line */
	proc transpose data=gdoc_cell
				   out=gdoc_cells(drop=_NAME_); 
				   var line;
				   id id;
				   by rank;
	run;

	/* Remove temporary datasets */
	proc datasets library=work nodetails nolist;
		delete gdoc_cell;
	run;
	quit;

	/* Get cell URL for new cell */
	data _null_;
		set gdoc_cells;
		call symputx('CELL_URL',(trim(EDIT_URL)));
	run;

	/* Create the header record (first line in spreadsheet) */
	data _null_;
		file "&DRIVE.:\&TEMP.\payload.xml";

		put '<entry xmlns="http://www.w3.org/2005/Atom"';
		put '		xmlns:gs="http://schemas.google.com/spreadsheets/2006">';
		put "	<id>&CELLS./private/full/R&ROW.C&COLUMN.</id>";
		put '	<link rel="edit" type="application/atom+xml"';
		put "		href=""&&CELL_URL.""/>";
		put "	<gs:cell row=""&ROW."" col=""&COLUMN."" inputValue=""&DATA""/>";
		put '</entry>';
	run;

	/* Set header and make create call to API to read cell data for row/column cell creation*/
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH_S.
               , header2=If!Match: *);
	filename in "&DRIVE.:\&TEMP.\payload.xml";
	proc http headerin=hin headerout=hout in=in out=out ct="application/atom+xml" url="&&CELLS./private/full/R&ROW.C&COLUMN." method="PUT"; run;
%mend;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_ADDROW
Desc 	    : Add a single row to a worksheet
Parameters: TITLE - Title of new worksheet
			      PAYLOAD - XML data to pass to the API for a single row.
\**************************************************************************/;
%macro gdoc_addRow(title, payload);
	/* Get worksheet ID for new sheet */
	data _null_;
		set gdoc_sheetlst;
		where worksheet_name = "&TITLE.";
		call symputx('URL',trim(WORKSHEET_ID));
	run;

	/* Set header and send off row creation payload */
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH_S.);
	filename in "&&PAYLOAD.";
	proc http headerin=hin headerout=hout in=in out=out ct="application/atom+xml" url="&URL./private/full" method="POST"; run;
%mend;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_CREATEWS
Desc 	    : Creates a new worksheet (tab) and uploads a dataset's contents into the new worksheet.
Parameters: SPREADSHEET_TITLE - Title of spreadsheet
			      NEW_SHEET_TITLE - Title of the newly created sheet (e.g. SHEET2)
			      DS - Dataset name
			      DS_LIB - Dataset libname (i.e. SASHELP).
\**************************************************************************/;
%macro gdoc_createWS(spreadsheet_title, new_sheet_title, ds, ds_lib);
	/* We need to remove special characters to do the upload and then replace them with the original column names later */
	data VCOLUMN; set SASHELP.VCOLUMN(where=(LIBNAME=UPCASE(resolve('&DS_LIB.')) AND MEMNAME=UPCASE(resolve('&DS.')))); NAME_LCASE = compress(lowcase(NAME),',<>?:;"{}[]|\~`!@#$%^&*()_-+=/*.'); run;
		%array(col_names, data=VCOLUMN, var=NAME_LCASE);
	data VCOLUMN(where=(UPCASE(NAME_LCASE) ne NAME)); set SASHELP.VCOLUMN(where=(LIBNAME=UPCASE(resolve('&DS_LIB.')) AND MEMNAME=UPCASE(resolve('&DS.')))); NAME_LCASE = compress(lowcase(NAME),',<>?:;"{}[]|\~`!@#$%^&*()_-+=/*.'); run;
		%array(rename_cols_to, data=VCOLUMN, var=NAME_LCASE);
		%array(rename_cols_from, data=VCOLUMN, var=NAME);
	data VCOLUMN; set SASHELP.VCOLUMN(where=(LIBNAME=UPCASE(resolve('&DS_LIB.')) AND MEMNAME=UPCASE(resolve('&DS.')))); run;
		%array(orig_col_names, data=VCOLUMN, var=NAME);
	proc datasets library=work nodetails nolist; delete VCOLUMN; run; quit;
	data MODIFIED_ORIG_DS; set &DS_LIB..&DS.; run;
	proc datasets library=WORK nodetails nolist; modify MODIFIED_ORIG_DS; %do_over(rename_cols_from rename_cols_to, phrase=rename ?rename_cols_from = ?rename_cols_to;); run; quit;

	/* Add blank sheet */	
	data _null_;
		set gdoc_wslist;
		where SHEET_NAME = "&SPREADSHEET_TITLE.";
		call symputx('SPREADSHEET',(trim(WORKSHEET_ID)));
	run;

	/* Open the dataset and determine the number of rows */
	%let DSID = %sysfunc(open(&DS_LIB..&DS., IS)); %let ROWS = %eval(%sysfunc(attrn(&DSID, NLOBS))+1); %let CLOSE_ID = %sysfunc(close(&DSID));

	/* Add blank sheet */
	%gdoc_addSheet(url=&SPREADSHEET.
			 	 , title=&NEW_SHEET_TITLE.
				 , rows=&ROWS.
				 , cols=&COL_NAMESN.));

	/* Add header row column by column (requirement of the API) */
	%do i = 1 %to &COL_NAMESN.;
		%gdoc_updateCell(&NEW_SHEET_TITLE.,1,&i.,&&COL_NAMES&i.);
	%end;

	/* Add the rest of the data and create payload */
	data _null_;
		set MODIFIED_ORIG_DS end=last_rec;

		/* Open the dataset and determine the number of rows */
		%let DSID = %sysfunc(open(&DS_LIB..&DS., IS)); %let ROWS = %eval(%sysfunc(attrn(&DSID, NLOBS))); %let CLOSE_ID = %sysfunc(close(&DSID));

		%do i = 1 %to &ROWS.;
			if _N_ = &i. then do;
				file "&DRIVE.:\&TEMP.\payload&i..xml";
				put '<entry xmlns="http://www.w3.org/2005/Atom"';
				put '		xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">';
				%do_over(COL_NAMES, phrase=put "	<gsx:?>" ? "</gsx:?>";);
				put '</entry>'; 
			end;
		%end;
		%do i = 1 %to &ROWS.;
			%gdoc_addRow(&NEW_SHEET_TITLE., &DRIVE.:\&TEMP.\payload&i..xml);
		%end;
		option noxwait;
		x "del &DRIVE.:\&TEMP.\payload*.xml";
	run;

	/* Fix column headers back to expected names */
	%do i = 1 %to &ORIG_COL_NAMESN.;
		%gdoc_updateCell(&NEW_SHEET_TITLE.,1,&i.,&&ORIG_COL_NAMES&i.);
	%end;

	/* Remove temp dataset */
	proc datasets library=work nodetails nolist; delete modified_orig_ds; run; quit;
%mend;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_GETWORKSHEETLIST
Desc 	    : Get a list of Google Docs worksheets (e.g. SHEET1)
Parameters: SPREADSHEET_NAME
\**************************************************************************/;
%macro gdoc_getWorksheetList(SPREADSHEET_NAME);

	/* Get worksheet ID for new sheet */
	data _null_;
		set gdoc_wslist;
		where sheet_name = "&SPREADSHEET_NAME.";
		call symputx('URL',trim(WORKSHEET_ID));
	run;

	/* Set header and get list */	
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH_S.);
	proc http headerin=hin headerout=hout out=out url="&URL." method="GET"; run;
	x "&DRIVE.:\&TIDY. -xml -indent -wrap 256 -quiet -m &DRIVE.:\&TEMP.\out.txt";

	/* Read the response file and get the URLs */
	data gdoc_worksheets;
		format line $256.;
		infile out lrecl=256 DLM='0A'x;
		input line;
	run;

	/* Grab URLs for the documents and strip extraenous tags */
	data gdoc_worksheets(where=(id ^= ''));
		set gdoc_worksheets;
		format id $100.;

		/* New list entry */
		if prxmatch('/etag/',line) then rank + 1;

		/* ID */
		if prxmatch ('/<id>/',line) then do;
			line = prxchange('s/^.*<id>//',-1,line);
			line = prxchange('s/<\/id>.*$//',-1,line);
			line = prxchange('s/worksheets/list/',-1,line);
			id = 'worksheet_id';
		end;

		/* Edit Link */
		if prxmatch ('/link rel=.edit/',line) then do;
			line = prxchange("s/^.*href=.?//",-1,line);
			line = prxchange("s/'.\/>.*$//",-1,line);
			id = 'edit_url';
		end;

		/* Name of document */
		if prxmatch ('/<\/title>/',line) then do;
			line = prxchange('s/^.*<title>//',-1,line);
			line = prxchange('s/<\/title>.*$//',-1,line);
			id = 'worksheet_name';
		end;
	run;

	/* Put each spreadsheet listing on a single line */
	proc transpose data=gdoc_worksheets
				   out=gdoc_sheetlst(drop=_NAME_ rank); 
				   var line;
				   id id;
				   by rank;
	run;

	/* Remove temporary datasets */
	proc datasets library=work nodetails nolist;
		delete gdoc_worksheets;
	run;
	quit;
%mend;

/**************************************************************************\
Copyright (c) 2013 MarketShare Partners, LLC. All rights reserved
Content is proprietary to MarketShare, and this paper is not intended to convey any rights therein.
Macro	    : %GDOC_DELETESHEET
Desc 	    : Delete a worksheet
Parameters: TITLE - Title of worksheet
\**************************************************************************/;
%macro gdoc_deleteSheet(title);
	/* Get EDIT_URL for sheet */
	data _null_;
		set gdoc_sheetlst;
		where WORKSHEET_NAME = "&TITLE.";
		call symputx('URL',trim(EDIT_URL));
	run;

	/* Set header and delete worksheet (this is final and gives no warning) */	
	%sas_setHeader(header=Authorization: GoogleLogin Auth=&AUTH_S.
				 , header2=%str(If!Match: *));
	proc http headerin=hin headerout=hout out=out ct="application/atom+xml" url="&URL." method="DELETE"; run;
%mend;
