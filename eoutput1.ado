* eoutput1 - Export one-way tables into Excel with formatting - see eoutput2 for two-way tables
* Philippe Dufresne - p.or.duff@gmail.com
* v1.0, 29-Nov-2018

program define eoutput1, rclass

	version 14
	
	syntax varlist(min=1 max=1) [if] [in] using/, [ sheet(string) Format(string) replace modify]
	
	*~~~~~~~~~~~~~~~ export to Excel ~~~~~~~~~~~~~~~*
	* --> export both data and labels to Excel; that is the most intense part due to write opearions; takes ~90% computational time
	* Set default sheet name
	if "`sheet'" == "" local sheet = "Data"
	
	* Replace or modify options
	if "`replace'" == "" & "`modify'" == "" local replace = "replace"
	
	* Prepare an Excel file
	local ext = substr("`using'",-4,.)
	//writing into xls is faster than to xlsx but formatting (merged cells, borders) doesn't work as it should
	if !inlist("`ext'", "xlsx", ".xls")  	local using "`using'.xlsx"	
	qui putexcel clear
	qui putexcel set `"`using'"', sheet("`sheet'") `replace' `modify'
	
	
	*~~~~~~~~~~~~ program ~~~~~~~~~~~~~~*

	local row=2
	
	qui: sum `varlist'
	putexcel A`row'=("`varlist'")		
	
	local var_lab : var label `varlist'
	local row1 = `row'+1
	putexcel A`row1' = ("`var_lab'")
	
	local rowvar = `row'										/*Keep the starting row as a reference*/	
	
	if `r(min)'==1 {
		loc nbrow = `r(max)'
	}
	else {
		loc nbrow = `r(max)'+1
	}
	
	matrix mat1 = J(`nbrow',3,.)
	
	forvalues i = `r(min)'/`r(max)' {							/*creates table for every possible value of every variable above*/			
		local val_lab : value label `varlist'					/*puts name of value label in a macro*/
		local val_lab`i' : label `val_lab' `i'					/*puts indivual value labels in a macro*/
		local allval_lab = "`val_lab`i''"
	}	

	local row = `rowvar'										/*goes back from the start of labels, to input results*/
	
	qui: sum `varlist'
	forvalues i = `r(min)'/`r(max)' {
		if `i' == 0 {
			matrix mat1[`i'+1,1] = `i'
		}
		else {
			matrix mat1[`i',1] = `i'
		}
		
		qui: tab1 `varlist' if `varlist'==`i'
		loc n`i' = `r(N)'
		if `i' == 0 {
			matrix mat1[`i'+1,2] = `n`i''
		}
		else {
			matrix mat1[`i',2] = `n`i''
		}
		
		qui: tab1 `varlist',m
		loc N`i' = `r(N)'
		
		loc per`i' = round(`n`i''/`N`i''*100, 0.1)
		if `i' == 0 {
			matrix mat1[`i'+1,3] = `per`i''
		}
		else {
			matrix mat1[`i',3] = `per`i''
		}
	}
	
	forvalues i = `r(min)'/`r(max)'
		matrix rownames mat1 = "`allval_lab'"
	
	matrix colnames mat1 = value N %

	mat li mat1
	*putexcel B`row' = (matrix(mat1), names)
	local row = row + `i' +1
}

end
