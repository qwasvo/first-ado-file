* eoutput1 - Export one-way tables into Excel with formatting - see eoutput2 for two-way tables
* Philippe Dufresne - p.or.duff@gmail.com
* v1.0, 29-Nov-2018

program define eoutput1, rclass

	version 14
	
	syntax varlist [if] [in] using/, [ sheet(string) Format(string) replace modify]
	
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
	
	gettoken var1 : `varlist'	
		*Variables cat
	qui: sum `var1'
	forvalues g = `r(min)'/`r(max)' {	
		putexcel A`row' = ("`var1'==`g'")
		local ++row
		
		foreach var of varlist `var1' {
			qui: sum `var1'												/*creates macro with the number of the variable's categories in it*/

			putexcel A`row'=("`var1'")		
			
			local var_lab : var label `var1'
			local row1 = `row'+1
			putexcel A`row1' = ("`var_lab'")
			
			local rowvar = `row'										/*Keep the starting row as a reference*/	
		
			forvalues i = `r(min)'/`r(max)' {							/*creates table for every possible value of every variable above*/			

				local val_lab : value label `var'						/*puts name of value label in a macro*/
				local val_lab`i' : label `val_lab' `i'					/*puts indivual value labels in a macro*/
			}	

			local row = `rowvar'										/*goes back from the start of labels, to input results*/
			
			qui: svy: tab ``m'' ``k'', count
			qui: ereturn display										/*Creates matrix of results r(table)*/
			mat res = r(table)
			loc nbcol = colsof(res)
			
			qui: sum `var'
			forvalues i = `r(min)'/`nbcol' {
				loc br = round(res[1,`i'], 0.01)						/*return results. line 1 = coef, line 5 & 6 = 95% CI*/
				loc b = res[1,`i']
				loc ll = round(res[5,`i'], 0.01)
				loc ul = round(res[6,`i'], 0.01)
				loc se = res[2,`i']
				loc cv = `se'/`b'
				
				putexcel D`row'=("`br' (`ll' - `ul')")
				putexcel F`row'=("`cv'")
				
				local ++row
			}
		}
	}
}

end
