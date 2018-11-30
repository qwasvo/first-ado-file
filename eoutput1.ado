* eoutput1 - Export one-way tables into Excel with formatting - see eoutput2 for two-way tables
* Philippe Dufresne - p.or.duff@gmail.com
* v1.0, 29-Nov-2018

program define eoutput1, rclass

	version 14
	
	syntax varlist 
	
	gettoken strat var : varlist
	
	display "the strata variable is `strat'" 
	
	local row=2
	
	gettoken var1 : var	
		*Variables cat
	qui: sum `var'
	forvalues g = `r(min)'/`r(max)' {	
		putexcel A`row' = ("`var'==`g'")
		local ++row
		
		foreach var of varlist `var' {
			qui: sum `var'												/*creates macro with the number of the variable's categories in it*/

			putexcel A`row'=("`var'")		
			
			local var_lab : var label `var'
			local row1 = `row'+1
			putexcel A`row1' = ("`var_lab'")
			
			local rowvar = `row'										/*Keep the starting row as a reference*/	
		
			forvalues i = `r(min)'/`r(max)' {							/*creates table for every possible value of every variable above*/			
				putexcel B`row'=`i'

				local val_lab : value label `var'						/*puts name of value label in a macro*/
				local val_lab1 : label `val_lab' `i'					/*puts indivual value labels in a macro*/
				putexcel C`row' = ("`val_lab1'")
			
				local ++row
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
