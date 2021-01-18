*** Archive code for LR among ANC recipients *** 

*** Among ANC recipients *** 

* ANC provider type
svy: logistic SWdelivprob_convuls i.provider_code if anyanc==1 
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = PHCP both _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A59 = "ANC Provider Type", bold overwritefmt
putexcel E60 = "N = ", right bold overwritefmt
putexcel F60 = `e(N)', hcenter bold overwritefmt
putexcel A61 = matrix(results), names nformat(number_d2)

* Maternal assessment at ANC 
svy: logistic SWdelivprob_convuls maternal_assess_all if anyanc==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A67 = "Received all 5 assessment at ANC", bold overwritefmt
putexcel E68 = "N = ", right bold overwritefmt
putexcel F68 = `e(N)', hcenter bold overwritefmt
putexcel A69 = matrix(results), names nformat(number_d2)

svy: logistic SWdelivprob_convuls i.maternal_assess_cat if anyanc==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1-3 4 5 _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A73 = "Maternal assessment score at ANC", bold overwritefmt
putexcel E74 = "N = ", right bold overwritefmt
putexcel F74 = `e(N)', hcenter bold overwritefmt
putexcel A75 = matrix(results), names nformat(number_d2)

* BP measurement alone
svy: logistic SWdelivprob_convuls SWanc_bp if anyanc==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A81 = "BP measurement at ANC", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82 = `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)

* Weight measurement alone
svy: logistic SWdelivprob_convuls SWanc_weight if anyanc==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A87 = "Weight measurement at ANC", bold overwritefmt
putexcel E88 = "N = ", right bold overwritefmt
putexcel F88 = `e(N)', hcenter bold overwritefmt
putexcel A89 = matrix(results), names nformat(number_d2)

* Urine test alone
svy: logistic SWdelivprob_convuls SWanc_urine if anyanc==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A93 = "Urine test at ANC", bold overwritefmt
putexcel E94 = "N = ", right bold overwritefmt
putexcel F94 = `e(N)', hcenter bold overwritefmt
putexcel A95 = matrix(results), names nformat(number_d2)

* Birth readiness discussion 
* Categorical 
svy: logistic SWdelivprob_convuls i.birth_readiness_cat if anyanc==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "ll"], temp[2..., "ul"], temp[2..., "pvalue"]
matrix rownames results = 1-3 4 5 _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A99 = "Birth readiness discussion (categorial)", bold overwritefmt
putexcel E100 = "N = ", right bold overwritefmt
putexcel F100 = `e(N)', hcenter bold overwritefmt
putexcel A101 = matrix(results), names nformat(number_d2)

* Binary 
svy: logistic SWdelivprob_convuls birth_readiness_all if anyanc==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A107 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel E108 = "N = ", right bold overwritefmt
putexcel F108 = `e(N)', hcenter bold overwritefmt
putexcel A109 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls danger_sign_coun if anyanc==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A113 = "PE-related danger sign counseling", bold overwritefmt
putexcel E114 = "N = ", right bold overwritefmt
putexcel F114 = `e(N)', hcenter bold overwritefmt
putexcel A115 = matrix(results), names nformat(number_d2)

* Nutrition counseling
svy: logistic SWdelivprob_convuls SWanc_nd_info_yn if anyanc==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A119 = "Nutritional counseling", bold overwritefmt
putexcel E120 = "N = ", right bold overwritefmt
putexcel F120 = `e(N)', hcenter bold overwritefmt
putexcel A121 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls SWanc_tt_inject if anyanc==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A125 = "Tetanus injection at ANC", bold overwritefmt
putexcel E126 = "N = ", right bold overwritefmt
putexcel F126 = `e(N)', hcenter bold overwritefmt
putexcel A127 = matrix(results), names nformat(number_d2)
