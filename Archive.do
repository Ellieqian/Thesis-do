*** GA at first ANC visit ***
* Generate GA at first visit variable
recode SWanc_phcp_timing(0 = 1)
recode SWanc_hew_timing(0 = 1)
gen ga_first_anc=SWanc_hew_timing
replace ga_first_anc=SWanc_phcp_timing if SWanc_phcp_timing<anc_hew_timing & SWanc_phcp_timing!=.
replace ga_first_anc=. if SWanc_hew_timing==. & SWanc_phcp_timing==.
replace ga_first_anc=. if anyanc==0
label var ga_first_anc "GA (in months) at first ANC"

* Create categories 
gen ga_first_anc_cat=0
replace ga_first_anc_cat=1 if ga_first_anc<4 & ga_first_anc>=1 
replace ga_first_anc_cat=2 if ga_first_anc>=4 & ga_first_anc<=6 
replace ga_first_anc_cat=3 if ga_first_anc>6 & ga_first_anc<=9 
replace ga_first_anc_cat=4 if ga_first_anc==.
replace ga_first_anc_cat=0 if anyanc==0
label define ga_anc_l 0 "No ANC" 1 "First trimester" 2 "Second trimester" 3 "Third trimester" 4 "Missing ANC timing"
label val ga_first_anc_cat ga_anc_l
tab ga_first_anc_cat

gen maternal_assess_score=SWanc_bp + SWanc_weight + SWanc_urine + SWanc_blood + SWanc_stool
replace maternal_assess_score=0 if anyanc==0

gen maternal_assess_cat=0 
replace maternal_assess_cat=1 if maternal_assess_score>=1 & maternal_assess_score<=3
replace maternal_assess_cat=2 if maternal_assess_score==4
replace maternal_assess_cat=3 if maternal_assess_score==5
replace maternal_assess_cat=0 if anyanc==0
label define maternal_assess_cat 0 "None" 1 "1-3 assessments" 2 "4 assessments" 3 "All 5 assessments"
label val maternal_assess_cat maternal_assess_cat


* Generate binary composite birth/complication readiness indicator
gen birth_readiness_all=0
replace birth_readiness_all=1 if (SWanc_delivery_place==1 & SWanc_delivery_skilled==1 & ///
									SWanc_emergency_transport==1 & SWanc_danger_place==1 & ///
									SWanc_danger_migraine==1 & SWanc_danger_hbp==1 & SWanc_danger_edema==1 & ///
									SWanc_danger_convuls==1 & SWanc_danger_bleeding==1) & anyanc==1
label var birth_readiness_all "At ANC discussed all birth/complication readiness topics"
label val birth_readiness_all yesno


gen birth_readiness_c=SWanc_delivery_place + SWanc_delivery_skilled + SWanc_emergency_transport + SWanc_danger_place + ///
						SWanc_danger_migraine + SWanc_danger_hbp + SWanc_danger_edema + SWanc_danger_convuls + SWanc_danger_bleeding
								
gen birth_readiness_cat=0
replace birth_readiness_cat=1 if birth_readiness_c<=3 & birth_readiness_c>=1 
replace birth_readiness_cat=2 if birth_readiness_c>3 & birth_readiness_c<=6 
replace birth_readiness_cat=3 if birth_readiness_c>6 & birth_readiness_c<=9 
replace birth_readiness_cat=0 if anyanc==0
label var birth_readiness_cat "Birth readiness discussion (categorial)"
label define birthl 0 "None" 1 "1-3 topics" 2 "4-6 topics" 3 "7-9 topics"
label val birth_readiness_cat birthl


* All 5 assessments
gen maternal_assess_all=0
replace maternal_assess_all=1 if (SWanc_bp==1 & SWanc_weight==1 & SWanc_urine==1 & SWanc_blood==1 & SWanc_stool==1)
label var maternal_assess_all "Received all 5 maternal assessments"
label val maternal_assess_all yesno
tab maternal_assess_all

* Key ANC services binary indicators 
gen anc_key_services=0 
replace anc_key_services=1 if (SWanc_bp==1 & SWanc_iron==1 & SWanc_blood==1 & ///
								SWanc_urine==1 & SWanc_syph_test==1 & SWanc_hiv_test==1) 
label var anc_key_services "At ANC had : BP, urine and blood sampled, tested for syphilis & HIV, + took iron during preg"
label val anc_key_services yesno



#delimit;
table1_mc,  			
by(SWdelivprob_convuls) 			
vars( 						
age_cat cat %4.1f \
urban bin %4.1f \
education cat %4.1f \
wealthquintile cat %4.1f \
parity4 cat %4.1f \
unintended_preg bin %4.1f \
facility_skilled cat %4.1f \
preg_comp bin %4.1f \
anc_num_cat cat %4.1f \
) 
nospace onecol 
saving(Tables/table1_$date.xlsx, replace) // save table 1
;
#delimit cr



*=========================================================
* EDA *
*=========================================================

* Tabulation of outcome and potential confounders
* Background characteristics
/*
tabout SWdelivprob_convuls age_cat [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls education [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls urban [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls married [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls parity_cat [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 

* ANC and delivery care 
tabout SWdelivprob_convuls anyanc [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls ANC4 [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls anc_key_services [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls birth_readiness_all [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls preg_comp [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls facility_deliv [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls sba [aw=SWweight]  using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
tabout SWdelivprob_convuls SWcaesarean_delivery [aw=SWweight]  using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 


* Lowess smoothed regressions -- continuous Xs

* Age
lowess SWdelivprob_convuls FQ_age, name(age,replace) logit msymbol(i) bwidth(0.8) yline(0) t1("Lowess plot: log odds age -vs- any complication")
graph export graphs/lowess_age.png,replace

* Parity
lowess SWdelivprob_convuls total_birth, name(parity,replace) logit msymbol(i) bwidth(0.8) yline(0) t1("Lowess plot: log odds parity -vs- any complication")
graph export graphs/lowess_parity.png,replace

* Number of ANC 
lowess SWdelivprob_convuls anc_tot, name(anc,replace) logit msymbol(i) bwidth(0.8) yline(0) t1("Lowess plot: log odds ANC -vs- any complication")
graph export graphs/lowess_anc.png,replace

* GA at first ANC
lowess SWdelivprob_convuls ga_first_anc if ga_first_anc!=0, name(ga_anc,replace) logit msymbol(i) bwidth(0.8) yline(0) t1("Lowess plot: log odds GA at first ANC -vs- any complication")
graph export graphs/lowess_ga.png,replace

* k x 2 tables -- categorical Xs
for var age_cat:  tabulate X SWdelivprob_convuls, chi2 exact column
for var parity_cat:  tabulate X SWdelivprob_convuls, chi2 exact column
for var education:  tabulate X SWdelivprob_convuls, chi2 exact column
for var anc_num_cat:  tabulate X SWdelivprob_convuls, chi2 exact column


graph box FQ_age, name(age_provider,replace) over(provider_code, relabel(1 "No ANC" 2 "HEW only" 3 "PHCP only" 4 "Both")) ///
	marker(1,mlab(FQ_age)) t1("Age by provider type") scale(1.2) ytitle("Age (years)") 
	graph export graphs/age_provider.png, replace

graph box total_birth, name(parity_provider,replace) over(provider_code, relabel(1 "No ANC" 2 "HEW only" 3 "PHCP only" 4 "Both")) ///
	marker(1,mlab(total_birth)) t1("Parity by provider type") scale(1.2) ytitle("Number of children") 
	graph export graphs/parity_provider.png, replace
	
	

*=========================================================
* Logistic regression *
*=========================================================

*======================
* TEST ON WEIGHTING
*======================
/*
svyset EA [pweight=SWweight], strata(strata) singleunit(scaled)
svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_num_cat if select

svyset EA, strata(strata) singleunit(scaled)
svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_num_cat if select

svyset EA [pweight=SWweight], strata(strata) singleunit(scaled)
svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_place if select

svyset EA, strata(strata) singleunit(scaled)
svy: logistic eclampsia i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_place if select
*/



*** Crude effect among all women ***

* ANC frequency
svy: logistic SWdelivprob_convuls i.anc_num_cat if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4+ _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) replace
putexcel A1 = "ANC Frequency", bold overwritefmt
putexcel E2 = "N = ", right bold overwritefmt
putexcel F2 = `e(N)', hcenter bold overwritefmt
putexcel A3 = matrix(results), names nformat(number_d2)

* ANC provider type and location 
svy: logistic SWdelivprob_convuls i.anc_place if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = HEW PHCP_HF PHCP_hos _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A8 = "ANC Provider Type and Location", bold overwritefmt
putexcel E9 = "N = ", right bold overwritefmt
putexcel F9 = `e(N)', hcenter bold overwritefmt
putexcel A10 = matrix(results), names nformat(number_d2)

* ANC timing 
svy: logistic SWdelivprob_convuls i.ga_first_anc_cat if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = first_trimester second_trimester third_trimester missing _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A16 = "Timing at first ANC", bold overwritefmt
putexcel E17 = "N = ", right bold overwritefmt
putexcel F17 = `e(N)', hcenter bold overwritefmt
putexcel A18 = matrix(results), names nformat(number_d2)

* ANC in the first trimester 
svy: logistic SWdelivprob_convuls i.first_tri if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = AfterFristTri InFristTri _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A25 = "Received ANC in the first trimester", bold overwritefmt
putexcel E26 = "N = ", right bold overwritefmt
putexcel F26 = `e(N)', hcenter bold overwritefmt
putexcel A27 = matrix(results), names nformat(number_d2)

* BP measurement 
svy: logistic SWdelivprob_convuls i.bp if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A33 = "BP measurement at ANC", bold overwritefmt
putexcel E34 = "N = ", right bold overwritefmt
putexcel F34 = `e(N)', hcenter bold overwritefmt
putexcel A35 = matrix(results), names nformat(number_d2)

* Weight measurement 
svy: logistic SWdelivprob_convuls i.weight if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A41 = "Weight measurement at ANC", bold overwritefmt
putexcel E42 = "N = ", right bold overwritefmt
putexcel F42 = `e(N)', hcenter bold overwritefmt
putexcel A43 = matrix(results), names nformat(number_d2)

* Urine test 
svy: logistic SWdelivprob_convuls i.urine if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A49 = "Urine tested at ANC", bold overwritefmt
putexcel E50 = "N = ", right bold overwritefmt
putexcel F50 = `e(N)', hcenter bold overwritefmt
putexcel A51 = matrix(results), names nformat(number_d2)

* Blood sample
svy: logistic SWdelivprob_convuls i.blood if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A57 = "Blood sample taken at ANC", bold overwritefmt
putexcel E58 = "N = ", right bold overwritefmt
putexcel F58 = `e(N)', hcenter bold overwritefmt
putexcel A60 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls i.danger_sign_coun if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A65 = "PE-related danger sign counseling", bold overwritefmt
putexcel E66 = "N = ", right bold overwritefmt
putexcel F66 = `e(N)', hcenter bold overwritefmt
putexcel A67 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls i.tt if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A73 = "Tetanus injection at ANC", bold overwritefmt
putexcel E74 = "N = ", right bold overwritefmt
putexcel F74 = `e(N)', hcenter bold overwritefmt
putexcel A75 = matrix(results), names nformat(number_d2)

* Iron counseling at ANC 
svy: logistic SWdelivprob_convuls i.iron if select
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude) modify
putexcel A81 = "Iron counseling at ANC", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82= `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)


*** Crude effect among women with signs of preeclampsia ***

* ANC frequency
svy: logistic SWdelivprob_convuls i.anc_num_cat if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = 1-3 4+ _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A1 = "ANC Frequency", bold overwritefmt
putexcel E2 = "N = ", right bold overwritefmt
putexcel F2 = `e(N)', hcenter bold overwritefmt
putexcel A3 = matrix(results), names nformat(number_d2)

* ANC provider type and location 
svy: logistic SWdelivprob_convuls i.anc_place if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = HEW PHCP_HF PHCP_hos _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A8 = "ANC Provider Type and Location", bold overwritefmt
putexcel E9 = "N = ", right bold overwritefmt
putexcel F9 = `e(N)', hcenter bold overwritefmt
putexcel A10 = matrix(results), names nformat(number_d2)

* ANC timing 
svy: logistic SWdelivprob_convuls i.ga_first_anc_cat if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = first_trimester second_trimester third_trimester missing _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A16 = "Timing at first ANC", bold overwritefmt
putexcel E17 = "N = ", right bold overwritefmt
putexcel F17 = `e(N)', hcenter bold overwritefmt
putexcel A18 = matrix(results), names nformat(number_d2)

* ANC in the first trimester 
svy: logistic SWdelivprob_convuls i.first_tri if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = AfterFirstTri InFirstTri _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A25 = "Received ANC in the first trimester", bold overwritefmt
putexcel E26 = "N = ", right bold overwritefmt
putexcel F26 = `e(N)', hcenter bold overwritefmt
putexcel A27 = matrix(results), names nformat(number_d2)

* BP measurement 
svy: logistic SWdelivprob_convuls i.bp if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A33 = "BP measurement at ANC", bold overwritefmt
putexcel E34 = "N = ", right bold overwritefmt
putexcel F34 = `e(N)', hcenter bold overwritefmt
putexcel A35 = matrix(results), names nformat(number_d2)

* Weight measurement 
svy: logistic SWdelivprob_convuls i.weight if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A41 = "Weight measurement at ANC", bold overwritefmt
putexcel E42 = "N = ", right bold overwritefmt
putexcel F42 = `e(N)', hcenter bold overwritefmt
putexcel A43 = matrix(results), names nformat(number_d2)

* Urine test 
svy: logistic SWdelivprob_convuls i.urine if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A49 = "Urine test at ANC", bold overwritefmt
putexcel E50 = "N = ", right bold overwritefmt
putexcel F50 = `e(N)', hcenter bold overwritefmt
putexcel A51 = matrix(results), names nformat(number_d2)

* Blood sample
svy: logistic SWdelivprob_convuls i.blood if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A57 = "Blood sample taken at ANC", bold overwritefmt
putexcel E58 = "N = ", right bold overwritefmt
putexcel F58 = `e(N)', hcenter bold overwritefmt
putexcel A60 = matrix(results), names nformat(number_d2)

* PE-related danger sign counseling
svy: logistic SWdelivprob_convuls i.danger_sign_coun if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A65 = "PE-related danger sign counseling", bold overwritefmt
putexcel E66 = "N = ", right bold overwritefmt
putexcel F66 = `e(N)', hcenter bold overwritefmt
putexcel A67 = matrix(results), names nformat(number_d2)

* Tetanus injection
recode SWanc_tt_inject(. -88 -99 = 0)
svy: logistic SWdelivprob_convuls i.tt if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A73 = "Tetanus injection at ANC", bold overwritefmt
putexcel E74 = "N = ", right bold overwritefmt
putexcel F74 = `e(N)', hcenter bold overwritefmt
putexcel A75 = matrix(results), names nformat(number_d2)

* Iron counseling at ANC 
svy: logistic SWdelivprob_convuls i.iron if select & preg_comp==1
mat temp = r(table)'
mat results = temp[2..., "b"], temp[2..., "pvalue"], temp[2..., "ll"], temp[2..., "ul"]
matrix rownames results = No Yes _cons
matrix colnames results = OR P-value CI-lower CI-upper 
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A81 = "Iron counseling at ANC", bold overwritefmt
putexcel E82 = "N = ", right bold overwritefmt
putexcel F82= `e(N)', hcenter bold overwritefmt
putexcel A83 = matrix(results), names nformat(number_d2)


*** Archive code for LR among ANC recipients *** 

*** Among ANC recipients *** 

* ANC provider type
svy: logistic SWdelivprob_convuls i.provider_code if preg_comp==1 
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
svy: logistic SWdelivprob_convuls maternal_assess_all if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A67 = "Received all 5 assessment at ANC", bold overwritefmt
putexcel E68 = "N = ", right bold overwritefmt
putexcel F68 = `e(N)', hcenter bold overwritefmt
putexcel A69 = matrix(results), names nformat(number_d2)

svy: logistic SWdelivprob_convuls i.maternal_assess_cat if preg_comp==1
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
svy: logistic SWdelivprob_convuls SWanc_bp if preg_comp==1
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
svy: logistic SWdelivprob_convuls SWanc_weight if preg_comp==1
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
svy: logistic SWdelivprob_convuls SWanc_urine if preg_comp==1
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
svy: logistic SWdelivprob_convuls i.birth_readiness_cat if preg_comp==1
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
svy: logistic SWdelivprob_convuls birth_readiness_all if preg_comp==1
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
svy: logistic SWdelivprob_convuls danger_sign_coun if preg_comp==1
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
svy: logistic SWdelivprob_convuls SWanc_nd_info_yn if preg_comp==1
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
svy: logistic SWdelivprob_convuls SWanc_tt_inject if preg_comp==1
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "ll"], temp[1..., "ul"], temp[1..., "pvalue"]
matrix rownames results = Yes _cons
matrix colnames results = OR CI-lower CI-upper P-value
putexcel set Regression_output_$date.xlsx, sheet(crude-1) modify
putexcel A125 = "Tetanus injection at ANC", bold overwritefmt
putexcel E126 = "N = ", right bold overwritefmt
putexcel F126 = `e(N)', hcenter bold overwritefmt
putexcel A127 = matrix(results), names nformat(number_d2)



*======================================================*
* Among women who were pregnancy at enrollment  * 
*======================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_num_cat if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H1 = "Adjusted logistic regression results among women who were pregnancy at enrollment", bold overwritefmt
putexcel H3 = "ANC Frequency", bold overwritefmt
putexcel L4 = "N = ", right bold overwritefmt
putexcel M4 = `e(N)', hcenter bold overwritefmt
putexcel H5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.provider_code if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H23 = "ANC Provider Type", bold overwritefmt
putexcel L24 = "N = ", right bold overwritefmt
putexcel M24 = `e(N)', hcenter bold overwritefmt
putexcel H25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.maternal_assess_cat if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel L45 = "N = ", right bold overwritefmt
putexcel M45 = `e(N)', hcenter bold overwritefmt
putexcel H46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled maternal_assess_all if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel L66 = "N = ", right bold overwritefmt
putexcel M66 = `e(N)', hcenter bold overwritefmt
putexcel H67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled danger_sign_coun if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H83 = "PE/E danger sign counseling", bold overwritefmt
putexcel L84 = "N = ", right bold overwritefmt
putexcel M84 = `e(N)', hcenter bold overwritefmt
putexcel H85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.birth_readiness_cat if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel L101 = "N = ", right bold overwritefmt
putexcel M101 = `e(N)', hcenter bold overwritefmt
putexcel H102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled birth_readiness_all if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel L123 = "N = ", right bold overwritefmt
putexcel M123 = `e(N)', hcenter bold overwritefmt
putexcel H124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_nd_info_yn if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H140 = "Nutritional counseling", bold overwritefmt
putexcel L141 = "N = ", right bold overwritefmt
putexcel M141 = `e(N)', hcenter bold overwritefmt
putexcel H142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_tt_inject if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H158 = "Tetanus injection", bold overwritefmt
putexcel L159 = "N = ", right bold overwritefmt
putexcel M159 = `e(N)', hcenter bold overwritefmt
putexcel H160 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_bp if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H176 = "BP measurement at ANC", bold overwritefmt
putexcel L177 = "N = ", right bold overwritefmt
putexcel M177 = `e(N)', hcenter bold overwritefmt
putexcel H178= matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_weight if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H194 = "Weight measurement at ANC", bold overwritefmt
putexcel L195 = "N = ", right bold overwritefmt
putexcel M195 = `e(N)', hcenter bold overwritefmt
putexcel H196= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_urine if select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel H212 = "Urine test at ANC", bold overwritefmt
putexcel L213 = "N = ", right bold overwritefmt
putexcel M213 = `e(N)', hcenter bold overwritefmt
putexcel H214= matrix(results), names nformat(number_d2)

*======================================================*
* Excluding 5-9 postpartum women * 
*======================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.anc_num_cat if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O1 = "Adjusted logistic regression results excluding 5-9 postpartum women", bold overwritefmt
putexcel O3 = "ANC Frequency", bold overwritefmt
putexcel S4 = "N = ", right bold overwritefmt
putexcel T4 = `e(N)', hcenter bold overwritefmt
putexcel O5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.provider_code if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O23 = "ANC Provider Type", bold overwritefmt
putexcel S24 = "N = ", right bold overwritefmt
putexcel T24 = `e(N)', hcenter bold overwritefmt
putexcel O25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.maternal_assess_cat if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel S45 = "N = ", right bold overwritefmt
putexcel T45 = `e(N)', hcenter bold overwritefmt
putexcel O46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled maternal_assess_all if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel S66 = "N = ", right bold overwritefmt
putexcel T66 = `e(N)', hcenter bold overwritefmt
putexcel O67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled danger_sign_coun if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O83 = "PE/E danger sign counseling", bold overwritefmt
putexcel S84 = "N = ", right bold overwritefmt
putexcel T84 = `e(N)', hcenter bold overwritefmt
putexcel O85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled i.birth_readiness_cat if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel S101 = "N = ", right bold overwritefmt
putexcel T101 = `e(N)', hcenter bold overwritefmt
putexcel O102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled birth_readiness_all if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel S123 = "N = ", right bold overwritefmt
putexcel T123 = `e(N)', hcenter bold overwritefmt
putexcel O124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_nd_info_yn if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O140 = "Nutritional counseling", bold overwritefmt
putexcel S141 = "N = ", right bold overwritefmt
putexcel T141 = `e(N)', hcenter bold overwritefmt
putexcel O142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_tt_inject if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O158 = "Tetanus injection", bold overwritefmt
putexcel S159 = "N = ", right bold overwritefmt
putexcel T159 = `e(N)', hcenter bold overwritefmt
putexcel O160 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_bp if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O176 = "BP measurement at ANC", bold overwritefmt
putexcel S177 = "N = ", right bold overwritefmt
putexcel T177 = `e(N)', hcenter bold overwritefmt
putexcel O178= matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_weight if select & baseline_status!=3
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O194 = "Weight measurement at ANC", bold overwritefmt
putexcel S195 = "N = ", right bold overwritefmt
putexcel T195 = `e(N)', hcenter bold overwritefmt
putexcel O196= matrix(results), names nformat(number_d2)

*** Urine test*** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat preg_comp i.facility_skilled SWanc_urine if select & baseline_status!=3 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted) modify
putexcel O212 = "Urine test at ANC", bold overwritefmt
putexcel S213 = "N = ", right bold overwritefmt
putexcel T213 = `e(N)', hcenter bold overwritefmt
putexcel O214= matrix(results), names nformat(number_d2)



*==================================================================*
* Among women who were pregnancy at enrollment and had sign of PE* 
*==================================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.anc_num_cat if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H1 = "Adjusted logistic regression results among women who were pregnancy at enrollment and had PE signs", bold overwritefmt
putexcel H3 = "ANC Frequency", bold overwritefmt
putexcel L4 = "N = ", right bold overwritefmt
putexcel M4 = `e(N)', hcenter bold overwritefmt
putexcel H5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.provider_code if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H23 = "ANC Provider Type", bold overwritefmt
putexcel L24 = "N = ", right bold overwritefmt
putexcel M24 = `e(N)', hcenter bold overwritefmt
putexcel H25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.maternal_assess_cat if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel L45 = "N = ", right bold overwritefmt
putexcel M45 = `e(N)', hcenter bold overwritefmt
putexcel H46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled maternal_assess_all if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel L66 = "N = ", right bold overwritefmt
putexcel M66 = `e(N)', hcenter bold overwritefmt
putexcel H67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled danger_sign_coun if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H83 = "PE/E danger sign counseling", bold overwritefmt
putexcel L84 = "N = ", right bold overwritefmt
putexcel M84 = `e(N)', hcenter bold overwritefmt
putexcel H85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.birth_readiness_cat if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel L101 = "N = ", right bold overwritefmt
putexcel M101 = `e(N)', hcenter bold overwritefmt
putexcel H102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled birth_readiness_all if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel L123 = "N = ", right bold overwritefmt
putexcel M123 = `e(N)', hcenter bold overwritefmt
putexcel H124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_nd_info_yn if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H140 = "Nutritional counseling", bold overwritefmt
putexcel L141 = "N = ", right bold overwritefmt
putexcel M141 = `e(N)', hcenter bold overwritefmt
putexcel H142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_tt_inject if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H158 = "Tetanus injection", bold overwritefmt
putexcel L159 = "N = ", right bold overwritefmt
putexcel M159 = `e(N)', hcenter bold overwritefmt
putexcel H160 = matrix(results), names nformat(number_d2)

*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_bp if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H176 = "Blood pressure measurement at ANC", bold overwritefmt
putexcel L177 = "N = ", right bold overwritefmt
putexcel M177 = `e(N)', hcenter bold overwritefmt
putexcel H178 = matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_weight if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H194 = "Weight measurement at ANC", bold overwritefmt
putexcel L195 = "N = ", right bold overwritefmt
putexcel M195 = `e(N)', hcenter bold overwritefmt
putexcel H196 = matrix(results), names nformat(number_d2)

*** Urine test *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_urine if preg_comp==1 & select & baseline_status==1 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel H212 = "Urine test at ANC", bold overwritefmt
putexcel L213 = "N = ", right bold overwritefmt
putexcel M213 = `e(N)', hcenter bold overwritefmt
putexcel H214 = matrix(results), names nformat(number_d2)


*==================================================================*
* Among women with PE, exclusing 5-9 weeks postpartum women * 
*==================================================================*

*** ANC Frequency *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.anc_num_cat if preg_comp==1 & select & baseline_status!=3 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O1 = "Adjusted logistic regression results among with PE, exclusing 5-9 weeks postpartum women", bold overwritefmt
putexcel O3 = "ANC Frequency", bold overwritefmt
putexcel S4 = "N = ", right bold overwritefmt
putexcel T4 = `e(N)', hcenter bold overwritefmt
putexcel O5 = matrix(results), names nformat(number_d2)

*** ANC provider type *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.provider_code if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O23 = "ANC Provider Type", bold overwritefmt
putexcel S24 = "N = ", right bold overwritefmt
putexcel T24 = `e(N)', hcenter bold overwritefmt
putexcel O25 = matrix(results), names nformat(number_d2)

*** Maternal assessment at ANC *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.maternal_assess_cat if preg_comp==1 & select & baseline_status!=3 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O44 = "Maternal assessment (categorial)", bold overwritefmt
putexcel S45 = "N = ", right bold overwritefmt
putexcel T45 = `e(N)', hcenter bold overwritefmt
putexcel O46 = matrix(results), names nformat(number_d2)

*** All maternal assessment *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled maternal_assess_all if preg_comp==1 & select & baseline_status!=3 
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O65 = "Received all 5 maternal assessments", bold overwritefmt
putexcel S66 = "N = ", right bold overwritefmt
putexcel T66 = `e(N)', hcenter bold overwritefmt
putexcel O67 = matrix(results), names nformat(number_d2)

*** PE/E danger sign counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled danger_sign_coun if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O83 = "PE/E danger sign counseling", bold overwritefmt
putexcel S84 = "N = ", right bold overwritefmt
putexcel T84 = `e(N)', hcenter bold overwritefmt
putexcel O85 = matrix(results), names nformat(number_d2)

*** Birth readiness discussion *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled i.birth_readiness_cat if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O100 = "Birth readiness discussion (categorical)", bold overwritefmt
putexcel S101 = "N = ", right bold overwritefmt
putexcel T101 = `e(N)', hcenter bold overwritefmt
putexcel O102 = matrix(results), names nformat(number_d2)


svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled birth_readiness_all if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O122 = "Received all 9 birth readiness discussions", bold overwritefmt
putexcel S123 = "N = ", right bold overwritefmt
putexcel T123 = `e(N)', hcenter bold overwritefmt
putexcel O124 = matrix(results), names nformat(number_d2)

*** Nutritional counseling *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_nd_info_yn if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O140 = "Nutritional counseling", bold overwritefmt
putexcel S141 = "N = ", right bold overwritefmt
putexcel T141 = `e(N)', hcenter bold overwritefmt
putexcel O142= matrix(results), names nformat(number_d2)

*** Tetanus injection *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_tt_inject if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O158 = "Tetanus injection", bold overwritefmt
putexcel S159 = "N = ", right bold overwritefmt
putexcel T159 = `e(N)', hcenter bold overwritefmt
putexcel O160 = matrix(results), names nformat(number_d2)


*** BP measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_bp if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O176 = "Blood pressure measurement at ANC", bold overwritefmt
putexcel S177 = "N = ", right bold overwritefmt
putexcel T177 = `e(N)', hcenter bold overwritefmt
putexcel O178 = matrix(results), names nformat(number_d2)

*** Weight measurement *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_weight if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O194 = "Weight measurement at ANC", bold overwritefmt
putexcel S195 = "N = ", right bold overwritefmt
putexcel T195 = `e(N)', hcenter bold overwritefmt
putexcel O196 = matrix(results), names nformat(number_d2)

*** Urine test *** 

svy: logistic SWdelivprob_convuls i.age_cat urban i.parity_cat i.facility_skilled SWanc_urine if preg_comp==1 & select & baseline_status!=3  
mat temp = r(table)'
mat results = temp[1..., "b"], temp[1..., "pvalue"], temp[1..., "ll"], temp[1..., "ul"]
mat list results
matrix colnames results = OR P-value CI-lower CI-upper
putexcel set Regression_output_$date.xlsx, sheet(adjusted-all-PE) modify
putexcel O212 = "Urine test at ANC", bold overwritefmt
putexcel S213 = "N = ", right bold overwritefmt
putexcel T213 = `e(N)', hcenter bold overwritefmt
putexcel O214 = matrix(results), names nformat(number_d2)

