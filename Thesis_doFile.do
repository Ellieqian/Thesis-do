*=========================================================
* Thesis .do file *
*=========================================================

clear
clear matrix
clear mata
capture log close
set maxvar 15000
set more off
numlabel, add


/*

Outcome: 
	- Convulsions/fits during delivery 


Covariates of interest: 
	- ANC frequency 
		* 0, 1-3, 4+ 
	- ANC components: 
		* Nutrition counseling
		* Key services received (Had BP, urine and blood sampled, tested for syphilis/HIV, + took iron)
		* Birth readiness discussion 
		
Things to adjust for: 
	- Background characteristics
		* Age
		* Residence
		* Parity
		* Education 
		* Wealth? 
		* Marital status? 
	- Pregnancy intention
	- Facility delivery and/or skilled birth attendant
	- Complications experienced during pregnancy 
	
	
2,853 interviews completed and consented to follow up 
2,578 6-week interviews expected 

*/


*=========================================================
* Set up macros and clean dataset *
*=========================================================

* Change directory and make sub-folder
	global datadir "/Users/Ellie/Desktop/THESIS/Data"
	global dofiledir "/Users/Ellie/Desktop/THESIS/Thesis-do"
	global output "/Users/Ellie/Desktop/THESIS/Output"
	
*Year Macros
	local COHORT Cohort1 

* Set local/global macros for current date
	local today=c(current_date)
	local c_today= "`today'"
	global date=subinstr("`c_today'", " ", "",.)
	
* Run .do file to merge baseline and 6-week data 
	run "$dofiledir/6Week_Merge.do"

* Load data
	use "$datadir/`COHORT'_6W_Merged_$date.dta"
	
* Create log file
	cd $output
	capture log close
	log using "Thesis.log", replace

* Generate strata
	drop strata
	egen strata=concat(region ur) if region==1 | region==3 | region==4 | region==7, punc(-)
	replace strata=string(region) if strata==""

* Create weight for all women with complete forms
	gen SWweight=FQweight if baseline_status==3
	replace SWweight=SWFUweight if SWresult!=.

* 6-week response rate among pregnant and 0-4 weeks postpartum women
	gen sw_responserate=0 if SWresult>=1 & SWresult<6 & SWresult!=. & baseline_status!=3
	replace sw_responserate=1 if SWresult==1 & baseline_status!=3
	label define responselist 0 "Not complete" 1 "Complete"
	label val sw_responserate responselist

	tabout sw_responserate if baseline_status!=3 using "Thesis_output_$date.xls", replace cells(freq col) h2("6W response rate") f(0 1) clab(n %)

* Baseline response rate among pregnant and 5-9 weeks postpartum women
	gen responserate=0 if FRS_>=1 & FRS_<6 & FRS_!=. & baseline_status==3
	replace responserate=1 if FRS_result==1 & baseline_status==3
	label val responserate responselist
	
	tabout responserate if baseline_status==3 using "Thesis_output_$date.xls", replace cells(freq col) h2("Basline response rate") f(0 1) clab(n %)

* Regional breakdown of the women
	tabout region if FRS_result==1 | SWresult==1 using "Thesis_output_$date.xls", append cells(freq col) h2("Number of Women by Region") f(0 1) clab(n %) 

* Source of the data 
	gen source=1 if FRS_result!=.
	replace source=2 if SWresult_orig!=.
	replace source=3 if SWresult_cc!=.
	label var source "Source of 6-week data"
	label define source_list 1 "Baseline (5-9 wks pp at baseline)" 2 "6-Week pre-COVID" 3 "6-Week during COVID"
	label val source source_list

	tabout source if FRS_result==1 | SWresult==1 using "Thesis_output_$date.xls", append cells(freq col) h2("Source of 6-Week Data-Complete forms") f(0 1) clab(n %) 

* Restrict analysis to women who completed questionnaire 
	keep if (FRS_result==1 & baseline_status==3) | (SWresult==1 & baseline_status!=3)

* Set survey weights
	svyset EA [pweight=SWweight], strata(strata) singleunit(scaled)

* Generate 0/1 urban/rural variable
	gen urban=ur==1
	label variable urban "Urban/rural place of residence"
	label define urban 1 "Urban" 0 "Rural"
	label value urban urban

*=========================================================
* Covariates set up *
*=========================================================

* Replace -99 and -88 to missing
	foreach var in total_birth pregnancy_desired marital_status SWanc_hew_timing SWanc_phcp_timing ///
					SWanc_hew_yn SWanc_phcp_yn SWanc_hew_yn SWanc_phcp_yn ///
					SWanc_hew_num SWanc_hew_num SWanc_phcp_num SWanc_phcp_num ///
					SWanc_delivery_place SWanc_delivery_skilled SWanc_emergency_transport ///
					SWanc_danger_place SWanc_danger_migraine SWanc_danger_hbp SWanc_danger_edema ///
					SWanc_danger_convuls SWanc_danger_bleeding SWanc_nd_info_nr ///
					SWanc_bp SWanc_weight SWanc_urine SWanc_blood SWanc_stool SWanc_syph_test ///
					SWanc_syph_result SWanc_syph_couns SWanc_hiv_test SWanc_hiv_result SWanc_hiv_couns ///
					SWanc_lam_couns SWanc_ppfp_couns SWwho_assisted_delivery ///
					SWdelivprob_bleed SWdelivprob_leakmemb24hr SWdelivprob_leakmembpre9mo ///
					SWdelivprob_malposition SWdelivprob_prolonglab SWdelivprob_convuls ///
					SWpostdelivprob_retainpl SWpostdelivprob_fever SWpostdelivprob_bleed ///
					SWpostdelivprob_convuls SWpregprob_hbp SWpregprob_edema SWpregprob_convuls ///
					SWpregprob_vagbleed SWpregprob_fever SWpregprob_abnormdisch SWpregprob_abpain SWpregprob_vision {
		replace `var'=. if `var'==-99 | `var'==-88 
		}
		
	foreach var in anc_delivery_place anc_delivery_skilled anc_emergency_transport ///
					anc_danger_place anc_danger_migraine anc_danger_hbp anc_danger_edema ///
					anc_danger_convuls anc_danger_bleeding anc_nd_info_nr ///
					anc_bp anc_weight anc_urine anc_blood anc_stool anc_syph_test ///
					anc_syph_result anc_syph_couns anc_hiv_test anc_hiv_result anc_hiv_couns ///
					anc_lam_couns anc_ppfp_couns who_assisted_delivery ///
					delivprob_bleed delivprob_leakmemb24hr delivprob_leakmembpre9mo ///
					delivprob_malposition delivprob_prolonglab delivprob_convuls ///
					postdelivprob_retainpl postdelivprob_fever postdelivprob_bleed ///
					postdelivprob_convuls pregprob_hbp pregprob_edema pregprob_convuls ///
					pregprob_vagbleed pregprob_fever pregprob_abnormdisch pregprob_abpain pregprob_vision {
		replace SW`var'= `var' if baseline_status==3
		}
		
	replace SWanc_hew_yn=anc_hew_yn_pp if baseline_status==3
	replace SWanc_phcp_yn=anc_phcp_yn_pp if baseline_status==3
	replace SWanc_hew_num=anc_hew_num_pp if baseline_status==3
	replace SWanc_phcp_num=anc_phcp_num_pp if baseline_status==3

* Label yes/no response options
	capture label define yesno 0 "No" 1 "Yes"	

* Generate age categories
	gen age_cat=0
	replace age_cat=1 if FQ_age>=20 & FQ_age<35
	replace age_cat=2 if FQ_age>=35
	label define agel 0 "<20" 1 "20-34" 2 "35+"
	label value age_cat agel

	
* Generate categorical variable for parity 
	egen parity4=cut(total_births), at(0, 1, 2, 4, 30) icodes
	replace parity4=0 if total_births==.
	lab def parity4l 0 "0 children" 1 "1 children" 2 "2-3 children" 3 "4+ children"
	lab val parity4 parity4l
	lab var parity4 "Parity" 

* Group secondary, technical & vocational, and higher education 
	gen education=school
	replace education=2 if school>=3
	lab def edul 0 "No education" 1 "Primary" 2 "Secondary+" 
	lab val education edul
	lab var education "Education level" 
	
	gen any_edu=0 if education==0
	replace any_edu=1 if education==1 | education==2

* Generate unintended pregnancy binary variable
	gen unintended_preg=0
	replace unintended_preg=1 if pregnancy_desired==2 | pregnancy_desired==3 
	lab var unintended_preg "Unintended pregnancy" 
	label val unintended_preg yesno

* Generate binary marital status variable 
	gen married=0 if marital_status!=.
	replace married=1 if marital_status==1 | marital_status==2
	label define marriedl 0 "Not married" 1 "Married or living with a partner"
	label val married marriedl
	
* Facility versus home delivery
	gen facility_deliv=0 if SWdelivery_place==1 | SWdelivery_place==2
	replace facility_deliv=1 if SWdelivery_place>2 & SWdelivery_place!=96
	label val facility_deliv yesno
	
* Skilled birth attendant binary variable
	gen skilled_birth=0 if SWwho_assisted_delivery!=.
	replace skilled_birth=1 if (SWwho_assisted_delivery==1 | SWwho_assisted_delivery==2 | ///
								SWwho_assisted_delivery==3 | SWwho_assisted_delivery==4 | ///
								SWwho_assisted_delivery==5 | SWwho_assisted_delivery==6 | ///
								SWwho_assisted_delivery==7)
	label val skilled_birth yesno
								
*=========================================================
* ANC-related covariste *
*=========================================================

* ANC frequency 
	* Total number of ANC 
	replace SWanc_hew_num=0 if SWanc_hew_num==.
	replace SWanc_phcp_num=0 if SWanc_phcp_num==.
	gen anc_tot=SWanc_hew_num + SWanc_phcp_num 
	label var anc_tot "Total number of ANC" 
	
	* 0, 1-3 and 4+ ANC categories 
	gen anc_num_cat=0 if anc_tot==0
	replace anc_num_cat=1 if anc_tot>=1 & anc_tot<=3
	replace anc_num_cat=2 if anc_tot>=4
	label define anc_num_l 0 "No ANC" 1 "1-3 ANC" 2 "4+ ANC"
	label val anc_num_cat anc_num_l
	
	* 4+ ANC
	gen ANC4=0
	replace ANC4=1 if anc_num_cat==2
	label var ANC4 "Had mroe than 4 ANC" 
	
	* Any ANC
	gen anyanc=0 if anc_tot==0
	replace anyanc=1 if anc_num_cat==1 | anc_num_cat==2
	label var anyanc "Had at least one ANC" 
	
	* 8+ ANC
	gen anc_num_cat1=0 if anc_tot==0
	replace anc_num_cat1=1 if anc_tot>=1 & anc_tot<=3
	replace anc_num_cat1=2 if anc_tot>=4 & anc_tot<=5
	replace anc_num_cat1=3 if anc_tot>=6

	label define anc_num_lab 0 "No ANC" 1 "1-3 ANC" 2 "4-5 ANC" 3 "6+ANC" 
	label val anc_num_cat1 anc_num_lab

* GA at first ANC visit

	* Generate GA at first visit variable
	gen ga_first_anc=SWanc_hew_timing if SWanc_hew_timing!=. & anyanc==1
	replace ga_first_anc=SWanc_phcp_timing if SWanc_phcp_timing<=anc_hew_timing & SWanc_phcp_timing!=. 
	label var ga_first_anc "GA (in months) at first ANC"
	
	* Create categories 
	gen ga_first_anc_cat=0 if ga_first_anc<4
	replace ga_first_anc_cat=1 if ga_first_anc>=4 & ga_first_anc<=6
	replace ga_first_anc_cat=2 if ga_first_anc>6 & ga_first_anc<=9
	label define ga_anc_l 0 "First trimester" 1 "Second trimester" 2 "Third trimester"
	label val ga_first_anc_cat ga_anc_l
	
* ANC by provider type
	
	/* Generate provider code 
			0: No anc 
			1: HEW only 
			2. PHCP only
			3. Both
	*/
			
	gen provider_code=0 if anyanc==0
	replace provider_code=1 if SWanc_hew_yn==1 & SWanc_phcp_yn!=1
	replace provider_code=2  if SWanc_phcp_yn==1 & SWanc_hew_yn!=1
	replace provider_code=3 if SWanc_phcp_yn==1 & SWanc_hew_yn==1
	label define providerl 0 "No ANC" 1 "HEW only" 2 "PHCP only" 3 "Both" 
	label val provider_code providerl
	
* Generate binary composite birth/complication readiness indicator
	gen anc_birth_readiness_b=0
	replace anc_birth_readiness_b=1 if (SWanc_delivery_place==1 & SWanc_delivery_skilled==1 & ///
										SWanc_emergency_transport==1 & SWanc_danger_place==1 & ///
										SWanc_danger_migraine==1 & SWanc_danger_hbp==1 & SWanc_danger_edema==1 & ///
										SWanc_danger_convuls==1 & SWanc_danger_bleeding==1)
	label var anc_birth_readiness_b "At ANC discussed all birth/complication readiness topics"
	label val anc_birth_readiness_b yesno
	
	/*
	continuous variable, then categorize  
	look up egen xtile to divide into three even groups
	
	*/

	drop if anyanc==1 & (SWanc_delivery_place==.| SWanc_delivery_skilled==. | SWanc_danger_place==. | ///
		SWanc_emergency_transport==. | SWanc_danger_migraine==. | SWanc_danger_hbp==. | ///
		SWanc_danger_edema==. | SWanc_danger_convuls==. | SWanc_danger_bleeding==.)
		
	gen anc_birth_readiness_c=0 
	replace anc_birth_readiness_c=SWanc_delivery_place + SWanc_delivery_skilled + SWanc_emergency_transport + SWanc_danger_place + ///
									SWanc_danger_migraine + SWanc_danger_hbp + SWanc_danger_edema + SWanc_danger_convuls + SWanc_danger_bleeding
									
	gen birth_readiness_cat=0 if anyanc==1
	replace birth_readiness_cat=1 if anc_birth_readiness_c<=3 & anyanc==1 
	replace birth_readiness_cat=2 if anc_birth_readiness_c>3 & anc_birth_readiness_c<=6 & anyanc==1
	replace birth_readiness_cat=3 if anc_birth_readiness_c>6 & anc_birth_readiness_c<=9 & anyanc==1

	
* PE/E danger sign counseling
	gen danger_sign_coun=0
	replace danger_sign_coun=1 if (SWanc_danger_migraine==1 | anc_danger_hbp==1 | anc_danger_edema==1 | SWanc_danger_convuls==1)

* Key ANC services binary indicators 
	gen anc_key_services=0 
	replace anc_key_services=1 if (SWanc_bp==1 & SWanc_iron==1 & SWanc_blood==1 & ///
									SWanc_urine==1 & SWanc_syph_test==1 & SWanc_hiv_test==1) 
	label var anc_key_services "At ANC had : BP, urine and blood sampled, tested for syphilis/HIV, + took iron during preg"
	label val anc_key_services yesno
	tab anc_key_services

* PE/E-related Pregnancy complications binary variable 
	gen preg_comp=0
	replace preg_comp=1 if (SWpregprob_hbp==1 | SWpregprob_edema==1 | SWpregprob_convuls==1 | SWpregprob_vision==1) 
	label var preg_comp "Had PE/E-related complications during pregnancy" 
	label val preg_comp yesno

* Facility delivery and skilled birth attendant combined variable
	gen facility_skilled=0 if facility_deliv!=. & skilled_birth!=.
	replace facility_skilled=1 if (facility_deliv==0 & skilled_birth==0)
	replace facility_skilled=2 if (facility_deliv==0 & skilled_birth==1)
	replace facility_skilled=3 if facility_deliv==1
	tab facility_skilled
	label define facility_skilledl 1 "Home delivery with no SBA" 2 "Home delivery with SBA" 3 "Facility delivery with SBA"
	label val facility_skilled facility_skilledl

*=========================================================
* EDA *
*=========================================================

* Tabulation of outcome and potential confounders
* Background characteristics
	tabout SWdelivprob_convuls age_cat [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls education [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls urban [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls married [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls parity4 [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls unintended_preg [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 


* ANC and delivery care 
	tabout SWdelivprob_convuls anyanc [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls ANC4 [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls anc_key_services [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls anc_birth_readiness_b [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls preg_comp [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls facility_deliv [aw=SWweight] using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls skilled_birth [aw=SWweight]  using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 
	tabout SWdelivprob_convuls SWcaesarean_delivery [aw=SWweight]  using "Thesis_output_$date.xls", append cells(freq col) f(0 1) npos(col) clab(n %) 


* Lowess smoothed regressions -- continuous Xs
/*
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
	for var parity4:  tabulate X SWdelivprob_convuls, chi2 exact column
	for var education:  tabulate X SWdelivprob_convuls, chi2 exact column
	for var anc_num_cat:  tabulate X SWdelivprob_convuls, chi2 exact column


graph box FQ_age, name(age_provider,replace) over(provider_code, relabel(1 "No ANC" 2 "HEW only" 3 "PHCP only" 4 "Both")) ///
		marker(1,mlab(FQ_age)) t1("Age by provider type") scale(1.2) ytitle("Age (years)") 
		graph export graphs/age_provider.png, replace

graph box total_birth, name(parity_provider,replace) over(provider_code, relabel(1 "No ANC" 2 "HEW only" 3 "PHCP only" 4 "Both")) ///
		marker(1,mlab(total_birth)) t1("Parity by provider type") scale(1.2) ytitle("Number of children") 
		graph export graphs/parity_provider.png, replace
		
graph box ga_first_anc if provider_code!=0, name(ga_provider,replace) over(provider_code, relabel(1 "HEW only" 2 "PHCP only" 3 "Both")) ///
		marker(1,mlab(ga_first_anc)) t1("GA at first ANC by provider type") scale(1.2) ytitle("Gestational age (months)") 
		graph export graphs/ga_provider.png, replace
	
* Table 1

#delimit;
table1_mc,  			
	by(SWdelivprob_convuls) 			
	vars( 						
	age_cat cat %4.1f \
	education cat %4.1f \
	urban bin %4.1f \
	parity4 cat %4.1f \
	unintended_preg bin %4.1f \
	facility_deliv bin %4.1f \
	skilled_birth bin %4.1f \
	preg_comp bin %4.1f \
	provider_code cat %4.1f \
	) 
	nospace onecol 
	saving(tables/table1_$date.xlsx, replace) // save table 1
;
#delimit cr
*/
*=========================================================
* Logistic regression *
*=========================================================


svy: logistic SWdelivprob_convuls i.anc_num_cat i.facility_skilled i.age_cat urban preg_comp
svy: logistic SWdelivprob_convuls i.anc_num_cat i.facility_skilled i.age_cat urban if preg_comp==1  //unintended_preg married i.education
 
assert 0
svy: logistic SWdelivprob_convuls i.provider_code i.facility_skilled i.age_cat urban preg_comp   //unintended_preg married i.education
svy: logistic SWdelivprob_convuls i.provider_code i.facility_skilled i.age_cat urban if preg_comp==1   //unintended_preg married i.education

	svy: logistic SWdelivprob_convuls i.birth_readiness_cat i.facility_skilled i.provider_code i.age_cat urban if anyanc==1 & preg_comp==1 //unintended_preg married i.education

	svy: logistic SWdelivprob_convuls i.ga_first_anc_cat i.facility_skilled i.age_cat urban preg_comp if anyanc==1  //unintended_preg married i.education

	svy: logistic SWdelivprob_convuls SWanc_nd_info_yn i.facility_skilled i.age_cat urban preg_comp if anyanc==1  //unintended_preg married i.education

* Indicator variable for those who had PE/E danger signs during pregnancy but did not have eclampsia 
	gen pe_no_e=0 if preg_comp==1 & SWdelivprob_convuls==1 // pe & e
	replace pe_no_e=1 if preg_comp==1 & SWdelivprob_convuls==0 // pe but no e

svy: logistic pe_no_e i.provider_code i.facility_skilled  i.age_cat urban 
svy: logistic pe_no_e i.anc_num_cat i.facility_skilled  i.age_cat urban 

// compared to those who had pe & e, those who had pe but no e have a 2.8-fold increase in the odds of seeking care from both a HEW and PHCP //

/*
	NOTE: 
		Among those who had pre-eclampsia danger signs but did not have eclampsia,
		a much higher proportion sought care 

*/

log close
