/* PMA Ethiopia 6-Week Merge file
This .do file takes the deidentified recoded 6-week data from PMA Ethiopia and merges it with the deidentified Baseline data */

	clear
	clear matrix
	clear mata
	capture log close
	set maxvar 15000
	set more off
	numlabel, add

*******************************************************************************
* SET MACROS 
*******************************************************************************

*Year Macros
	local COHORT Cohort1 

* Set macros for data sets
	local baselinedata "/Users/Ellie/Dropbox/PMAET2_Datasets/1-Cohort1/1-Baseline/Prelim100/Cohort1_Baseline_WealthWeightAll_3Nov2020.dta"
	local sixweekdata "/Users/Ellie/Dropbox/PMAET2_Datasets/1-Cohort1/2-6Week/95Prelim/Cohort1_NoName_6W_Clean_7Dec2020.dta"
	global datadir "/Users/Ellie/Desktop/THESIS/Data"

*******************************************************************************
***PREPARATION OF DATA
*******************************************************************************
	cd $datadir

*Prepare Baseline for merge
	use "`baselinedata'", clear
	keep if FRS_result!=.

*Create a dummy participant ID for women who did not complete the baseline
	replace participant_id=EA+"_"+member_number if participant_id==""
	duplicates drop participant_id, force

	tempfile base
	save `base'.dta, replace

*Prepare 6-week for Merge
	use "`sixweekdata'" , clear
	foreach var of varlist _all {
		rename `var' SW`var'
		}
	rename SWSWmetainstanceID SWmetainstanceID
	rename SWparticipant_id participant_id
	rename SWSWFUweight SWFUweight

	duplicates drop participant_id, force

	tempfile sw
	save `sw'.dta, replace

*Merge Baseline, Panel Pregnancy and 6 Week
	use `base'.dta, clear

	merge 1:1 participant_id using `sw'.dta, gen(sw_merge)
	drop if sw_merge==2

*Replace dummy participant ID=.
	replace participant_id="" if participant_id==EA+"_"+member_number & FRS_result!=1


	save `COHORT'_6W_Merged_$date.dta, replace

