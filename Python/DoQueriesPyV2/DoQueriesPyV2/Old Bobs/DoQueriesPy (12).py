﻿__author__ = 'mplace'
#Version 3.02

import os
import time
import shutil

# date becomes the current date and is then placed in MM-DD-YY format
date = time.strftime("%x").replace("/", "-")
month_folder = date[:2] + "-20" + date[-2:]
###############################
test = False
###############################


def rename(name, new_name, attach_list, i=2):
    this_name = os.path.realpath(os.path.join(os.listdir("."), name))
    this_new_name = os.path.realpath(os.path.join(os.listdir("."), new_name))
    this_attach_list = attach_list
    num = i
    try:
        os.rename(this_name, this_new_name)
        this_attach_list.append(this_new_name)
    except WindowsError:
        try:
            final_name = this_new_name[:-4] + " (" + str(num) + ").xlsx"
            os.rename(this_name, final_name)
            this_attach_list.append(final_name)
        except WindowsError:
            rename(this_name, this_new_name, this_attach_list, num + 1)


def move(name, to_directory):
    move_name = name
    move_directory = to_directory
    try:
        shutil.move(move_name, move_directory)
    except shutil.Error:
        print "Already a file with the name: " + name + " at location."


def mailer(text, subject, recipient, cc, attachments):
    import win32com.client as win32

    list(attachments)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.cc = cc
    mail.Subject = subject
    mail.HtmlBody = text
    for each in attachments:
        mail.Attachments.Add(Source=each)
    mail.Display()


def do_query(name, new_name, destination, attach_list, i=2):
    this_name = name
    this_new_name = new_name
    this_destination = destination
    this_attach_list = attach_list
    num = i
    if num == 2:
        move(this_name, this_destination)
        rename(destination + "/" + this_name, destination + "/" + this_new_name, this_attach_list)

# region Email and Attachment Groups
a_mail = "Anne Maxwell (anne.maxwell@utah.edu)"
ac_mail = "Amber Cook (acook@sa.utah.edu); Carol Bergstrom <cbergstrom@sa.utah.edu>"
aca_mail = "Amber Cook (acook@sa.utah.edu); Carol Bergstrom <cbergstrom@sa.utah.edu>; Amy Capps (acapps@sa.utah.edu)"
acakr_mail = "Amber Cook (acook@sa.utah.edu); Carol Bergstrom <cbergstrom@sa.utah.edu>;Amy Capps (acapps@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu);Ryan Christensen (rchristensen@sa.utah.edu)"
act_mail = "Amber Cook (acook@sa.utah.edu); Carol Bergstrom <cbergstrom@sa.utah.edu>;"
acvj_mail = "Amber Cook (acook@sa.utah.edu); Carol Bergstrom <cbergstrom@sa.utah.edu>;Veronica Christensen (vchristensen@sa.utah.edu); Jennifer Berry (jberry@sa.utah.edu)"
ak_mail = "acapps@sa.utah.edu; karen.henriquez@utah.edu"
aka_mail = "Amy Capps (acapps@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu);Anne Maxwell (anne.maxwell@utah.edu)"
akar_mail = "Amy Capps (acapps@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu);Anne Maxwell (anne.maxwell@utah.edu); Ryan Christensen (rchristensen@sa.utah.edu)"
akc_mail = "acapps@sa.utah.edu; Karen.Henriquez@utah.edu; cspringer@sa.utah.edu"
akk_mail = "Amy Capps (acapps@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu);Krista Burton (kburton@sa.utah.edu)"
akr_mail = "acapps@sa.utah.edu; karen.henriquez@utah.edu; rchristensen@sa.utah.edu"
akl_mail = "Amy Capps (acapps@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu); Linh Ly <lly@sa.utah.edu>"
akrc_mail = "acapps@sa.utah.edu; karen.henriquez@utah.edu; rchristensen@sa.utah.edu; cspringer@sa.utah.edu"
akrk_mail = "Amy Capps (acapps@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu);Ryan Christensen (rchristensen@sa.utah.edu); Krista Burton (kburton@sa.utah.edu)"
akrv_mail = "Amy Capps (acapps@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu);Ryan Christensen (rchristensen@sa.utah.edu); Veronica Christensen (vchristensen@sa.utah.edu)"
akv_mail = "Amy Capps (acapps@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu);Veronica Christensen (vchristensen@sa.utah.edu)"
alt_mail = "Krista Burton <kburton@sa.utah.edu>; Amy Capps <acapps@sa.utah.edu>; Jennifer Berry <jberry@sa.utah.edu>;Mathew Edward Place <mplace@sa.utah.edu>; Scott Wilgar <swilgar@sa.utah.edu>;Veronica Christensen <vchristensen@sa.utah.edu>;LGaray@sa.utah.edu"
atcj_mail = "Amber Cook (acook@sa.utah.edu); Carol Bergstrom <cbergstrom@sa.utah.edu>; Jonathan Leon <JLeon@sa.utah.edu>"
ca_mail = "Amy Capps (acapps@sa.utah.edu); cspringer@sa.utah.edu)"
cak_mail = "Carol Bergstrom <cbergstrom@sa.utah.edu>; Amy Capps (acapps@sa.utah.edu);Karen Henriquez (Karen.Henriquez@utah.edu)"
disb_mail = "Karen Henriquez (Karen.Henriquez@utah.edu); Steffany Forrest (steffany.forrest@income.utah.edu);Leila Ames (lames@sa.utah.edu); KAYLA JOY MC CLOYN <kmccloyn@sa.utah.edu>;Lisa Zaelit (lisa.zaelit@admin.utah.edu);Veronica Christensen (vchristensen@sa.utah.edu); Jennifer Berry (jberry@sa.utah.edu);Carol Bergstrom <cbergstrom@sa.utah.edu>; Amy Capps (acapps@sa.utah.edu);Krista Burton (kburton@sa.utah.edu)"
dl_mail = "Krista Burton <kburton@sa.utah.edu>; Amy Capps <acapps@sa.utah.edu>;Karen Henriquez <Karen.Henriquez@utah.edu>;Jennifer Berry <jberry@sa.utah.edu>; Mathew Edward Place <mplace@sa.utah.edu>;Scott Wilgar <swilgar@sa.utah.edu>; Veronica Christensen <vchristensen@sa.utah.edu>;LGaray@sa.utah.edu"
jen_mail = "Jennifer Berry <jberry@sa.utah.edu>"
ka_mail = "Krista Burton (kburton@sa.utah.edu); Amy Capps (acapps@sa.utah.edu)"
kaca_mail = "Karen Henriquez (Karen.Henriquez@utah.edu); Amber Cook (acook@sa.utah.edu);Amy Capps (acapps@sa.utah.edu)"
kak_mail = "Krista Burton (kburton@sa.utah.edu); Amy Capps (acapps@sa.utah.edu);Karen Henriquez (Karen.Henriquez@utah.edu)"
kc_mail = "Karen.Henriquez@utah.edu; cspringer@sa.utah.edu"
lkjk_mail = "Leila Ames (lames@sa.utah.edu); John Curl (jcurl@sa.utah.edu); Karen Henriquez (Karen.Henriquez@utah.edu); KAYLA JOY MC CLOYN <kmccloyn@sa.utah.edu>"
lkj_mail = "Leila Ames (lames@sa.utah.edu);KAYLA JOY MC CLOYN (kmccloyn@sa.utah.edu); John Curl (jcurl@sa.utah.edu)"
mat_mail = "mplace@sa.utah.edu"
ms_mail = "mplace@sa.utah.edu; Scott Wilgar (swilgar@sa.utah.edu)"
null_mail = ""
rac_mail = "Amber Cook (acook@sa.utah.edu); Carol Bergstrom <cbergstrom@sa.utah.edu>; Raenetta King (rking@sa.utah.edu)n"
scott_mail = "Scott Wilgar (swilgar@sa.utah.edu)"
ss_mail = "Amber Cook <acook@sa.utah.edu>; Carol Bergstrom <cbergstrom@sa.utah.edu>;Cary Lopez <cary.lopez@utah.edu>;Jonathan Leon <JLeon@sa.utah.edu>; MARY SNOW <MSnow@sa.utah.edu>; Sheryl Hansen <shansen@sa.utah.edu>; Jennifer Berry <jberry@sa.utah.edu>;Leonel Garay <LGaray@sa.utah.edu>; Mathew Place <mplace@sa.utah.edu>; Scott Wilgar <swilgar@sa.utah.edu>;Veronica Christensen <vchristensen@sa.utah.edu>"
ssj_mail = "Carol Bergstrom <cbergstrom@sa.utah.edu>; Amber Cook (acook@sa.utah.edu); Leila Ames (lames@sa.utah.edu); Brenda Burke <BBurke@sa.utah.edu>; Jonathan Leon <JLeon@sa.utah.edu>; MARY SNOW <MSnow@sa.utah.edu>; HILERIE A HARRIS <hilerie.harris@sa.utah.edu>; Veronica Christensen (vchristensen@sa.utah.edu); Scott Wilgar (swilgar@sa.utah.edu); Leonel Garay <LGaray@sa.utah.edu>; Mathew Place <mplace@sa.utah.edu>; Jennifer Berry (jberry@sa.utah.edu)"
sys_mail = "Jennifer Berry <jberry@sa.utah.edu>; Leonel Garay <LGaray@sa.utah.edu>; Mathew Place <mplace@sa.utah.edu>;Scott Wilgar <swilgar@sa.utah.edu>; Veronica Christensen <vchristensen@sa.utah.edu>"
v_mail = "Veronica Christensen (vchristensen@sa.utah.edu)"
vm_mail = "Veronica Christensen (vchristensen@sa.utah.edu); Mathew Place <mplace@sa.utah.edu>"
vs_mail = "Veronica Christensen (vchristensen@sa.utah.edu);Scott Wilgar (swilgar@sa.utah.edu)"

a_attachment_list = []
ac_attachment_list = []
aca_attachment_list = []
acakr_attachment_list = []
act_attachment_list = []
acvj_attachment_list = []
ak_attachment_list = []
aka_attachment_list = []
akar_attachment_list = []
akc_attachment_list = []
akk_attachment_list = []
akr_attachment_list = []
akl_attachment_list = []
akrk_attachment_list = []
akrv_attachment_list = []
akv_attachment_list = []
alt_attachment_list = []
atcj_attachment_list = []
cak_attachment_list = []
disb_attachment_list = []
dl_attachment_list = []
jen_attachment_list = []
ka_attachment_list = []
kaca_attachment_list = []
kak_attachment_list = []
kc_attachment_list = []
lkj_attachment_list = []
mat_attachment_list = []
ms_attachment_list = []
null_attachment_list = []
rac_attachment_list = []
scott_attachment_list = []
ss_attachment_list = []
ssj_attachment_list = []
sys_attachment_list = []
v_attachment_list = []
vm_attachment_list = []

# endregion


def do_dailies():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_IL_CMT_CODE_OVR_AGR_") \
                | query_name.startswith("UUFA_IL_CMT_CDE_OVR_AGR"):
            if date[6:] in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
            if str(int(date[6:]) + 1) in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break

    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Daily', aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Daily', aid_year, month_folder))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_IL_ATHLETE_OVERAWARD_"):
            do_query(query, date + " Athlete Aid Overaward " + year + ".xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_IL_CMT_CODE_OVR_AGR_LMT_" + year) | \
                query.startswith("UUFA_IL_CMT_CDE_OVR_AGR_LMT_" + year):
            do_query(query, date + " Comment Code Over Aggregate 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_COMMENT_CODE_298_" + year) | \
                query.startswith("UUFA_IL_COMMENT_CODE_298_" + year):
            do_query(query, date + " IASG - Pell Eligible 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FDEG_FFBD_FBLK_FFBC_" + year) | \
                query.startswith("UUFA_IL_FDEG_FFBD_FBLK_FFBC_" + year):
            do_query(query, date + " Complete FDEG FFBD FBLK FFBC " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_COMPLETE_FDEG_" + year):
            do_query(query, date + " FDEG Update 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_CORR_NOT_MARKED_SENT_" + year) | \
                query.startswith("UUFA_IL_CORR_NOT_MARK_SENT_" + year):
            do_query(query, date + " Corrections not Marked to Sent 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_CORR_SENT_REJECT_CD1_" + year) | \
                query.startswith("UUFA_IL_CORR_SENT_RJCT_CD1_" + year):
            do_query(query, date + " Correction Sent Reject Code 1 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_ENRL_GR_DATE_ERRORS_" + year):
            do_query(query, date + " Place FDIP" + year + " Checklist 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FBKP" + year + "_CHECKLIST_" + year) | \
                query.startswith("UUFA_IL_FBKP16_CHECKLIST_" + year):
            do_query(query, date + " FBKP" + year + " Checklist Initiated.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FOUT" + year + "_INC_" + year) | \
                query.startswith("UUFA_IL_FOUT" + year + "_INC_" + year):
            do_query(query, date + " Outside Resources 20" + year + ".xlsx", directory,
                     rac_attachment_list)
        if query.startswith("FA_IL_FP1B" + year + "_CHECKLIST_" + year):
            do_query(query, date + " FP1B" + year + " Checklist " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FP2B" + year + "_CHECKLIST_" + year):
            do_query(query, date + " FP2B" + year + " Checklist " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FP1N" + year + "_CHECKLIST_" + year):
            do_query(query, date + " FP1N" + year + " Checklist " + year + ".xlsx", directory,
                     rac_attachment_list)
        if query.startswith("FA_IL_FP2N" + year + "_CHECKLIST_" + year):
            do_query(query, date + " FP2N" + year + " Checklist " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FPJ" + year + "_CHECKLIST_" + year) | \
                query.startswith("UUFA_IL_FPJ" + year + "_CHECKLIST_" + year):
            do_query(query, date + " FPJ" + year + " Checklist 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_ISIR_02_IND_UP_DOWN_" + year) | \
                query.startswith("UUFA_IL_ISIR_02_IND_UP_DWN_" + year):
            do_query(query, date + " ISIR Service IND UP Down 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FEHU" + year + "_INITIATED") | \
                query.startswith("UUFA_IL_FEHU" + year + "_INITIATED"):
            do_query(query, date + " Initiated FEHU" + year + " Checklist.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_IL_IRS_DRT_02_20" + year) | \
                query.startswith("UUFA_IL_IRS_DRT_02_20" + year):
            do_query(query, date + " IRS Data Retrieval Equal to 02 " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_ISIR_CMT_CODE_359_360_" + year):
            do_query(query, date + " ISIR Comment Code 359 or 360 " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_IL_ISIR_GRAD_I_UG_FATERM_" + year) | \
                query.startswith("UUFA_IL_ISIR_GRD_I_UG_FATRM_" + year):
            do_query(query, date + " ISIR Graduate Independent UG FATERM 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_ISIR_PRMARY_EFC_DIF_20" + year) | \
                query.startswith("UUFA_IL_ISIR_PRMARY_EFC_DIF_" + year):
            do_query(query, date + " Primary EFC Difference 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_NURSING_LOANS_TILA_20" + year) | \
                query.startswith("UUFA_IL_NURSING_LOANS_TILA_" + year):
            do_query(query, date + " Nursing Loans 20" + year + ".xlsx", directory,
                     akc_attachment_list)
        if query.startswith("FA_IL_OTHER_ATB_20" + year):
            do_query(query, date + " ISIR Other ATB 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_OTHER_ATTND_" + year):
            do_query(query, date + " Attend Other Institution 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_PELL_LEU_C_20" + year) | \
                query.startswith("UUFA_IL_PELL_LEU_C_" + year):
            do_query(query, date + " Pell LEU Limit Flag C " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_PELL_LEU_E_20" + year) | \
                query.startswith("UUFA_IL_PELL_LEU_E_" + year):
            do_query(query, date + " Pell LEU Limit Flag E " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_PELL_MAX_ELIGIBILITY_" + year) | \
                query.startswith("UUFA_IL_PELL_MAX_ELIG_" + year):
            do_query(query, date + " Pell Max Eligibility " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_IL_SF_RFND_AWD_NO_POST_20" + year) | \
                query.startswith("UUFA_IL_SF_RFND_AWD_NO_POST_" + year):
            do_query(query, date + " Refund Post Third Party 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_SUB_ISIR_NO_PACKAGE_" + year) | \
                query.startswith("UUFA_IL_SUB_ISIR_NO_PACKAGE_" + year):
            do_query(query, date + " Subsequent ISIR Not Package Not Verified 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_VET_ACTIVE_DUTY_STAT_" + year) | \
                query.startswith("UUFA_IL_VET_ACTV_DUTY_STAT_" + year):
            do_query(query, date + " Veteran Active Duty Status 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_IL_VER_I_SUB_SUSP_ISIR_" + year):
            do_query(query, date + " FAVR Initiated Susp ISIR Psbl DRT " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_UPDATED_ATB_20" + year) | \
                query.startswith("UUFA_IL_UPDATED_ATB_" + year):
            do_query(query, date + " New ISIR Updated ATB " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FATN_INITIATED_20" + year) | \
                query.startswith("UUFA_IL_FATN_INITIATED_" + year):
            do_query(query, date + " Review FATN Checklist 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_ATHLETE_OVERAWARD_" + year):
            do_query(query, date + " Athlete Aid Overaward " + year + ".xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_IL_FED_AID_OVERAWARD_" + year) | \
                query.startswith("UUFA_IL_FED_AID_OVERAWARD_" + year):
            do_query(query, date + " Federal Aid Overaward " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_IL_FHST_I_HST_COMPLETE_" + year) \
                | query.startswith("UUFA_IL_FHST_I_HST_COMPLETE_" + year):
            do_query(query, date + " HS Transcript 'C' FHST" + year + " I.xlsx", directory,
                     akr_attachment_list)
        if year == "15" and query.startswith("FA_IL_SW_THESIS_HOURS"):
            do_query(query, date + " SW Thesis Hours.xlsx", directory,
                     ak_attachment_list)
        if year == "16" and query.startswith("ussf0034"):
            do_query(query, date + " " + query, directory,
                     acakr_attachment_list)
        if query.startswith("UUFA_IL_PKG_SCH_EXP_GRAD_FA_" + year):
            do_query(query, date + " Scholarship Aid Grad Date Fall 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_IL_PKG_FED_EXP_GRAD_FA_" + year):
            do_query(query, date + " Accepted Federal Aid Grad Date Fall 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_IL_FPEL" + year + "_NO_DB_MATCH"):
            do_query(query, date + " FPEL" + year + " No Database Match.xlsx", directory,
                     akr_attachment_list)

    if ak_attachment_list:
        mailer("", aid_year + " Daily Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if akc_attachment_list:
        mailer("", aid_year + " Daily Queries", akc_mail, "", akc_attachment_list)
        del akc_attachment_list[:]
    if akr_attachment_list:
        mailer("", aid_year + " Daily Queries", akr_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if akr_attachment_list:
        mailer("", aid_year + " Daily Queries", akrc_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if rac_attachment_list:
        mailer("", aid_year + " Daily Queries", rac_mail, "", rac_attachment_list)
        del rac_attachment_list[:]
    if acakr_attachment_list:
        mailer("", aid_year + " Daily Queries", acakr_mail, "", acakr_attachment_list)
        del acakr_attachment_list[:]
    if lkj_attachment_list:
        mailer("", aid_year + " Daily Queries", lkj_mail, "", lkj_attachment_list)
        del lkj_attachment_list[:]


def do_monday_weeklies():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_WR_AID_DISB_NO_ENRLD_ATH_") \
                | query_name.startswith("UUFA_WR_AID_DISB_NO_ENR_ATH_"):
            if date[6:] in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
            if str(int(date[6:]) + 1) in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Monday Weekly', aid_year, month_folder))
        directory2013 = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test\Monday Weekly', "2012-2013", month_folder))
        directory2014 = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test\Monday Weekly', "2013-2014", month_folder))
        packaging_directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Packaging', aid_year, month_folder))
        disb_failure_directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Disb Failure ' + aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monday Weekly', aid_year, month_folder))
        directory2013 = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monday Weekly', "2012-2013", month_folder))
        directory2014 = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monday Weekly', "2013-2014", month_folder))
        packaging_directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Packaging', aid_year, month_folder))
        disb_failure_directory = os.path.realpath(os.path.join('O:/Disbursement Failure/Disb Failure ' + aid_year))

    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(directory2013):
        os.makedirs(directory2013)
    if not os.path.isdir(directory2014):
        os.makedirs(directory2014)
    if not os.path.isdir(packaging_directory):
        os.makedirs(packaging_directory)
    if not os.path.isdir(disb_failure_directory):
        os.makedirs(disb_failure_directory)

    # Change File_Name to be query ac it is received and _new_file_name to what the new query should be.Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("FA_WR_ADM_DEFERRAL_" + year) \
                | query.startswith("UUFA_WR_ADM_DEFERRAL_" + year):
            do_query(query, date + " FA Admission Deferral " + year + ".xlsx", directory,
                     atcj_attachment_list)

        if query.startswith("UUFA_WR_AGG_CK_MLT_YR_AWDED_" + year):
            do_query(query, date + " Student Pkgd for " + str(int(year) - 1) + " after " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_AID_DISB_NO_ENRLD_ATH_" + year) \
                | query.startswith("UUFA_WR_AID_DISB_NO_ENR_ATH_" + year):
            do_query(query, date + " Athlete Disb Not Enrolled " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_WR_AID_DISB_NO_ENRLD_FED_" + year) \
                | query.startswith("UUFA_WR_AID_DISB_NO_ENR_FED_" + year):
            do_query(query, date + " Federal Disb Not Enrolled " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_AID_DISB_NO_ENRLD_SCH_" + year) \
                | query.startswith("UUFA_WR_AID_DISB_NO_ENR_SCH_" + year):
            do_query(query, date + " T 53 Sch Disb Not Enrolled " + year + ".xlsx", directory,
                     act_attachment_list)

        if query.startswith("FA_WR_AMERICORP_AWD_POSTING_" + year) \
                | query.startswith("UUFA_WR_AMERICORP_AWD_POST_" + year):
            do_query(query, date + " Americorp Awards " + year + ".xlsx", directory,
                     act_attachment_list)

        if query.startswith("FA_WR_ATHLETE_NOT_DISB_" + year) \
                | query.startswith("UUFA_WR_ATHLETE_NOT_DISB_" + year):
            do_query(query, date + " Athlete Not Disbursed " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_WR_ATHLETE_OVERAWARD_" + year):
            do_query(query, date + " Athlete Aid Overaward " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_WR_ATH_HRS_AFTER_CENSUS_" + year) \
                | query.startswith("UUFA_WR_ATH_HRS_AFTR_CENSUS_" + year):
            do_query(query, date + " Ath Hours After Census " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_WR_ATH_SF_TERM_BALANCE_" + year) \
                | query.startswith("UUFA_WR_ATH_SF_TERM_BALANCE_" + year):
            do_query(query, date + " Athlete Tuition Fee Balance " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_WR_AUDIT_CLASS_AID_DISB_" + year) \
                | query.startswith("UUFA_WR_AUDIT_CLSS_AID_DISB_" + year):
            do_query(query, date + " Audit Class Aid Disbursed " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_AWARD_UG_NOW_GRAD_ATH_" + year) \
                | query.startswith("UUFA_WR_AWD_UG_NOW_GRAD_ATH_" + year):
            do_query(query, date + " Ath Awards past Grad Term " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_WR_AWRDED_UG_NOW_GRAD_FC_" + year) \
                | query.startswith("UUFA_WR_AWD_UG_NOW_GRAD_FC_" + year):
            do_query(query, date + " Federal Awards past Grad Term " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_AWRDED_UG_NOW_GRAD_SV_" + year) \
                | query.startswith("UUFA_WR_AWD_UG_NOW_GRAD_SV_" + year):
            do_query(query, date + " Scholar Awards past Grad Term " + year + ".xlsx", directory,
                     act_attachment_list)

        if query.startswith("FA_WR_CHECKLST_STATUS_ERROR_" + year) \
                | query.startswith("UUFA_WR_CHKLST_STATUS_ERROR_" + year):
            do_query(query, date + " Checklist Status Error " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_DISB_ATH_FAILURE_" + year) \
                | query.startswith("UUFA_WR_DISB_ATH_FAILURE_" + year):
            do_query(query, date + " Authorization Failure 20" + year + ".xlsx", disb_failure_directory,
                     null_attachment_list)

        if query.startswith("FA_WR_DL_DISBURSED_LTHT_" + year) \
                | query.startswith("UUFA_WR_DL_DISBURSED_LTHT_" + year):
            do_query(query, date + " DL Disbursed LTHT " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_DL_EC_SUSPENDED_" + year) \
                | query.startswith("UUFA_WR_DL_EC_SUSPENDED_" + year):
            do_query(query, date + " DL Entrance Counseling Suspense " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_DL_ORIG_TRANS_PENDING_" + year) \
                | query.startswith("UUFA_WR_DL_ORIG_TRNS_PEND_" + year):
            do_query(query, date + " DL Orig Trans Pending " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_FAFSA_CKLST_INCMP_" + year) \
                | query.startswith("UUFA_WR_FAFSA_CKLST_INCMP_" + year):
            do_query(query, date + " PLUS FAFSA Incomplete " + year + ".xlsx", directory,
                     akr_attachment_list)

        # This query is being phased out after testing on the next query "wdrn" is completed.
        if query.startswith("FA_WR_FALL_TOTAL_WTHDRN_DRP_" + year) \
                | query.startswith("UUFA_WR_FALL_TOTAL_WDRN_DRP_" + year):
            do_query(query, date + " Fall Disb Total Withdrawn Drop " + year + " (old).xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_FALL_TOTAL_WDRN_DRP_" + str(int(year)-1)):
            do_query(query, date + " Fall Disb Total Withdrawn Drop " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_FARC_CHECKLIST_" + year) \
                | query.startswith("UUFA_WR_FARC_CHECKLIST_" + year):
            do_query(query, date + " FARC 30 Day Review " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_FARC_CMNT_CODES_" + year) \
                | query.startswith("UUFA_WR_FARC_CMNT_CODES_" + year):
            do_query(query, date + " Initiated FARC w ISIR Cmnt Codes " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_FED_AID_OVERAWARD_" + year):
            do_query(query, date + " Federal Aid Overaward " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_FGED_ISIR_DEGREE_" + year) \
                | query.startswith("UUFA_WR_FGED_ISIR_DEGREE_" + year):
            do_query(query, date + " FGED ISIR Degree 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_FPEL" + year + "_INITIATED_AWARDED") \
                | query.startswith("UUFA_WR_FPEL" + year + "_INITIATED_AWDED"):
            do_query(query, date + " FPEL" + year + " Initiated Pell.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_GENDER_20" + year) \
                | query.startswith("UUFA_WR_GENDER_" + year):
            do_query(query, date + " Gender Discrepancies 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_HEDU_PARAMEDIC_20" + year) \
                | query.startswith("UUFA_WR_HEDU_PARAMEDIC_" + year):
            do_query(query, date + " HEDU Paramedic Class 20" + year + "F.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_HOME_SCHOOLED_" + year):
            do_query(query, date + " Home Schooled Check " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_HRS_DECREASE_ATH_" + year) \
                | query.startswith("UUFA_WR_HRS_DECREASE_ATH_" + year):
            do_query(query, date + " Hours Decrease Athlete " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_WR_HRS_DECREASE_FC_" + year) \
                | query.startswith("UUFA_WR_HRS_DECREASE_FC_" + year):
            do_query(query, date + " Hours Decrease FC " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_HRS_DECREASE_SV_" + year) \
                | query.startswith("UUFA_WR_HRS_DECREASE_SV_" + year):
            do_query(query, date + " Hours Decrease SV " + year + ".xlsx", directory,
                     act_attachment_list)

        if query.startswith("FA_WR_FHST_I_HST_COMPLETE_" + year) \
                | query.startswith("UUFA_WR_FHST_I_HST_COMPLETE_" + year):
            do_query(query, date + " HS Transcript 'C' FHST" + year + " 'I'.xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_ISIR_AS_EFC_" + year) \
                | query.startswith("UUFA_WR_ISIR_AS_EFC_" + year):
            do_query(query, date + " ISIR Assumption EFC 20" + year + ".xlsx", directory,
                     mat_attachment_list)

        if query.startswith("FA_WR_ISIR_CORR_ASSESSMENT_" + year) \
                | query.startswith("UUFA_WR_ISIR_COR_ASSESSMENT_" + year):
            do_query(query, date + " ISIR Correction Assessment " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_ISIR_CORR_REJECT_" + year) \
                | query.startswith("UUFA_WR_ISIR_CORR_REJECT_" + year):
            do_query(query, date + " ISIR Correction Rejected " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_ISIR_DEGRE_ANSW_CHNGE_" + year) \
                | query.startswith("UUFA_WR_ISIR_DGR_ANSW_CHNG_" + year):
            do_query(query, date + " ISIR Degree Answer Change " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_ISIR_DEPNDCY_STAT_PRB_" + year) \
                | query.startswith("UUFA_WR_ISIR_DEP_STAT_PRB_" + year):
            do_query(query, date + " ISIR Dependency 20" + year + ".xlsx", directory,
                     mat_attachment_list)

        if query.startswith("FA_WR_ISIR_REJECTED_CORR_" + year) \
                | query.startswith("UUFA_WR_ISIR_REJECTED_CORR_" + year):
            do_query(query, date + " ISIR Rejected Corrections 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_ISIR_REJECT_CODES_20" + year) \
                | query.startswith("UUFA_WR_ISIR_REJECT_CODES_" + year) \
                | query.startswith("UUFA_WR_ISIR_REJECT_CODES_20" + year):
            do_query(query, date + " Rejected ISIR's 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_ISIR_SS_MTCH_NOT_CON_" + year) \
                | query.startswith("UUFA_WR_ISIR_SS_MCH_NOT_CON_" + year):
            do_query(query, date + " SS Match Not Confirmed 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_ISIR_SUSPENSE_20" + year) \
                | query.startswith("UUFA_WR_ISIR_SUSPENSE_20" + year):
            do_query(query, date + " ISIR Suspense " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_LEGAL_ALIEN_WORK_" + year) \
                | query.startswith("UUFA_WR_LEGAL_ALIEN_WORK_" + year):
            do_query(query, date + " Legal Alien Work 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_LN_ACCPT_STAFF_31_32_" + year) \
                | query.startswith("UUFA_WR_LN_ACCPT_STAF_31_32_" + year):
            do_query(query, date + " Stafford Accept Offer " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_LOAN_CENSUS_DATE_" + year) \
                | query.startswith("UUFA_WR_LOAN_CENSUS_DATE_" + year):
            do_query(query, date + " Loans Census Date 20" + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_LOAN_FA907_1_REVISED_" + year) \
                | query.startswith("UUFA_WR_LN_FA907_1_REVISED_" + year):
            do_query(query, date + " Loan Disbursed Report " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_LOAN_FA907_2_REVISED_" + year) \
                | query.startswith("UUFA_WR_LN_FA907_2_REVISED_" + year):
            do_query(query, date + " Loan Not Disbursed Report " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_LOAN_ORIG_DEPT_REVIEW_" + year):
            do_query(query, date + " Loan ORIG DEPT RVW " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_LOAN_SENT_NO_RESPONSE_" + year) \
                | query.startswith("UUFA_WR_LN_SENT_NO_RESPONSE_" + year):
            do_query(query, date + " Loan Sent No Response " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_LOAN_TRANSMIT_HOLD_" + year) \
                | query.startswith("UUFA_WR_LOAN_TRANSMIT_HOLD_" + year):
            do_query(query, date + " Loan Transmit Hold " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_LW_MD_DN_AWD_NOT_DISB_" + year) \
                | query.startswith("UUFA_WR_LW_MD_DN_AW_NO_DISB_" + year):
            do_query(query, date + " LW MD DN Awards Not Disbursed " + year + ".xlsx", directory,
                     akc_attachment_list)

        if query.startswith("FA_WR_MONTGMR_AMCORP_OVERAW_" + year) \
                | query.startswith("UUFA_WR_MNTGMR_AMCORP_OVRAW_" + year):
            do_query(query, date + " Montgomery Americorp Overaward " + year + ".xlsx", directory,
                     kaca_attachment_list)

        if query.startswith("FA_WR_MULTIPLE_EMPLIDS_" + year) \
                | query.startswith("UUFA_WR_MULTIPLE_EMPLIDS_" + year):
            do_query(query, date + " Multiple EMPLIDS 20" + year + ".xlsx", directory,
                     mat_attachment_list)

        if query.startswith("FA_WR_NO_COMMENT_CODE_" + year) \
                | query.startswith("UUFA_WR_NO_COMMENT_CODE_" + year):
            do_query(query, date + " Sub ISIR Checklist No ISIR Comment Code 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_NSLDS_LOAN_DATA_" + year) \
                | query.startswith("UUFA_WR_NSLDS_LOAN_DATA_" + year):
            do_query(query, date + " NSLDS Loan Data .xlsx", directory,
                     akk_attachment_list)

        if query.startswith("FA_WR_OVRD_ACAD_LVL_" + year) \
                | query.startswith("UUFA_WR_OVRD_ACAD_LVL_" + year):
            do_query(query, date + " FA Term Override Acad Level " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("UUFA_WR_PA_FDEG_CHECKLIST_" + year):
            do_query(query, date + " PA MPS FDEG Checklist " + year + ".xlsx", directory,
                     akc_attachment_list)

        if query.startswith("FA_WR_PELL_AWRD_LOCK_" + year) \
                | query.startswith("UUFA_WR_PELL_AWRD_LOCK_" + year):
            do_query(query, date + " Pell Award Lock No FPEL" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_PELL_OVERPAYMENT_" + year) \
                | query.startswith("UUFA_WR_PELL_OVERPAYMENT_" + year):
            do_query(query, date + " Pell Ovpy Check NSLDS 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_PELL_SUMMER_NO_PELL_" + year) \
                | query.startswith("UUFA_WR_PELL_SUMMER_NO_PELL_" + year):
            do_query(query, date + " Pell Summer No Pell 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_PELL_TERM_FT_" + year) \
                | query.startswith("UUFA_WR_PELL_TERM_FT_" + year):
            do_query(query, date + " Term Pell Awards FT 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_PELL_TERM_HT_" + year) \
                | query.startswith("UUFA_WR_PELL_TERM_HT_" + year):
            do_query(query, date + " Term Pell Awards HT 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_PELL_TERM_LH_" + year) \
                | query.startswith("UUFA_WR_PELL_TERM_LH_" + year):
            do_query(query, date + " Term Pell Awards LH 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_PELL_TERM_NL_" + year) \
                | query.startswith("UUFA_WR_PELL_TERM_NL_" + year):
            do_query(query, date + " Term Pell Awards NL 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_PELL_TERM_TQ_" + year) \
                | query.startswith("UUFA_WR_PELL_TERM_TQ_" + year):
            do_query(query, date + " Term Pell Awards TQ 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_PERK_SPLIT_MISMATCH_" + year) \
                | query.startswith("UUFA_WR_PERK_SPLIT_MISMATCH_" + year):
            do_query(query, date + " Perkins Plan Split Mismatch " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_QUALITY_ASSURANCE_" + year) \
                | query.startswith("UUFA_WR_QUALITY_ASSURANCE_" + year):
            do_query(query, date + " QA Students Complete Verification 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_RT4_DROPPED_CLASSES_" + year):
            do_query(query, date + " RT4 Dropped Classes 20" + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_SCHOLARSHIP_NOT_DISB_" + year) \
                | query.startswith("UUFA_WR_SCH_NOT_DISB_" + year):
            do_query(query, date + " Cash Non-Cash Sch Not Disb " + year + ".xlsx", directory,
                     act_attachment_list)

        # This query is being phased out after testing on the next query "wdrn" is completed.
        if query.startswith("FA_WR_SPR_TOTAL_WTHDRN_DRP_" + year):
            do_query(query, date + " Spr Disb Total Withdrawn Drop " + year + " (old).xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SPR_TOTAL_WDRN_DRP_" + str(int(year)-1)) \
                | query.startswith("UUFA_WR_SPR_TOTAL_WDRN_DRP_" + year):
            do_query(query, date + " Spr Disb Total Withdrawn Drop " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SSR_MATCH_NOT_CONFRMD_" + year) \
                | query.startswith("UUFA_WR_SSR_MATCH_NOT_CNFRM_" + year):
            do_query(query, date + " SSR Not Confirmed 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SSR_NOT_CONFRMED_VTRN_" + year) \
                | query.startswith("UUFA_WR_SSR_NOT_CNFRMD_VTRN_" + year):
            do_query(query, date + " VA Match SSR DB Override " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SS_DB_OVERRIDE_" + year) \
                | query.startswith("UUFA_WR_SS_DB_OVERRIDE_" + year):
            do_query(query, date + " SS DB Override " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUB_ISIR_PACKAGED_" + year) \
                | query.startswith("UUFA_WR_SUB_ISIR_PACKAGED_" + year):
            do_query(query, date + " Subsequent ISIR Packaged 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUB_ISIR_REAWARD_AID_" + year) \
                | query.startswith("UUFA_WR_SUB_ISIR_REAWD_AID_" + year):
            do_query(query, date + " Canceled FCOR Complete " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUB_ISIR_SYSG_20" + year) \
                | query.startswith("UUFA_WR_SUB_ISIR_SYSG_20" + year):
            do_query(query, date + " Subsequent ISIR System Generated 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUB_ISIR_VERIFIED_" + year) \
                | query.startswith("UUFA_WR_SUB_ISIR_VERIFIED_" + year):
            do_query(query, date + " Subsequent ISIR Verified 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        # This query is being phased out after testing on the next query "wdrn" is completed.
        if query.startswith("FA_WR_SUM_TOTAL_WTHDRN_DRP_" + year):
            do_query(query, date + " Sum Disb Total Withdrawn Drop " + year + " (old).xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUM_TOTAL_WDRN_DRP_" + str(int(year)-1)) \
                | query.startswith("UUFA_WR_SUM_TOTAL_WDRN_DRP_" + year):
            do_query(query, date + " Sum Disb Total Withdrawn Drop " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUMMER_NO_DL_" + year) \
                | query.startswith("UUFA_WR_SUMMER_NO_DL_" + year):
            do_query(query, date + " Summer Enroll No DL " + year + ".xlsx", directory,
                     akk_attachment_list)

        if query.startswith("FA_WR_SUSP_DOB_PRB_APPLCNT_" + year) \
                | query.startswith("UUFA_WR_SSP_DOB_PRB_APPLCNT_" + year):
            do_query(query, date + " Suspense Applicant DOB Problem 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUSP_NAME_PRB_APPLCNT_" + year) \
                | query.startswith("UUFA_WR_SSP_NAME_PRB_APLCNT_" + year):
            do_query(query, date + " Suspense Applicant Name Problem 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("UUFA_WR_SSP_SSN_PRB_APLCNT_" + year) \
                | query.startswith("UUFA_WR_SSP_SSN_PRB_APLCNT_20" + year):
            do_query(query, date + " Suspense Applicant SSN Problem 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_TERM_NSLDS_LOAN_YR_" + year) \
                | query.startswith("UUFA_WR_TERM_NSLDS_LOAN_YR_" + year):
            do_query(query, date + " NSLDS Loan Year Blank " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_TITLE_VII_MED_LOANS_" + year) \
                | query.startswith("UUFA_WR_TITLE_VII_MED_LOANS_" + year):
            do_query(query, date + " Title VII Medical Loans TILA 20" + year + ".xlsx", directory,
                     akc_attachment_list)

        if query.startswith("FA_WR_TRANSFER_ENT_CNS_" + year) \
                | query.startswith("UUFA_WR_TRANSFER_ENT_CNS_" + year):
            do_query(query, date + " Transfer Students Entrance Counseling 20" + year + ".xlsx", directory,
                     kak_attachment_list)

        if query.startswith("FA_WR_TRANSFER_STDNTS_FA_SP_" + year) \
                | query.startswith("UUFA_WR_TRANSFER_STU_FA_SP_" + year):
            do_query(query, date + " Transfer Students Fall-Spring 20" + year + ".xlsx", directory,
                     v_attachment_list)

        if query.startswith("UUFA_WR_UG_GR_PLUS_GR_TERM_" + year):
            do_query(query, date + " UG-GR PLUS Grad Term " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_UNDOCUMENTED_STUDENTS_" + year):
            do_query(query, date + " Undocumented Student Awards " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_VERI_CHKLST_MISSING_" + year) \
                | query.startswith("UUFA_WR_VERI_CHKLST_MISSING_" + year):
            do_query(query, date + " Verification Checklist Missing 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_VERI_INCOME_ADJ_20" + year) \
                | query.startswith("UUFA_WR_VERI_INCOME_ADJ_20" + year):
            do_query(query, date + " Income Adjustments 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_VER_NOT_CONSL_20" + year) \
                | query.startswith("UUFA_WR_VER_NOT_CONSL_20" + year):
            do_query(query, date + " Verification Not Consolidated 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_VETERAN_ACTIVE_DUTY_" + year) \
                | query.startswith("UUFA_WR_VETERAN_ACTIVE_DUTY_" + year):
            do_query(query, date + " Veteran Active Duty 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_VETERAN_NO_QUALIFY_" + year) \
                | query.startswith("UUFA_WR_VETERAN_NO_QUALIFY_" + year):
            do_query(query, date + " Veteran No Qualify 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_WEEKS_OF_INSTR_FIX_" + year) \
                | query.startswith("UUFA_WR_WEEKS_OF_INSTR_FIX_" + year):
            do_query(query, date + " Weeks of Instruction 20" + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_WR_DL_AY_SP_CANCELLED_" + year) \
                | query.startswith("UUFA_WR_DL_AY_SP_CANCELED_" + year):
            do_query(query, date + " DL AY SP Cancelled " + year + ".xlsx", directory,
                     kak_attachment_list)

        if query.startswith("FA_WR_LOAN_TRANSMIT_HOLD_13"):
            do_query(query, date + " Loan Transmit Hold 13.xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_WR_RT4_FALL_DROP_CLASSES_" + year) \
                | query.startswith("UUFA_WR_RT4_FA_DROP_CLASSES_" + year):
            do_query(query, date + " RT4 Fall Drop Classes 20" + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_RT4_SPR_DROP_CLASSES_" + year) \
                | query.startswith("UUFA_WR_RT4_SP_DROP_CLASSES_" + year):
            do_query(query, date + " RT4 Spring Drop Classes 20" + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_RT4_SUM_DROP_CLASSES_" + year) \
                | query.startswith("UUFA_WR_RT4_SU_DROP_CLASSES_" + year):
            do_query(query, date + " RT4 Summer Drop Classes 20" + year + ".xlsx", directory,
                     ak_attachment_list)

        # Manually run Queries
        if query.startswith("FA_WR_LOAN_EFT_DETAIL_ERROR") \
                | query.startswith("UUFA_WR_LOAN_EFT_DETAIL_ERROR"):
            do_query(query, date + " Loan EFT Detail Error.xlsx", directory,
                     akk_attachment_list)

        if query.startswith("FA_WR_NSL_PROMISSORY_NOTE") \
                | query.startswith("UUFA_WR_NSL_PROMISSORY_"):
            do_query(query, date + " NSL Promissory Note " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_SNGDO_CAMPUS_20" + year) \
                | query.startswith("UUFA_WR_SNGDO_CAMPUS_20" + year):
            do_query(query, date + " Asian-SNGDO Campus " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_FALL_TOTAL_WTHDRN_DRP_13"):
            do_query(query, date + " Fall Total Withdrawn Drop 13.xlsx", directory2013,
                     akr_attachment_list)

        if query.startswith("FA_WR_SPR_TOTAL_WTHDRN_DRP_13"):
            do_query(query, date + " Spring Total Withdrawn Drop 13.xlsx", directory2013,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUM_TOTAL_WTHDRN_DRP_13"):
            do_query(query, date + " Summer Total Withdrawn Drop 13.xlsx", directory2013,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUMMER_NO_DL_13"):
            do_query(query, date + " Summer Enroll No DL " + year + ".xlsx", directory2013,
                     ak_attachment_list)
            n
        if query.startswith("FA_WR_FALL_TOTAL_WTHDRN_DRP_14") | query.startswith("FA_WR_FALL_TOTAL_WDRN_DRP_14"):
            do_query(query, date + " Fall Total Withdrawn Drop 14.xlsx", directory2014,
                     akr_attachment_list)

        if query.startswith("FA_WR_SPR_TOTAL_WTHDRN_DRP_14") | query.startswith("FA_WR_SPR_TOTAL_WDRN_DRP_14"):
            do_query(query, date + " Spring Total Withdrawn Drop 14.xlsx", directory2014,
                     akr_attachment_list)

        if query.startswith("FA_WR_SUM_TOTAL_WTHDRN_DRP_14") | query.startswith("FA_WR_SUM_TOTAL_WDRN_DRP_14"):
            do_query(query, date + " Summer Total Withdrawn Drop 14.xlsx", directory2014,
                     akr_attachment_list)

        if query.startswith("FA_WR_LOAN_TRANSMIT_HOLD_13"):
            do_query(query, date + " Loan Transmit Hold 13.xlsx", directory2013,
                     ka_attachment_list)

        # Packaging queries that are being manually run.
        if query.startswith("FA_PRT_ATH_ACCEPT_FED_AID_" + year):
            do_query(query, date + " Athlete Accepted Federal Aid " + year + ".xlsx", packaging_directory,
                     lkj_attachment_list)

        if query.startswith("FA_PRT_ATH_AWD_CBA_GRANT_" + year):
            do_query(query, date + " Athlete Awarded CBA Grant " + year + ".xlsx", packaging_directory,
                     lkj_attachment_list)

        if query.startswith("FA_PRT_ATHLETE_GRAD_DATE_" + year):
            do_query(query, date + " Athlete Accepted Grad Date " + year + ".xlsx", packaging_directory,
                     lkj_attachment_list)

        if query.startswith("FA_PRT_ATH_OFFERED_FED_AID_" + year):
            do_query(query, date + " Athlete Offered Federal Aid " + year + ".xlsx", packaging_directory,
                     lkj_attachment_list)

    if atcj_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", atcj_mail, "", atcj_attachment_list)
        del atcj_attachment_list[:]
    if lkj_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", lkj_mail, "", lkj_attachment_list)
        del lkj_attachment_list[:]
    if akr_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", akr_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if act_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", act_mail, "", act_attachment_list)
        del act_attachment_list[:]
    if ak_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if akk_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", akk_mail, "", akk_attachment_list)
        del akk_attachment_list[:]
    if rac_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", rac_mail, "", rac_attachment_list)
        del rac_attachment_list[:]
    if vm_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", vm_mail, "", vm_attachment_list)
        del vm_attachment_list[:]
    if v_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", v_mail, "", v_attachment_list)
        del v_attachment_list[:]
    if ka_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", ka_mail, "", ka_attachment_list)
        del ka_attachment_list[:]
    if kak_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", kak_mail, "", kak_attachment_list)
        del kak_attachment_list[:]
    if kc_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", kc_mail, "", kc_attachment_list)
        del kc_attachment_list[:]
    if akc_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", akc_mail, "", akc_attachment_list)
        del akc_attachment_list[:]
    if kak_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", kak_mail, "", kak_attachment_list)
        del kak_attachment_list[:]
    if kaca_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", kaca_mail, "", kaca_attachment_list)
        del kaca_attachment_list[:]
    if mat_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", mat_mail, "", mat_attachment_list)
        del mat_attachment_list[:]


def do_budget_queries():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_BR_ACAD_LVLS_OUT_OF_SYNC") \
                | query_name.startswith("UUFA_BR_ACAD_LVLS_OUT_SYNC"):
            if date[6:] in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
            if str(int(date[6:]) + 1) in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break

    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    year = aid_year[-2:]

    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Budgets', aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Budgets', aid_year, month_folder))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be query AC it is received and _new_file_name to what the new query should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("FA_BR_ACAD_LVLS_OUT_OF_SYNC_" + year) \
                | query.startswith("UUFA_BR_ACAD_LVLS_OUT_SYNC_" + year):
            do_query(query, date + " GR Academic Levels Out of Sync " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_BR_ATH_TUITION_INCREASE_" + year) \
                | query.startswith("UUFA_BR_ATH_TUIT_INCR_NR_" + year):
            do_query(query, date + " Athlete Tuition Increase " + year + ".xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_BR_ATHLETE_TUIT_INCR_NR_" + year) \
                | query.startswith("UUFA_BR_ATH_TUITION_INCRS_" + year):
            do_query(query, date + " Athlete Tuition Increase Non Resident 20" + year + ".xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_BR_BUDGET_DOUBLE_BUDGETS_" + year) \
                | query.startswith("UUFA_BR_BDGT_DOUBLE_BUDGETS_" + year):
            do_query(query, date + " Double Budget " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("UUFA_BR_COA_LESS_HT_" + year):
            do_query(query, date + " PELL COA Less Than Half Time Enrollment " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_BR_COA_TUIT_ZERO_" + year) \
                | query.startswith("UUFA_BR_COA_TUIT_ZERO_" + year):
            do_query(query, date + " COA Tuition Amount Zero " + year + ".xlsx", directory,
                     akv_attachment_list)
        if query.startswith("FA_BR_DN_LW_MD_STU_AID_ATRB_" + year) \
                | query.startswith("UUFA_BR_DN_LW_MD_AID_ATRB_" + year):
            do_query(query, date + " DN-LW-MD Student Aid Career " + year + ".xlsx", directory,
                     akc_attachment_list)
        if query.startswith("FA_BR_FT_CLASS_OVERRIDES_" + year) \
                | query.startswith("UUFA_BR_FT_CLASS_OVERRIDES_" + year):
            do_query(query, date + " Class Overrides " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_BR_ISIR_SCHOLARSHIP_" + year):
            do_query(query, date + " Scholarship ISIR Received 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_BR_NO_BUDGET_ATTEND_20" + year) \
                | query.startswith("UUFA_BR_NO_BUDGET_ATTEND_" + year):
            do_query(query, date + " NO Budget Attend 20" + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_BR_NSLDS_NO_MATCH_DB_FLG_" + year) \
                | query.startswith("UUFA_BR_NSLDS_NO_MCH_DB_FLG_" + year):
            do_query(query, date + " NSLDS No Match DB Flag " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_BR_OMBAMBA_" + year) \
                | query.startswith("UUFA_BR_OMBAMBA_" + year):
            do_query(query, date + " Academic Plan OMBAMBA " + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_BR_PELL_COA_BLANK_" + year) \
                | query.startswith("UUFA_BR_PELL_COA_BLANK_" + year):
            do_query(query, date + " PELL COA Blank " + year + ".xlsx", directory,
                     aka_attachment_list)
        if query.startswith("FA_BR_PELL_COA_DOUBLE_" + year) \
                | query.startswith("UUFA_BR_PELL_COA_DOUBLE_" + year):
            do_query(query, date + " PELL COA Double " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_BR_PELL_COA_LESS_HT_20" + year):
            do_query(query, date + " PELL COA Less HT Enrollment " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_BR_PROC_STAT_REVIEW_STAT_" + year) \
                | query.startswith("UUFA_BR_PROC_STAT_RVW_STAT_" + year):
            do_query(query, date + " Reset Processing Status to 1 " + year + ".xlsx", directory,
                     ms_attachment_list)
        if query.startswith("FA_BR_RES_NON_RES_BUDGET_" + year) \
                | query.startswith("UUFA_BR_RES_NON_RES_BDGT_" + year):
            do_query(query, date + " Resident - Non-Resident Budget " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_BR_SCH_TUITION_FEES_NR_" + year) \
                | query.startswith("UUFA_BR_SCH_TUITION_FEES_NR_" + year):
            do_query(query, date + " Wavier-Scholar Tuition Fees Res " + year + ".xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_BR_SCH_TUITION_ONLY_NR_" + year) \
                | query.startswith("UUFA_BR_SCH_TUITION_ONLY_NR_" + year):
            do_query(query, date + " Wavier-Scholar Tuition Fees NR " + year + ".xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_BR_SCHOLAR_TUITION_FEES_" + year) \
                | query.startswith("UUFA_BR_SCHOLAR_TUIT_FEES_" + year):
            do_query(query, date + " Wavier-Scholar Tuition Only Res " + year + ".xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_BR_SCHOLAR_TUITION_ONLY_" + year) \
                | query.startswith("UUFA_BR_SCHOLAR_TUIT_ONLY_" + year):
            do_query(query, date + " Wavier-Scholar Tuition Only NR " + year + ".xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_BR_UFORM_CHANGE_BUD_DUR_" + year):
            do_query(query, date + " Correct Budget Duration " + year + ".xlsx", directory,
                     akr_attachment_list)

    if ak_attachment_list:
        mailer("", aid_year + " Budget Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if lkj_attachment_list:
        mailer("", aid_year + " Budget Queries", lkjk_mail, "", lkj_attachment_list)
        del lkj_attachment_list[:]
    if akc_attachment_list:
        mailer("", aid_year + " Budget Queries", akc_mail, "", akc_attachment_list)
        del akc_attachment_list[:]
    if akr_attachment_list:
        mailer("", aid_year + " Budget Queries", akr_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if akv_attachment_list:
        mailer("", aid_year + " Budget Queries", akv_mail, "", akv_attachment_list)
        del akv_attachment_list[:]
    if aka_attachment_list:
        mailer("", aid_year + " Budget Queries", aka_mail, "", aka_attachment_list)
        del aka_attachment_list[:]
    if ms_attachment_list:
        mailer("", aid_year + " Budget Queries", vs_mail, "", ms_attachment_list)
        del ms_attachment_list[:]
    if act_attachment_list:
        mailer("", aid_year + " Budget Queries", act_mail, "", act_attachment_list)
        del act_attachment_list[:]


def do_packaging_queries():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_PRT_ACAD_PROG_REVIEW") \
                | query_name.startswith("UUFA_PRT_ACAD_PROG_REVIEW"):
            if date[6:] in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
            if str(int(date[6:]) + 1) in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Packaging', aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Packaging', aid_year, month_folder))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be query AC it is received and _new_file_name to what the new query should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("FA_PRT_ACAD_LVLS_OUT_OF_SYNC") \
                | query.startswith("UUFA_PRT_ACAD_LVLS_OUT_OF_SYNC"):
            do_query(query, date + " UG Acad Levels Out of Sync.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_ACAD_PROG_REVIEW") \
                | query.startswith("UUFA_PRT_ACAD_PROG_REVIEW_" + year):
            do_query(query, date + " Academic Progress Review.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_ATHLETE_ACCEPT_FED_AID") \
                | query.startswith("UUFA_PRT_ATH_ACCEPT_FED_AID"):
            do_query(query, date + " Athlete Accepted Federal Aid.xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_PRT_ATHLETE_AWD_CBA_GRANT") \
                | query.startswith("UUFA_PRT_ATH_AWD_CBA_GRANT"):
            do_query(query, date + " Athlete Awarded CBA Grant.xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_PRT_ATHLETE_GRAD_DATE") \
                | query.startswith("UUFA_PRT_ATH_GRAD_DATE"):
            do_query(query, date + " Athlete Expected Grad Date.xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_PRT_ATHLETE_OFFERED_FED_AID") \
                | query.startswith("UUFA_PRT_ATH_OFFRD_FED_AID_"):
            do_query(query, date + " Athlete Offered Federal Aid.xlsx", directory,
                     lkj_attachment_list)

            if query.startswith("UUFA_PRT_ATH_OFFR_ACCPT_AID_"):
                do_query(query, date + " ATH Fed State Inst O/A.xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_PRT_AWARD_DATE_HAD_SAT") \
                | query.startswith("UUFA_PRT_AWARD_DATE_HAD_SAT"):
            do_query(query, date + " SAT Hold Date Award Review.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_AWARD_TERM_HAD_SAT") \
                | query.startswith("UUFA_PRT_AWARD_TERM_HAD_SAT"):
            do_query(query, date + " SAT Hold Term Award Review.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_AWARDS_OTHER_INST") \
                | query.startswith("UUFA_PRT_AWARDS_OTHER_INST"):
            do_query(query, date + " Checklist FAOI" + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_AWD_COMB_OVER_AGG_RVW") \
                | query.startswith("UUFA_PRT_AWD_COMB_OVER_AGG_RVW"):
            do_query(query, date + " Award Combined Over Aggregate.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_AWD_MASS_P_NO_AWARDS"):
            do_query(query, date + " Award Mass Packaging No Awards.xlsx", directory,
                     ms_attachment_list)

        if query.startswith("FA_PRT_AWD_PELL_ELG_NO_PELL_" + year) \
                | query.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_" + year):
            do_query(query, date + " PELL ELIGIBLE NO PELL 20" + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_AWD_SUB_OVER_AGG_RVW") \
                | query.startswith("UUFA_PRT_AWD_SUB_OVER_AGG_RVW"):
            do_query(query, date + " SUB Over Aggregate.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_CTZN_IND_AWRD_NO_LOANS") \
                | query.startswith("UUFA_PRT_CTZN_IND_AWD_NO_LOANS"):
            do_query(query, date + " LA-wk eligible - Award No Loans.xlsx", directory,
                     akr_attachment_list)

        if query.startswith("UUFA_PRT_DEFR_ENROLLMENT_") \
                | query.startswith("UUFA_PRT_DEFR_ENROLLMENT_" + year):
            do_query(query, date + " DEFER Enrollment " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_PRT_DEPEND_PRNT_SSN_REVIEW") \
                | query.startswith("UUFA_PRT_DEPEND_PRNT_SSN_RVW"):
            do_query(query, date + " Parent SSN Review.xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_PRT_DIAG_AWARD_PELL_TERM_" + year) \
                | query.startswith("UUFA_PRT_DIAG_AWD_PELL_TERM_" + year):
            do_query(query, date + " Term Pell Awards 20" + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_DISB_PLAN_SPLT_CODE_" + year):
            do_query(query, date + " Disb Plan FY Split Code XX.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_DL_DPAY_SCSP_20" + year) \
                | query.startswith("UUFA_PRT_DL_DPAY_SCSP_" + year):
            do_query(query, date + " Disb Plan AY-Split Code SP.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_DSB_PLN_SPLT_CD_FD_" + year):
            do_query(query, date + " Federal Disb Plan FY Split Code XX " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_DSB_PLN_SPLT_CD_SC_" + year):
            do_query(query, date + " Scholarship Disb Plan FY Split Code XX " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_EXPECT_GRAD_TERM_11"):
            do_query(query, date + " Expected Grad Term 1" + str(int(year) - 1) + "8.xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_PRT_DL_GRAD_TERM_FALL_" + year) \
                | query.startswith("UUFA_PRT_DL_GRAD_TERM_FALL_" + year):
            do_query(query, date + " DL Expected Grad Term Fall " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_GRAD_TERM_FALL_" + year) \
                | query.startswith("UUFA_PRT_GRAD_TRM_FALL_" + year):
            do_query(query, date + " Loan Proration Grad Term Fall " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_GRAD_TERM_SPRING_" + year) \
                | query.startswith("UUFA_PRT_GRAD_TRM_SPRING_" + year):
            do_query(query, date + " Loan Proration Grad Term Spring " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_GRANT_UG_5TH_YR_2BACH") \
                | query.startswith("UUFA_PRT_GRANT_UG_5TH_YR_2BACH"):
            do_query(query, date + " UG 5th YR 2ND Bachelor.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_HEAL_20"):
            do_query(query, date + " Heal 20 " + year + ".xlsx", directory,
                     kak_attachment_list)

        if query.startswith("FA_PRT_LEU_C_FSEOG_20" + year) \
                | query.startswith("UUFA_PRT_LEU_C_FSEOG_" + year):
            do_query(query, date + " LEU C Flag Awarded FSEOG " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_LEU_C_PELL_AWARD_" + year) \
                | query.startswith("UUFA_PRT_LEU_C_PELL_AWARD_" + year):
            do_query(query, date + " LEU C Flag Awarded Pell " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_LEU_E_FSEOG_20" + year) \
                | query.startswith("UUFA_PRT_LEU_E_FSEOG_" + year):
            do_query(query, date + " LEU E Flag Awarded FSEOG " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_LOAN_CBA_AWD_NOT_ELIG") \
                | query.startswith("UUFA_PRT_LOAN_CBA_AWD_NOT_ELIG"):
            do_query(query, date + " Loan CBA Review Eligible.xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_PRT_NO_FALL_1" + str(int(year) - 1) + "8") \
                | query.startswith("UUFA_PRT_NO_FALL_11" + str(int(year) - 1) + "8"):
            do_query(query, date + " Packaging No Fall 1" + str(int(year) - 1) + "8 (2).xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_PRT_NSL_LOAN_RPT_VERI_20" + year) \
                | query.startswith("UUFA_PRT_NSL_LOAN_RPT_VERI_" + year):
            do_query(query, date + " NSL Loan Need Verification.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_NURSING_LOAN_RPT_20" + year) \
                | query.startswith("UUFA_PRT_NURSING_LOAN_RPT_" + year):
            do_query(query, date + " NSL Needs NSL P-N Checklist " + year + ".xlsx", directory,
                     kc_attachment_list)

        if query.startswith("FA_PRT_ON_LINE_PACKAGING") \
                | query.startswith("UUFA_PRT_ON_LINE_PACKAGING"):
            do_query(query, date + " Manual Packaging Counselors.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_PELL_COMMENT_037_" + year) \
                | query.startswith("UUFA_PRT_PELL_COMMENT_037_" + year):
            do_query(query, date + " Pell Comment Code 037 20" + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_PELL_EL_CTZN_NOT_INDCT") \
                | query.startswith("UUFA_PRT_PLL_EL_CTZN_NOT_INDCT"):
            do_query(query, date + " Pell Eligible Citizenship Not Indicated.xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_PRT_PELL_FPEL" + year + "_INITIATED") \
                | query.startswith("UUFA_PRT_PELL_FPEL" + year + "_INITIATED"):
            do_query(query, date + " Pell FPEL" + year + " Initiated.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_PELL_UG_5TH_YR_2ND_BACH") \
                | query.startswith("UUFA_PRT_PELL_UG_5TH_YR_2ND_BA"):
            do_query(query, date + " Pell UG 5th YR.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_PHARM_NO_HEAL") \
                | query.startswith("UUFA_PRT_PHARM_NO_HEAL"):
            do_query(query, date + " Pharmacy students with NO HEAL.xlsx", directory,
                     akk_attachment_list)

        if query.startswith("UUFA_PRT_PKG_AWARD_NO_BDGT"):
            do_query(query, date + " Pell Com Award NO Budget for Term.xlsx", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_PRIOR_TERM_STFFRD_OFR"):
            do_query(query, date + " Cancel Prior Term Stafford Offer " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_PRT_READY_PKG_" + year + "_ACTIVE"):
            do_query(query, date + " Manual Awd Pkg Active 20" + year + ".xlsx", directory,
                     ms_attachment_list)

        if query.startswith("FA_PRT_SCHOL_GRAD_DATE") \
                | query.startswith("UUFA_PRT_SCHOL_GRAD_DATE"):
            do_query(query, date + " Scholarship-Expected Grad Date.xlsx", directory,
                     act_attachment_list)

        if query.startswith("FA_PRT_SET_HEAL_FLAG_" + year) \
                | query.startswith("UUFA_PRT_SET_HEAL_FLAG_" + year):
            do_query(query, date + " MD - Pharmacy Heal Eligible Flag.xlsx", directory,
                     akrk_attachment_list)

        if query.startswith("FA_PRT_STATE_OF_RES_FM_MH_PW") \
                | query.startswith("UUFA_PRT_STATE_OF_RES_FM_MH_PW"):
            do_query(query, date + " State of Residence FM MH PW.xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_PRT_STILL_UNPRCD_AFTER_PKG"):
            do_query(query, date + " Students Not Packaged (old).xlsx", directory,
                     ms_attachment_list)

        if query.startswith("FA_PRT_STUDENT_NOT_PACKAGED_" + year) \
                | query.startswith("UUFA_PRT_STDNT_NOT_PACKAGED_" + year):
            do_query(query, date + " Students Not Packaged " + year + ".xlsx", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_TEACH_CREDENTIAL_" + year) \
                | query.startswith("UUFA_PRT_TEACH_CREDENTIAL_" + year):
            do_query(query, date + " Teach Credential 20" + year + ".xlsx", directory,
                     ak_attachment_list)

    if ak_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if lkj_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", lkj_mail, "", lkj_attachment_list)
        del lkj_attachment_list[:]
    if akr_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", akr_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if kc_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", kc_mail, "", kc_attachment_list)
        del kc_attachment_list[:]
    if ms_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", vs_mail, "", ms_attachment_list)
        del ms_attachment_list[:]
    if akk_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", akk_mail, "", akk_attachment_list)
        del akk_attachment_list[:]
    if act_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", act_mail, "", act_attachment_list)
        del act_attachment_list[:]
    if akrk_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", akrk_mail, "", akrk_attachment_list)
        del akrk_attachment_list[:]
    if akar_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", akar_mail, "", akar_attachment_list)
        del akar_attachment_list[:]
    if kak_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", kak_mail, "", kak_attachment_list)
        del kak_attachment_list[:]


def do_monthlies():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_MR_ATHLETE_RESIDENCY_") \
                | query_name.startswith("UUFA_MR_ATHLETE_RESIDENCY_"):
            if date[6:] in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
            if str(int(date[6:]) + 1) in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Monthly', aid_year, month_folder))
        acct_directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/acct/Chartfields', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monthly', aid_year, month_folder))
        acct_directory = os.path.realpath(os.path.join('O:/acct/Chartfields', aid_year))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(acct_directory):
        os.makedirs(acct_directory)

    # Change File_Name to be file ac it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("FA_MR_THIRD_PARTY_CROSSWALK_" + year) \
                | query.startswith("UUFA_MR_3RD_PARTY_CROSSWALK_" + year):
            do_query(query, date + " Third Party Crosswalk " + year + ".xlsx", directory,
                     rac_attachment_list)

        if query.startswith("FA_MR_THRD_PRTY_MNTR_IA_ALL_" + year) \
                | query.startswith("UUFA_MR_3RD_PRT_MNTR_IA_ALL_" + year):
            do_query(query, date + " Third Party Monitor " + year + ".xlsx", directory,
                     rac_attachment_list)

        if query.startswith("FA_MR_ACAD_LVLS_OUT_OF_SYNC_" + year) \
                | query.startswith("UUFA_MR_ACAD_LVLS_NOT_SYNC_" + year):
            do_query(query, date + " Academic Levels out of SYNC " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_MR_ATHLETE_RESIDENCY_" + year) \
                | query.startswith("UUFA_MR_ATHLETE_RESIDENCY_" + year):
            do_query(query, date + " Residency for Athlete Student " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_MR_ATHLETE_T53_AWARDS_" + year) \
                | query.startswith("UUFA_MR_ATHLETE_T53_AWARDS_" + year):
            do_query(query, date + " Athlete T53 Awards Accepted " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_MR_COD_DL_20" + year) \
                | query.startswith("UUFA_MR_COD_DL_20" + year):
            do_query(query, date + " COD DL FATB" + year + " FCRD" + year + " FHMS" + year + ".xlsx", directory,
                     aka_attachment_list)

        if query.startswith("FA_MR_COD_PELL_TEACH_IASG_20" + year) \
                | query.startswith("UUFA_MR_COD_PELL_TEACH_IASG_" + year):
            do_query(query, date + " COD Grant FCRD" + year + "-FHMS" + year + " Report.xlsx", directory,
                     aka_attachment_list)

        if query.startswith("FA_MR_DISB_ATH_AWD_NOPOST_" + year) \
                | query.startswith("UUFA_MR_DISB_ATH_AWD_NOPOST_" + year):
            do_query(query, date + " Athlete Waiver Disbursed Not Posted " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_MR_DISB_CASH_AWD_NOPOST_" + year) \
                | query.startswith("UUFA_MR_DSB_CASH_AWD_NOPOST_" + year):
            do_query(query, date + " Cash Disbursed Not Posted " + year + ".xlsx", directory,
                     aca_attachment_list)

        if query.startswith("FA_MR_DISB_WAVR_AWD_NOPOST_" + year) \
                | query.startswith("UUFA_MR_DSB_WAVR_AWD_NOPOST_" + year):
            do_query(query, date + " Waiver-Scholarship Disbursed Not Posted " + year + ".xlsx", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_MR_DN_INC_CHECKLISTS_" + year):
            do_query(query, date + " Dental Students with I Checklists " + year + ".xlsx", directory,
                     kc_attachment_list)

        if query.startswith("UUFA_MR_GRAD_TERM_PRB_" + year):
            do_query(query, date + " Grad Term Wrong " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_MR_ITEM_CHARTFIELD_SETUP_" + year) \
                | query.startswith("UUFA_MR_ITEM_CHARTFLD_SETUP_" + year):
            do_query(query, date + " Item Chartfield Setup.xlsx", acct_directory,
                     ac_attachment_list)

        if query.startswith("FA_MR_ITEM_TYPE_DISB_RULE_" + year) \
                | query.startswith("UUFA_MR_ITEM_TYPE_DISB_RULE_" + year):
            do_query(query, date + " Item Type Career - Match Disb Rule Career " + year + ".xlsx", directory,
                     acvj_attachment_list)

        if query.startswith("FA_MR_LAW_INC_CHECKLISTS_" + year) \
                | query.startswith("UUFA_MR_LAW_INC_CHECKLISTS_" + year):
            do_query(query, date + " Law Students with I Checklists " + year + ".xlsx", directory,
                     kc_attachment_list)

        if query.startswith("FA_MR_LOAN_AWD_PARTIAL_DISB_" + year) \
                | query.startswith("UUFA_MR_LOAN_AWD_PARTL_DISB_" + year):
            do_query(query, date + " Loan Awards Partial Disbursed " + year + ".xlsx", directory,
                     ka_attachment_list)

        if query.startswith("FA_MR_MED_INC_CHECKLISTS_" + year) \
                | query.startswith("UUFA_MR_MED_INC_CHECKLISTS_" + year):
            do_query(query, date + " Med Students with I Checklists " + year + ".xlsx", directory,
                     kc_attachment_list)

        if query.startswith("FA_MR_MED_LAW_LEVEL_REVIEW_" + year) \
                | query.startswith("UUFA_MR_MED_LAW_LVL_REVIEW_" + year) \
                | query.startswith("UUFA_MR_DN_LW_MD_LVL_RVW_" + year):
            do_query(query, date + " MED-LAW Academic Level Review " + year + ".xlsx", directory,
                     ac_attachment_list)

        if query.startswith("FA_MR_PARTIAL_TW_OTHER_SCH_" + year) \
                | query.startswith("UUFA_MR_PART_TW_OTHER_SCH_" + year):
            do_query(query, date + " Partial TW Other Scholarship " + year + ".xlsx", directory,
                     act_attachment_list)

        if query.startswith("FA_MR_PELL_AWARD_ADJUSTMENT_" + year) \
                | query.startswith("UUFA_MR_PELL_AWD_ADJUSTMENT_" + year):
            do_query(query, date + " Pell Award Adjust " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_MR_PELL_ONLY_" + year) \
                | query.startswith("UUFA_MR_PELL_ONLY_" + year):
            do_query(query, date + " Pell Awd  Zero Grants Loans " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_MR_PELL_SSN_MISMATCH_" + year) \
                | query.startswith("UUFA_MR_PELL_SSN_MISMATCH_" + year):
            do_query(query, date + " SSN Mismatch " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_MR_PERKINS_CLASS_LIMITS_" + year) \
                | query.startswith("UUFA_MR_PERKINS_CLASS_LIMIT_" + year):
            do_query(query, date + " Perkins Class Limits " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("UUFA_MR_PERK_MISC_LN_CNCLD_" + year):
            do_query(query, date + " Perkins - Misc Loans Cancelled " + year + ".xlsx", directory,
                     rac_attachment_list)

        if query.startswith("UUFA_MR_PERK_MISC_LOAN_DISB_" + year):
            do_query(query, date + " Perkins - Misc Loans Disbursed " + year + ".xlsx", directory,
                     rac_attachment_list)

        if query.startswith("UUFA_MR_SCHOLAR_LOA_" + year):
            do_query(query, date + " Scholarship LOA " + year + ".xlsx", directory,
                     atcj_attachment_list)

        if query.startswith("FA_MR_SF_DIS_AWD_PT_ERR_FC_" + year) \
                | query.startswith("UUFA_MR_SF_DIS_AWD_PT_ER_FC_" + year):
            do_query(query, date + " Federal Award Disb Post Error " + year + ".xlsx", directory,
                     kak_attachment_list)

        if query.startswith("FA_MR_SF_DIS_AWD_PT_ERR_SV_" + year) \
                | query.startswith("UUFA_MR_SF_DIS_AWD_PT_ER_SV_" + year):
            do_query(query, date + " SCHOL-ATH Award Disb Post Error " + year + ".xlsx", directory,
                     act_attachment_list)

        if query.startswith("FA_MR_STATE_FM_MH_PW_" + year) \
                | query.startswith("UUFA_MR_STATE_FM_MH_PW_" + year):
            do_query(query, date + " Palau - Micronesia - Marshall Islands Students " + year + ".xlsx", directory,
                     akar_attachment_list)

        if query.startswith("FA_MR_SUSPEND_RC2_" + year) \
                | query.startswith("UUFA_MR_SUSPEND_RC2_" + year):
            do_query(query, date + " ISIR Suspended Reason Code 2 " + year + ".xlsx", directory,
                     akrv_attachment_list)

        if query.startswith("FA_MR_UFORM_GRAD_TERM_PRB_" + year) \
                | query.startswith("FA_MR_UFORM_GRAD_TERM_PRB_" + year):
            do_query(query, date + " Grad Term Wrong " + year + ".xlsx", directory,
                     rac_attachment_list)

        if query.startswith("FA_MR_UNDS_OFFER_SCHOLAR_" + year) \
                | query.startswith("UUFA_MR_UNDS_OFFER_SCHOLAR_" + year):
            do_query(query, date + " Scholarship Awards UNDS Career " + year + ".xlsx", directory,
                     act_attachment_list)

        if query.startswith("FA_MR_UNDS_OFRD_AMT_ATHLETE_" + year) \
                | query.startswith("UUFA_MR_UNDS_OFRD_AMT_FDRL_" + year):
            do_query(query, date + " Athlete Awards UNDS Career " + year + ".xlsx", directory,
                     lkj_attachment_list)

        if query.startswith("FA_MR_UNDS_OFRD_AMT_FEDERAL_" + year) \
                | query.startswith("UUFA_MR_UNDS_OFRD_AMT_ATH_" + year):
            do_query(query, date + " Federal Awards UNDS Career " + year + ".xlsx", directory,
                     akr_attachment_list)

        if query.startswith("FA_MR_VERIFY_DEPND_OVERRIDE_" + year) \
                | query.startswith("UUFA_MR_VERIFY_DEP_OVERRIDE_" + year):
            do_query(query, date + " Verification Dependency Override " + year + ".xlsx", directory,
                     akr_attachment_list)

    if aka_attachment_list:
        mailer("", aid_year + " Monthly Queries", aka_mail, "", aka_attachment_list)
        del aka_attachment_list[:]
    if akr_attachment_list:
        mailer("", aid_year + " Monthly Queries", akr_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if lkj_attachment_list:
        mailer("", aid_year + " Monthly Queries", lkj_mail, "", lkj_attachment_list)
        del lkj_attachment_list[:]
    if act_attachment_list:
        mailer("", aid_year + " Monthly Queries", act_mail, "", act_attachment_list)
        del act_attachment_list[:]
    if aca_attachment_list:
        mailer("", aid_year + " Monthly Queries", aca_mail, "", aca_attachment_list)
        del aca_attachment_list[:]
    if kc_attachment_list:
        mailer("", aid_year + " Monthly Queries", kc_mail, "", kc_attachment_list)
        del kc_attachment_list[:]
    if ac_attachment_list:
        mailer("", aid_year + " Monthly Queries", ca_mail, "", ac_attachment_list)
        del ac_attachment_list[:]
    if rac_attachment_list:
        mailer("", aid_year + " Monthly Queries", rac_mail, "", rac_attachment_list)
        del rac_attachment_list[:]
    if kak_attachment_list:
        mailer("", aid_year + " Monthly Queries", kak_mail, "", kak_attachment_list)
        del kak_attachment_list[:]
    if ka_attachment_list:
        mailer("", aid_year + " Monthly Queries", ka_mail, "", ka_attachment_list)
        del ka_attachment_list[:]
    if acvj_attachment_list:
        mailer("", aid_year + " Monthly Queries", acvj_mail, "", acvj_attachment_list)
        del acvj_attachment_list[:]
    if akar_attachment_list:
        mailer("", aid_year + " Monthly Queries", akar_mail, "", akar_attachment_list)
        del akar_attachment_list[:]
    if atcj_attachment_list:
        mailer("", aid_year + " Monthly Queries", atcj_mail, "", atcj_attachment_list)
        del atcj_attachment_list[:]
    if akrv_attachment_list:
        mailer("", aid_year + " Monthly Queries", akrv_mail, "", akrv_attachment_list)
        del akrv_attachment_list[:]


def do_end_of_term_queries():
    global aid_year
    term = 'F'
    input = "Error"
    for query_name in os.listdir("."):
        if query_name.startswith("FA_ACADEMIC_PLAN_RVW_FRAP") \
                | query_name.startswith("UUFA_ACADEMIC_PLAN_RVW_FRAP"):
            prompt = "Enter Term: (e.g. 2016U, 2017F, or 2032S):"
            while True:
                input = str.upper(raw_input(prompt))
                aid_year = input[2:4]
                term = input[4]
                if term == 'S' or term == 'U' or term == 'F':
                    break

    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/SAT/', "20" + year + term))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems\QUERIES/SAT/', "20" + year + term))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        if query.startswith("FA_ACADEMIC_PLAN_RVW_FRAP") \
                | query.startswith("UUFA_ACADEMIC_PLAN_RVW_FRAP"):
            do_query(query, date + " Academic plan RVW FRAP " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_ATHLETE_ACAD_PROG_REVIEW") \
                | query.startswith("UUFA_ATHLETE_ACAD_PROG_REVIEW"):
            do_query(query, date + " Athlete Academic Progress Review.xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_ATHLETE_CUR_PART_ALL_AWARDS") \
                | query.startswith("UUFA_ATHLETE_CUR_PART_ALL_AWRD"):
            do_query(query, date + " Athlete Participation Report.xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_CBA_ACAD_PROG_BELOW_FT") \
                | query.startswith("UUFA_CBA_ACAD_PROG_BELOW_FT"):
            do_query(query, date + " CBA Acad Prog below FT.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_CBA_UNDISBURSED") \
                | query.startswith("UUFA_CBA_UNDISBURSED"):
            do_query(query, date + " CBA UNDISBURSED .xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_EOT_ALT_LOAN_SAT") \
                | query.startswith("UUFA_EOT_ALT_LOAN_SAT"):
            do_query(query, date + " Alt Loans Awards with SAT Holds.xlsx", directory,
                     kak_attachment_list)
        if query.startswith("FA_LOANS_ORIG_FAILED_PENDING") \
                | query.startswith("UUFA_LOANS_ORIG_FAILED_PENDING"):
            do_query(query, date + " Loans Originated Failed Pending.xlsx", directory,
                     ka_attachment_list)
        if query.startswith("FA_LOAN_ACAD_PROG_BELOW_HT_UND") \
                | query.startswith("UUFA_LOAN_ACAD_PROG_BLW_HT_UND"):
            do_query(query, date + " Loan Acad Prog below HT Undisbursed.xlsx", directory,
                     ka_attachment_list)
        if query.startswith("FA_LOAN_ACAD_PROG_BELOW_HT_SUB") \
                | query.startswith("UUFA_LOAN_ACAD_PROG_BLW_HT_SUB"):
            do_query(query, date + " Loan Acad Prog below HT Subsq Disb.xlsx", directory,
                     ka_attachment_list)
        if query.startswith("UUFA_PRO_STDNTS_SAT_WARNING"):
            do_query(query, date + " SAT Warning Professional Students.xlsx", directory,
                     akc_attachment_list)
        if query.startswith("FA_LW1_LW2_LW3_SAT_WARNING") \
                | query.startswith("UUFA_LW1_LW2_LW3_SAT_WARNING"):
            do_query(query, date + " SAT Warning LW1 LW2 LW3.xlsx", directory,
                     ka_attachment_list)
        if query.startswith("FA_MED_SAT") \
                | query.startswith("UUFA_MED_SAT"):
            do_query(query, date + " Medical SAT Review.xlsx", directory,
                     kc_attachment_list)
        if query.startswith("FA_MR_FWS_WITH_NSI_HOLD") \
                | query.startswith("UUFA_MR_FWS_WITH_NSI_HOLD_" + year):
            do_query(query, date + " FWS with NSI Holds.xlsx", directory,
                     a_attachment_list)
        if query.startswith("FA_PELL_ACAD_PROG_LESS_THAN") \
                | query.startswith("UUFA_PELL_ACAD_PROG_LES_THN_" + year):
            do_query(query, date + " Pell Less Than " + year + ".xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_PELL_OFFERED_NOT_DIS") \
                | query.startswith("UUFA_PELL_OFFERED_NOT_DIS"):
            do_query(query, date + " Pell Offered Not Disbursed.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_THESIS_STUDENTS_NONRES") \
                | query.startswith("UUFA_THESIS_STUDENTS_NONRES"):
            do_query(query, date + " Thesis Students Non-Resident .xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_SCH_PROBATION_ACAD_PROG_RVW") \
                | query.startswith("UUFA_SCH_PROB_ACAD_PROG_RVW"):
            do_query(query, date + " U Tradition Review.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SCH_U_TRAD_ACAD_PROG_REN") \
                | query.startswith("UUFA_SCH_U_TRAD_ACAD_PROG_REN"):
            do_query(query, date + " U Tradition Renewal.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SCH_U_TRAD_ACAD_PROG_RVW") \
                | query.startswith("UUFA_SCH_U_TRAD_ACAD_PROG_RVW"):
            do_query(query, date + " Scholarship Probation Review.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SCHOLAR_ACAD_PROG_REVIEW") \
                | query.startswith("UUFA_SCHOLAR_ACAD_PROG_REVIEW"):
            do_query(query, date + " Scholarship Academic Progress Review Basic.xlsx", directory,
                     act_attachment_list)
        if query.startswith("UUFA_SAT_AGGCP_LAW") \
                | query.startswith("UUFA_SAT_AGGCP_LAW"):
            do_query(query, date + " SAT Aggregate Law Career.xlsx", directory,
                     akc_attachment_list)
        if query.startswith("UUFA_SAT_AGGCP_MED") \
                | query.startswith("UUFA_SAT_AGGCP_MED"):
            do_query(query, date + " SAT Aggregate Med Career.xlsx", directory,
                     akc_attachment_list)
        if query.startswith("UUFA_SAT_AGGCP_DN"):
            do_query(query, date + " SAT Aggregate Dental Career.xlsx", directory,
                     akc_attachment_list)
        if query.startswith("FA_EU_FALL_GRADE") \
                | query.startswith("UUFA_EU_FALL_GRADE"):
            do_query(query, date + " EU Grade Fall " + str(int(year) - 1) + ".xlsx", directory,
                     akl_attachment_list)
        if query.startswith("FA_EU_SPRING_GRADE") \
                | query.startswith("UUFA_EU_SPRING_GRADE"):
            do_query(query, date + " EU Grade Spring " + year + ".xlsx", directory,
                     akl_attachment_list)
        if query.startswith("FA_EU_SUMMER_GRADE") \
                | query.startswith("UUFA_EU_SUMMER_GRADE"):
            do_query(query, date + " EU Grade Summer " + year + ".xlsx", directory,
                     akl_attachment_list)
        if query.startswith("FA_PELL_ELIG_ENROLL_NO_AWARD") \
                | query.startswith("UUFA_PELL_ELIG_ENROLL_NO_AWARD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_SAT_FSAP") \
                | query.startswith("UUFA_SAT_FSAP"):
            do_query(query, date + " FSAP Students.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_SCHOLAR_ACAD_PROG_RVW_TT") \
                | query.startswith("UUFA_SCHOLAR_ACAD_PROG_RVW_TT"):
            do_query(query, date + " Scholarship Academic Progress Review Top Ten.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SCHOLAR_ALUMNI_CGPA") \
                | query.startswith("UUFA_SCHOLAR_ALUMNI_CGPA"):
            do_query(query, date + " Scholarship Alumni CGPA.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SCHOLAR_CASH_NO_AWARD") \
                | query.startswith("UUFA_SCHOLAR_CASH_NO_AWARD"):
            do_query(query, date + " Scholarship Cash NO Award.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SCHOLAR_LEADER_CGPA") \
                | query.startswith("UUFA_SCHOLAR_LEADER_CGPA"):
            do_query(query, date + " Scholarship CGPA Leadership.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SU_SCHOLAR_ACAD_PROG_REVIEW"):
            do_query(query, date + " Scholarship Academic Progress Review Summer.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SU_SCHOLAR_ACAD_PROG_RVW_TT") \
                | query.startswith("UUFA_SU_SCHLR_ACAD_PROG_RVW_TT"):
            do_query(query, date + " Scholarship Academic Progress Review Top 10 Summer.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SU_SCHOLAR_ALUMNI_CGPA") \
                | query.startswith("UUFA_SU_SCHOLAR_ALUMNI_CGPA"):
            do_query(query, date + " Scholarship CGPA Alumni - Summer.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_SU_SCHOLAR_LEADER_CGPA") \
                | query.startswith("UUFA_SU_SCHOLAR_LEADER_CGPA"):
            do_query(query, date + " Scholarship CGPA Leader - Summer.xlsx", directory,
                     act_attachment_list)

    if a_attachment_list:
        mailer("", "End of Term Queries", a_mail, "", a_attachment_list)
        del a_attachment_list[:]
    if act_attachment_list:
        mailer("", "End of Term Queries", act_mail, "", act_attachment_list)
        del act_attachment_list[:]
    if ak_attachment_list:
        mailer("", "End of Term Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if akc_attachment_list:
        mailer("", "End of Term Queries", akc_mail, "", akc_attachment_list)
        del akc_attachment_list[:]
    if akr_attachment_list:
        mailer("", "End of Term Queries", akr_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if akl_attachment_list:
        mailer("", "End of Term Queries", akl_mail, "", akl_attachment_list)
        del akl_attachment_list[:]
    if ka_attachment_list:
        mailer("", "End of Term Queries", ka_mail, "", ka_attachment_list)
        del ka_attachment_list[:]
    if kak_attachment_list:
        mailer("", "End of Term Queries", kak_mail, "", kak_attachment_list)
        del kak_attachment_list[:]
    if kc_attachment_list:
        mailer("", "End of Term Queries", kc_mail, "", kc_attachment_list)
        del kc_attachment_list[:]
    if lkj_attachment_list:
        mailer("", "End of Term Queries", lkj_mail, "", lkj_attachment_list)
        del lkj_attachment_list[:]


def do_disb_queries():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_DQ_AUTHORIZED_NOT_DISB") \
                | query_name.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB"):
            aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
            break
    disb_date = str(raw_input("Enter Date the Disbursement ran in 'MM-DD-YY' format:"))
    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Disbursement',
                                                  aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems\QUERIES\Disbursement', aid_year, month_folder))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        if query.startswith("FA_DQ_AUTHORIZED_NOT_DISB_" + year) \
                | query.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB_" + year):
            do_query(query, disb_date + " Item Types Authorized Not Disbursed " + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_DQ_ATHLETE_RM_BD_" + year):
            do_query(query, disb_date + " Athlete Room and Board " + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_DQ_ATH_OFF_SCHED_RM_BD_" + year):
            do_query(query, disb_date + " Athlete Off Schedule R&B " + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_CASH_DISB_TOTALS_20" + year) \
                | query.startswith("UUFA_CASH_DISB_TOTALS_20" + year):
            do_query(query, disb_date + " Cash Disbursement Totals " + year + ".xlsx", directory,
                     sys_attachment_list)
        if query.startswith("FA_DL_FALL_20" + year) \
                | query.startswith("UUFA_DQ_FALL_" + year):
            do_query(query, disb_date + " DL Fall Awards 20" + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_DL_FALL_SPRING_20" + year) \
                | query.startswith("UUFA_DQ_FALL_SPRING_" + year):
            do_query(query, disb_date + " DL Fall Spring Awards 20" + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_DL_SPRING_20" + year) \
                | query.startswith("UUFA_DQ_SPRING_" + year):
            do_query(query, disb_date + " DL Spring Awards 20" + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_DL_UG_PLUS_REFUND_IA_" + year) \
                | query.startswith("UUFA_DQ_UG_PLUS_REFUND_IA_" + year):
            do_query(query, disb_date + " DL UG PLUS Refund Borrower " + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_MISC_DISB_TOTALS_20" + year) \
                | query.startswith("UUFA_MISC_DISB_TOTALS_20" + year):
            do_query(query, disb_date + " Misc Disbursement Totals " + year + ".xlsx", directory,
                     sys_attachment_list)
        if query.startswith("FA_MISC_RESOURCES_DISB_20" + year) \
                | query.startswith("UUFA_DQ_MISC_RESOURCE_DISB_" + year):
            do_query(query, disb_date + " Misc Resources Disbursement " + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_NONCASH_DISB_TOTALS_20" + year) \
                | query.startswith("UUFA_NONCASH_DISB_TOTALS_20" + year):
            do_query(query, disb_date + " Non-cash Disbursement Totals " + year + ".xlsx", directory,
                     sys_attachment_list)
        if query.startswith("FA_PELL_AWDS_ACPT_GRTR_DISB_" + year) \
                | query.startswith("UUFA_DQ_PELL_ACPT_GR8_DISB_" + year):
            do_query(query, disb_date + " Pell Accepted Awards Greater Disb " + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_SF_ITEM_TYPE_ERROR") \
                | query.startswith("UUFA_DQ_SF_ITEM_TYPE_ERROR"):
            do_query(query, disb_date + " FA SF Item Type Error " + year + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_TEACH_GRANT_RECIPIENTS_" + str(int(year) - 1)) \
                | query.startswith("UUFA_DQ_TEACH_GRANT_" + str(int(year) - 1)):
            do_query(query, disb_date + " Teach Grant Recipients 20" + str(int(year) - 1) + ".xlsx", directory,
                     disb_attachment_list)
        if query.startswith("FA_TEACH_GRANT_RECIPIENTS_" + year) \
                | query.startswith("UUFA_DQ_TEACH_GRANT_" + year):
            do_query(query, disb_date + " Teach Grant Recipients 20" + year + ".xlsx", directory,
                     disb_attachment_list)

    if disb_attachment_list:
        mailer("", aid_year + " Disbursement Queries", disb_mail, "", disb_attachment_list)
        del disb_attachment_list[:]
    if sys_attachment_list:
        mailer("", aid_year + " Disbursement Queries", sys_mail, "", sys_attachment_list)
        del sys_attachment_list[:]


def do_2nd_ldr():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_PRT_AWD_PELL_ELG_NO_PELL_") \
                | query_name.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_"):
            aid_year = "20" + str(int(query_name[-18:-16]) - 1) + "-20" + query_name[-18:-16]
            break

    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Term', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Term', aid_year))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        if query.startswith("FA_CBA_UNDISBURSED") \
                | query.startswith("UUFA_CBA_UNDISBURSED"):
            do_query(query, date + " CBA Undisbursed.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("FA_HRS_DECREASE_ATH") \
                | query.startswith("UUFA_HRS_DECREASE_ATH"):
            do_query(query, date + " Hours Decrease Athlete.xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_HRS_DECREASE_FC") \
                | query.startswith("UUFA_HRS_DECREASE_FC"):
            do_query(query, date + " Hours Decrease FC.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_HRS_DECREASE_SV") \
                | query.startswith("UUFA_HRS_DECREASE_SV"):
            do_query(query, date + " Hours Decrease SV.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_PELL_ELIG_ENROLL_NO_AWARD") \
                | query.startswith("UUFA_PELL_ELIG_ENROLL_NO_AWARD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_PELL_OFFERED_NOT_DISB") \
                | query.startswith("UUFA_PELL_OFFERED_NOT_DISB"):
            do_query(query, date + " Pell Offered Not Disbursed.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_PELL_DL_MATH990") \
                | query.startswith("UUFA_PELL_DL_MATH990"):
            do_query(query, date + " Pell-DL Math 990.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_PELL_DL_ELI575_ELI685") \
                | query.startswith("UUFA_PELL_DL_ELI575_ELI685"):
            do_query(query, date + " Pell-DL ELI0075 ELI0085.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_SF_DISB_ATHLETIC_AWD_NOPOST") \
                | query.startswith("UUFA_SF_DISB_ATHLETIC_AWD_NOPOST"):
            do_query(query, date + " Athlete Awd Disb not Posted.xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_SF_DISB_WAIVER_AWD_NOPOST") \
                | query.startswith("UUFA_SF_DISB_WAIVER_AWD_NOPOST"):
            do_query(query, date + " Waiver Awd Disb not Posted.xlsx", directory,
                     act_attachment_list)
        if query.startswith("FA_PRT_AWD_PELL_ELG_NO_PELL") \
                | query.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_" + year):
            do_query(query, date + " Pell Eligible No Pell 20" + year + ".xlsx", directory,
                     akr_attachment_list)

    if ak_attachment_list:
        mailer("", "Second Session LDR Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if akr_attachment_list:
        mailer("", "Second Session LDR Queries", akr_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if lkj_attachment_list:
        mailer("", "Second Session LDR Queries", lkj_mail, "", lkj_attachment_list)
        del lkj_attachment_list[:]
    if act_attachment_list:
        mailer("", "Second Session LDR Queries", act_mail, "", act_attachment_list)
        del act_attachment_list[:]


def do_day_after_ldr():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_PRT_AWD_PELL_ELG_NO_PELL_") \
                | query_name.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_"):
            aid_year = "20" + str(int(query_name[-12:-10]) - 1) + "-20" + query_name[-12:-10]
            break

    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/LDR', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/LDR', aid_year))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("FA_LDR_MINIMUM_ENROLLMENT_ATH") \
                | query.startswith("UUFA_LDR_MIN_ENROLLMENT_ATH"):
            do_query(query, date + " Minimum Enrollment Athlete.xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_LDR_MINIMUM_ENROLLMENT_FC") \
                | query.startswith("UUFA_LDR_MINIMUM_ENROLLMENT_FC"):
            do_query(query, date + " Minimum Enrollment FC (Federal & Campus Based Aid).xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_LDR_MINIMUM_ENROLLMENT_SV") \
                | query.startswith("UUFA_LDR_MINIMUM_ENROLLMENT_SV"):
            do_query(query, date + " Minimum Enrollment SV (Scholarships & Waivers).xlsx", directory,
                     ac_attachment_list)
        if query.startswith("FA_LDR_PELL_AWARDS") \
                | query.startswith("UUFA_LDR_PELL_AWARDS"):
            do_query(query, date + " Pell Awards.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_ATHLETE_AWARD_DISBURSED") \
                | query.startswith("UUFA_ATHLETE_AWARD_DISBURSED"):
            do_query(query, date + " Athlete Award Disbursed.xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_CBA_UNDISBURSED") \
                | query.startswith("UUFA_CBA_UNDISBURSED"):
            do_query(query, date + " CBA Undisbursed.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_PELL_DL_MATH990") \
                | query.startswith("UUFA_PELL_DL_MATH990"):
            do_query(query, date + " Pell-DL MATH 990.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_PELL_DL_ELI575_ELI685") \
                | query.startswith("UUFA_PELL_DL_ELI575_ELI685"):
            do_query(query, date + " Pell-DL ELI575 ELI685.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_PELL_ELIG_ENROLL_NO_AWARD") \
                | query.startswith("UUFA_PELL_ELIG_ENROLL_NO_AWARD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_PELL_OFFERED_NOT_DISB") \
                | query.startswith("UUFA_PELL_OFFERED_NOT_DISB"):
            do_query(query, date + " Pell Offered Not Disbursed.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_PELL_SUMMER_ENROLLMENT") \
                | query.startswith("UUFA_PELL_SUMMER_ENROLLMENT"):
            do_query(query, date + " Pell Summer Enrollment Check.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_THESIS_STUDENTS_NONRES") \
                | query.startswith("UUFA_THESIS_STUDENTS_NONRES"):
            do_query(query, date + " Thesis Students Non-Resident.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_REGISTERED_CENSUS_DATE") \
                | query.startswith("UUFA_REGISTERED_CENSUS_DATE"):
            do_query(query, date + " LDR FA Load Check.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("FA_SF_DISB_ATHLETIC_AWD_NOPOST") \
                | query.startswith("UUFA_SF_DISB_ATH_AWD_NOPOST"):
            do_query(query, date + " Athlete Awd Disb not Posted.xlsx", directory,
                     lkj_attachment_list)
        if query.startswith("FA_SF_DISB_WAIVER_AWD_NOPOST") \
                | query.startswith("UUFA_SF_DISB_WAIVER_AWD_NOPOST"):
            do_query(query, date + " Waiver Awd Disb Not Posted.xlsx", directory,
                     ac_attachment_list)
        if query.startswith("FA_PRT_AWD_PELL_ELG_NO_PELL_" + year) \
                | query.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_" + year):
            do_query(query, date + " Pell Eligible No Pell 20" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_NO_MTRC_STU_ATH_BAL_OWING"):
            do_query(query, date + " Non-Matric Stu Athlete Balance Owing.xlsx", directory,
                     lkj_attachment_list)

    if lkj_attachment_list:
        mailer("", "Day After LDR Queries", lkj_mail, "", lkj_attachment_list)
        del lkj_attachment_list[:]
    if akr_attachment_list:
        mailer("", "Day After LDR Queries", akr_mail, "", akr_attachment_list)
        del akr_attachment_list[:]
    if ac_attachment_list:
        mailer("", "Day After LDR Queries", ac_mail, "", ac_attachment_list)
        del ac_attachment_list[:]


def dl_pre_outbound():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_DLR_ENTRANCE_COUNSELING") \
                | query_name.startswith("UUFA_DLR_ENTRANCE_COUNSEL_"):
            aid_year = "20" + str(int(query_name[-18:-16]) - 1) + "-20" + query_name[-18:-16]
            break
    year = aid_year[-2:]
    orig_file_doc = date + " DL ORIG 20" + year + ".doc"
    orig_file_docx = date + " DL ORIG 20" + year + ".docx"
    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Direct Loans', aid_year, 'DL Pre-Outbound'))
        orig_doc = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test\Direct Loans', aid_year, 'Origination', orig_file_doc))
        orig_docx = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test\Direct Loans', aid_year, 'Origination', orig_file_docx))

    else:
        directory = os.path.realpath(os.path.join('O:/Systems/Direct Loans', aid_year, 'DL Pre-Outbound'))
        orig_doc = os.path.realpath(os.path.join('O:/Systems/Direct Loans', aid_year, 'Origination', orig_file_doc))
        orig_docx = os.path.realpath(os.path.join('O:/Systems/Direct Loans', aid_year, 'Origination', orig_file_docx))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("FA_DLR_ENTRANCE_COUNSELING") \
                | query.startswith("UUFA_DLR_ENTRANCE_COUNSEL_" + year):
            do_query(query, date + " DL Entrance Counseling I  " + year + ".xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_LOAN_EFT_DT_LNDR_ERROR") \
                | query.startswith("UUFA_DLR_LOAN_EFT_DT_LNDR_ERR"):
            do_query(query, date + " Loan EFT Date Lender Error.xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_LOAN_NO_NSLDS") \
                | query.startswith("UUFA_DLR_LOAN_NO_NSLDS_" + year):
            do_query(query, date + " Loan No NSLDS " + year + ".xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_LOAN_ORIG_ACAD_LVL") \
                | query.startswith("UUFA_DLR_LOAN_ORIG_ACAD_LVL_" + year):
            do_query(query, date + " Loans Academic Level " + year + ".xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_LOAN_ORIG_EDIT_ERRORS") \
                | query.startswith("UUFA_DLR_LOAN_ORIG_EDIT_ERR"):
            do_query(query, date + " Loan Originate Edit Errors.xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_LOAN_ORIG_SPLT_CDES") \
                | query.startswith("UUFA_DLR_LOAN_ORIG_SPLT_CDS_" + year):
            do_query(query, date + " Loan Split Codes " + year + ".xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_LOAN_ORIG_VLOAN_REASONS") \
                | query.startswith("UUFA_DLR_LOAN_ORIG_VLOAN_RSN"):
            do_query(query, date + " Loan ORIG VLOAN Reasons.xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_LOAN_SPC_NEED_OVAWD") \
                | query.startswith("UUFA_DLR_LOAN_SPC_NEED_OVWD_" + year):
            do_query(query, date + " Loan Overaward Special Need " + year + ".xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_NOT_DISBURSED") \
                | query.startswith("UUFA_DLR_NOT_DISBURSED_" + year):
            do_query(query, date + " DL Disbursement Errors " + year + ".xlsx", directory,
                     dl_attachment_list)
        if query.startswith("FA_DLR_UG_PLUS_REFND_INDICATOR") \
                | query.startswith("UUFA_DLR_UG_PLUS_REFND_IND"):
            do_query(query, date + " DL UG PLUS Refund Indicator.xlsx", directory,
                     dl_attachment_list)
    if not test:
        while True:
            if os.path.isfile(orig_doc):
                dl_attachment_list.append(orig_doc)
                break
            if os.path.isfile(orig_docx):
                dl_attachment_list.append(orig_docx)
                break
            else:
                raw_input("\nCould not locate DL ORIG 20" + year + ".doc\nMake sure it is located in O:/Systems/Direct Loans/" + aid_year +
                          "/Origination\n\nPress Enter when ready.")

    mailer("", aid_year + " Pre-Outbound Queries", dl_mail, "", dl_attachment_list)
    del dl_attachment_list[:]


def al_pre_outbound():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("FA_ALR_LOAN_ORIG_LNDR_NT_CK_") \
                | query_name.startswith("UUFA_ALR_LOAN_ORG_LND_NT_CK_"):
            aid_year = "20" + str(int(query_name[-17:-15]) - 1) + "-20" + query_name[-17:-15]
            break
    skip = "n"
    year = aid_year[-2:]
    orig_file_doc = date + " ALT Loan ORIG 20" + year + ".doc"
    orig_file_docx = date + " ALT Loan ORIG 20" + year + ".docx"

    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/ALT Loans/', aid_year))
        orig_doc = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/ALT Loans/', aid_year, orig_file_doc))
        orig_docx = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/ALT Loans/', aid_year, orig_file_docx))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/ALT Loans/', aid_year))
        orig_doc = os.path.realpath(os.path.join('O:/Systems/QUERIES/ALT Loans/', aid_year, orig_file_doc))
        orig_docx = os.path.realpath(os.path.join('O:/Systems/QUERIES/ALT Loans/', aid_year, orig_file_docx))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("FA_ALR_110_CHNG_PENDING_TRANS") \
                | query.startswith("UUFA_ALR_110_CHNG_PNDNG_TRANS"):
            do_query(query, date + " Loan Pending Change Transactions.xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_CL_APP_RESPONSE_ERRS") \
                | query.startswith("UUFA_ALR_CL_APP_RESPONSE_ERRS"):
            do_query(query, date + " CL Response Load Errors.xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_LOAN_EFT_DT_LNDR_ERROR") \
                | query.startswith("UUFA_ALR_LOAN_EFT_DT_LNDR_ERR"):
            do_query(query, date + " Loan EFT Date Lender Errors.xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_LOAN_ORIG_ACAD_LVL_" + year) \
                | query.startswith("UUFA_ALR_LOAN_ORIG_ACAD_LVL_" + year):
            do_query(query, date + " Loans Academic Level 20" + year + ".xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_LOAN_ORIG_EDIT_ERRORS") \
                | query.startswith("UUFA_ALR_LOAN_ORIG_EDIT_ERRORS"):
            do_query(query, date + " Loan Originate Edit Errors.xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_LOAN_ORIG_FA_LOAD_" + year) \
                | query.startswith("UUFA_ALR_LOAN_ORIG_FA_LOAD_" + year):
            do_query(query, date + " Loan ORIG FA Load 20" + year + ".xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_LOAN_ORIG_LNDR_NT_CK_" + year) \
                | query.startswith("UUFA_ALR_LOAN_ORG_LND_NT_CK_" + year):
            do_query(query, date + " Loan ORIG Lender Note 20" + year + ".xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_LOAN_ORIG_SPLT_CDES_" + year) \
                | query.startswith("UUFA_ALR_LOAN_ORIG_SPLT_CDS_" + year):
            do_query(query, date + " Loan ORIG Split Codes 20" + year + ".xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_LOAN_ORIG_VLOAN_REASONS") \
                | query.startswith("UUFA_ALR_LOAN_ORIG_VLOAN_RSN"):
            do_query(query, date + " Loan ORIG VLOAN Reasons.xlsx", directory,
                     alt_attachment_list)
        if query.startswith("FA_ALR_LOAN_SPC_NEED_OVAWD_" + year) \
                | query.startswith("UUFA_ALR_LOAN_SPC_NEED_OVWD_" + year):
            do_query(query, date + " Loan Overaward Special Need " + year + ".xlsx", directory,
                     alt_attachment_list)
    if not test:
        while True:
            if os.path.isfile(orig_doc):
                alt_attachment_list.append(orig_doc)
                break
            if os.path.isfile(orig_docx):
                alt_attachment_list.append(orig_docx)
                break
            if str.capitalize(skip) == "Y":
                break
            else:
                skip = raw_input("\nCould not locate ALT Loan ORIG 20" + year +
                                 ".doc\nMake sure it is located in O:/Systems/Queries/ALT Loans/" +
                                 aid_year + "\n\nPress Enter when ready.")

    mailer("", aid_year + " Alt Loan Queries", alt_mail, "", alt_attachment_list)
    del alt_attachment_list[:]


def do_pre_repackaging():
    global aid_year
    strm = "no STRM found"
    prompt = "What STRM is this for? (ex. 1154 = Spring 2015:"
    aid_year = "not defined"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_PP_"):
            while True:
                strm = raw_input(prompt)
                if len(strm) == 4 and (0 / int(strm) == 0):
                    break
                else:
                    print "\nOnly STRMs in the form of 1(year)4/6/8 are acceptable. Ex: 1168 is Fall of 16."
            if strm[-1] == "8":
                aid_year = "20" + str(strm[1:3]) + "-20" + str(int(strm[1:3]) + 1)
                break
            else:
                aid_year = "20" + str(int(strm[1:3]) - 1) + "-20" + str(strm[1:3])
                break

    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Pell Repackaging', aid_year, strm))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Pell Repackaging', aid_year, strm))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        # Pre-Pell Repackaging Queries
        if query.startswith("UUFA_PP_RPKG_AGGREGATE_LIMITS"):
            do_query(query, date + " Pell AGG Limits Awards Reduced.xlsx", directory,
                     v_attachment_list)
        if query.startswith("UUFA_PP_RPKG_AWD_AY_NO_BDGT"):
            do_query(query, date + " Pell Award AY One STRM Budget.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("UUFA_PP_RPKG_AWD_STRM_INACTIVE"):
            do_query(query, date + " Pell Award STRM Inactive.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("UUFA_PP_RPKG_AWRD_LOCK"):
            do_query(query, date + " Pell Award Lock No FPEL.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_PP_RPKG_COA_DOUBLE"):
            do_query(query, date + " Pell COA Double.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("UUFA_PP_RPKG_LTHT_PELL_COA"):
            do_query(query, date + " Pell COA LTHT.xlsx", directory,
                     ak_attachment_list)
        if query.startswith("UUFA_PP_RPKG_RPKG_NO_BUDGET"):
            do_query(query, date + " Pell Repackaging No Budget.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_PP_RPKG_RPKG_SNAPSHOT"):
            do_query(query, date + " Pell Repackage Snapshot.xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_PP_RPKG_TOTAL_WDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop.xlsx", directory,
                     ak_attachment_list)

    if ak_attachment_list:
        mailer("", "Pre-Pell Only Repackaging Queries", ak_mail, sys_mail, ak_attachment_list)
        del ak_attachment_list[:]
    if akr_attachment_list:
        mailer("", "Pre-Pell Only Repackaging Queries", akr_mail, sys_mail, akr_attachment_list)
        del akr_attachment_list[:]
    if v_attachment_list:
        mailer("", "Pre-Pell Only Repackaging Queries", v_mail, sys_mail, v_attachment_list)
        del v_attachment_list[:]


def do_mid_repack_queries():
    global aid_year
    strm = "no STRM found"
    prompt = "What STRM is this for? (EX. 1154 = Spring 2015:)"
    aid_year = "not defined"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_MP_"):
            while True:
                strm = raw_input(prompt)
                if len(strm) == 4 and (0 / int(strm) == 0):
                    break
                else:
                    print "\nOnly STRMs in the form of 1(year)4/6/8 are acceptable."
            if strm[-1] == "8":
                aid_year = "20" + str(strm[1:3]) + "-20" + str(int(strm[1:3]) + 1)
                break
            else:
                aid_year = "20" + str(int(strm[1:3]) - 1) + "-20" + str(strm[1:3])
                break

    if test:
        directory = os.path.realpath(os.path.join('C:/QueryRunnerProj/Testing/Test/Pell Repackaging', aid_year, strm))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Pell Repackaging', aid_year, strm))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    year = aid_year[7:]

    for query in os.listdir("."):
        if query.startswith("UUFA_MP_RPKG_AID_PROC_STATUS_4"):
            do_query(query, date + " Aid Processing Status 4 Repackage.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_MP_RPKG_AWD_STRM_INACTIVE"):
            do_query(query, date + " Pell Award STRM Inactive.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_MP_RPKG_FCIT"):
            do_query(query, date + " Pell Repackage FCIT" + year + " FDEG" + year + ".xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_MP_RPKG_FDR"):
            do_query(query, date + " Pell Repackage FDR" + year + " Initiated.xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_MP_RPKG_FOVP_FARC_I"):
            do_query(query, date + " Pell REPKG FOVP FACR Initiated.xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_MP_RPKG_ISIR_CMT_346_347"):
            do_query(query, date + " Pell Repackage ISIR CMT 346 347.xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_MP_RPKG_LEAVE"):
            do_query(query, date + " Pell Repackaging Leave Absense.xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_MP_RPKG_TOTAL_WDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_MP_RPKG_TOTAL_WTHDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop (old).xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_MP_RPKG_VAR_1_2"):
            do_query(query, date + " Pell Repackage Flags.xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_MP_RPKG_VER_UNFLAG"):
            do_query(query, date + " Pell Verification Flag Unchecked.xlsx", directory,
                     null_attachment_list)

    if akr_attachment_list:
        mailer("", "Pell Repackaging Mid-Packaging Queries", akr_mail, sys_mail, akr_attachment_list)
        del akr_attachment_list[:]
        del null_attachment_list[:]


def do_after_repackaging():
    global aid_year
    strm = "no STRM found"
    prompt = "What STRM is this for? (ex. 1154 = Spring 2015:"
    aid_year = "not defined"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_AP_"):
            while True:
                strm = raw_input(prompt)
                if len(strm) == 4 and (0 / int(strm) == 0):
                    break
                else:
                    print "\nOnly STRMs in the form of 1(year)4/6/8 are acceptable."
            if strm[-1] == "8":
                aid_year = "20" + str(strm[1:3]) + "-20" + str(int(strm[1:3]) + 1)
                break
            else:
                aid_year = "20" + str(int(strm[1:3]) - 1) + "-20" + str(strm[1:3])
                break

    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/Pell Repackaging', aid_year, strm))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Pell Repackaging', aid_year, strm))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    year = aid_year[7:]

    for query in os.listdir("."):
        # Pell Repackaging Queries
        if query.startswith("UUFA_AP_RPKG_5TH_YR_2ND_BACH"):
            do_query(query, date + " UG 5th YR 2ND Bachelor.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_ACTN"):
            do_query(query, date + " Pell Award Activity.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_AWACT_C"):
            do_query(query, date + " Pell Awards Cancelled.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_AW_ACT"):
            do_query(query, date + " Pell Repackage Activity.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_FPEL_AWARD_LCK"):
            do_query(query, date + " Pell Award Lock FPEL" + year + ".xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_PLAN_ID_BLANK"):
            do_query(query, date + " Pell Repackaging Plan ID Blank.xlsx", directory,
                     v_attachment_list)
        if query.startswith("UUFA_AP_RPKG_RPKG_SNAPSHOT"):
            do_query(query, date + " Pell Repackage Snapshot.xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_AP_RPKG_SAT_HOLD_DELETED"):
            do_query(query, date + " Pell SAT Holds.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_SKIP"):
            do_query(query, date + " Pell Repackage Skip.xlsx", directory,
                     ms_attachment_list)
        if query.startswith("UUFA_AP_RPKG_TERM_FT"):
            do_query(query, date + " Term Pell Awards FT.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_TERM_HT"):
            do_query(query, date + " Term Pell Awards HT.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_TERM_LH"):
            do_query(query, date + " Term Pell Awards LH.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_TERM_NL"):
            do_query(query, date + " Term Pell Awards NL.xlsx", directory,
                     akr_attachment_list)
        if query.startswith("UUFA_AP_RPKG_TERM_TQ"):
            do_query(query, date + " Term Pell Awards TQ.xlsx", directory,
                     akr_attachment_list)

    if akr_attachment_list:
        mailer("", "Pell Only Repackaging Queries", akr_mail, sys_mail, akr_attachment_list)
        del akr_attachment_list[:]
    if ms_attachment_list:
        mailer("", "Pell Only Repackaging Queries", vs_mail, sys_mail, ms_attachment_list)
        del ms_attachment_list[:]
    if v_attachment_list:
        mailer("", "Pell Only Repackaging Queries", v_mail, sys_mail, v_attachment_list)
        del v_attachment_list[:]


def do_daily_scholarships():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_SCHOLAR_DISB_ZERO"):
            if date[6:] in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
            if str(int(date[6:]) + 1) in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break

    year = aid_year[-2:]

    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test', aid_year + ' Scholar\Queries'))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems', aid_year + ' Scholar\Queries'))

    # the list 'my_path' should be populated  with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_SCHOLAR_DISB_ZERO_" + year):
            do_query(query, date + " Scholarships awarded not disbursed " + year + ".xlsx", directory,
                     ss_attachment_list)
        if query.startswith("UUFA_SCHOLAR_TWO_CAREERS_" + year):
            do_query(query, date + " Scholarship Award with Two Careers " + year + ".xlsx", directory,
                     jen_attachment_list)
        if query.startswith("UUFA_SCHOLAR_AUTH_NOT_DISB_" + year):
            do_query(query, date + " Scholar Authorized Not Disbursed " + year + ".xlsx", directory,
                     jen_attachment_list)

    if ss_attachment_list:
        mailer("", aid_year + " Daily Scholarship Queries", ss_mail, "", ss_attachment_list)
        del ss_attachment_list[:]
    if jen_attachment_list:
        mailer("", aid_year + " Daily Scholarship Queries", jen_mail, "", jen_attachment_list)
        del jen_attachment_list[:]


def do_weekly_scholarships():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_SCH_ALL_NEED"):
            if date[6:] in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break
            if str(int(date[6:]) + 1) in query_name:
                aid_year = "20" + str(int(query_name[-15:-13]) - 1) + "-20" + query_name[-15:-13]
                break

    year = aid_year[-2:]

    if test:
        directory = os.path.realpath(os.path.join('C:\QueryRunnerProj\Testing\Test/', aid_year + ' Scholar\Queries'))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems', aid_year + ' Scholar\Queries'))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_DEPT_POST_WRNG_ITEM_" + year):
            do_query(query, date + " Depts posting to the wrong item type  " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_MISC_TOTAL_" + year):
            do_query(query, date + " IT Dept Misc Awards Total (10001) " + year + ".xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_MISC_NOPOST_20" + year) \
                | query.startswith("UUFA_MISC_NOPOST_" + year):
            do_query(query, date + " 7880013 Dept Misc Awards No Post " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_BOOKS_TOTAL_" + year):
            do_query(query, date + " 788 IT Dept Books Awards Total (10002) " + year + ".xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_BOOKS_NOPOST_20" + year) \
                | query.startswith("UUFA_BOOKS_NOPOST_" + year):
            do_query(query, date + " 7880015 Dept Books No Post " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_ROOMBOARD_TOTAL_" + year):
            do_query(query, date + " 788 IT Dept Room & Board Awards Total (10003) " + year + ".xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_ROOMBOARD_NOPOST_20" + year) \
                | query.startswith("UUFA_ROOMBOARD_NOPOST_" + year):
            do_query(query, date + " 7880029 Dept Room & Board No Post " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_TRAVEL_TOTAL_" + year):
            do_query(query, date + " 788 IT Dept Travel Awards Total (10004) " + year + ".xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_TRAVEL_NOPOST_20" + year) \
                | query.startswith("UUFA_TRAVEL_NOPOST_" + year):
            do_query(query, date + " 7880033 Dept Travel No Post " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_TRAINEESHIPS_TOTAL_" + year) \
                | query.startswith("UUFA_TRAINEESHIP_TOTAL_" + year):
            do_query(query, date + " 788 IT Dept Traineeship Awards Total (10005) " + year + ".xlsx", directory,
                     null_attachment_list)
        if query.startswith("UUFA_TRAINEESHIP_NOPOST_20" + year) \
                | query.startswith("UUFA_TRAINEESHIP_NOPOST_" + year):
            do_query(query, date + " 7880034 Dept Traineeship No Post " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_SCH_ALL_NEED_20" + year):
            do_query(query, date + " All Scholarships Need Based " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_SCH_ALL_NRFRESH_20" + year):
            do_query(query, date + " All Scholarships Non Res Freshman " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_SCH_ALL_NRTRAN_20" + year):
            do_query(query, date + " All Scholarships Non Res Transfer " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_SCH_ALL_RESFRESH_20" + year):
            do_query(query, date + " All Scholarships Res Freshman " + year + ".xlsx", directory,
                     ssj_attachment_list)
        if query.startswith("UUFA_SCH_ALL_RESTRAN_20" + year):
            do_query(query, date + " All Scholarships Res Transfer " + year + ".xlsx", directory,
                     ssj_attachment_list)

    if ssj_attachment_list:
        mailer("", aid_year + " Weekly Scholarship Queries", ssj_mail, "", ssj_attachment_list)
        del ssj_attachment_list[:]


for filename in os.listdir("."):
    # Daily Queries
    if filename.startswith("FA_IL_CMT_CODE_OVR_AGR_") \
            | filename.startswith("UUFA_IL_CMT_CDE_OVR_AGR"):
        do_dailies()
    # Monday Weekly Queries
    if filename.startswith("FA_WR_AID_DISB_NO_ENRLD_ATH_") \
            | filename.startswith("UUFA_WR_AID_DISB_NO_ENR_ATH_"):
        do_monday_weeklies()
    # Budget Queries
    if filename.startswith("FA_BR_ACAD_LVLS_OUT_OF_SYNC") \
            | filename.startswith("UUFA_BR_ACAD_LVLS_OUT_SYNC"):
        do_budget_queries()
    # Packaging Queries
    if filename.startswith("FA_PRT_ACAD_PROG_REVIEW") \
            | filename.startswith("UUFA_PRT_ACAD_PROG_REVIEW"):
        do_packaging_queries()
    # Monthly Queries
    if filename.startswith("FA_MR_ATHLETE_RESIDENCY_") \
            | filename.startswith("UUFA_MR_ATHLETE_RESIDENCY_"):
        do_monthlies()
    # Disbursement Queries
    if filename.startswith("FA_DQ_AUTHORIZED_NOT_DISB") \
            | filename.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB"):
        do_disb_queries()
    # 2nd LDR Queries
    if filename.startswith("FA_HRS") \
            | filename.startswith("UUFA_HRS_DECR"):
        do_2nd_ldr()
    # End of Term Queries
    if filename.startswith("FA_ACADEMIC_PLAN_RVW_FRAP") \
            | filename.startswith("UUFA_ACADEMIC_PLAN_RVW_FRAP_"):
        do_end_of_term_queries()
    # Day After LDR Queries
    if filename.startswith("FA_LDR_MIN") \
            | filename.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_"):
        do_day_after_ldr()
    # Direct Loans Pre-Outbound Queries
    if filename.startswith("FA_DLR_ENTRANCE_COUNSELING") \
            | filename.startswith("UUFA_DLR_ENTRANCE_COUNSEL"):
        dl_pre_outbound()
    # Alternative Loan Pre-Outbound Queries
    if filename.startswith("UUFA_ALR_LOAN_ORG_LND_NT"):
        al_pre_outbound()
    # Pre-Repackaging Queries
    if filename.startswith("UUFA_PP_"):
        do_pre_repackaging()
    # Mid-Repackaging Queries
    if filename.startswith("UUFA_MP_"):
        do_mid_repack_queries()
    # After Repackaging Queries
    if filename.startswith("UUFA_AP_"):
        do_after_repackaging()
    # Daily Scholarships Queries
    if filename.startswith("UUFA_SCHOLAR"):
        do_daily_scholarships()
    # Weekly Scholarships Queries
    if filename.startswith("UUFA_SCH_"):
        do_weekly_scholarships()

        # TEMPLATE
        # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
        # will be added.
        # for query in os.listdir("."):
        # if query.startswith("____________________"):
        #        do_query(query, date + " ________________" + year + ".xlsx", directory,
        #                 lkj_attachment_list)

        # if ak_attachment_list:
        #    mailer("", aid_year + " _____________", ak_mail, "", ak_attachment_list)
        #   del ak_attachment_list[:]