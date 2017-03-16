__author__ = 'mmason'
#Version 4.30

import os
import time
import shutil
import re

# date becomes the current date and is then placed in MM-DD-YY format
date = time.strftime("%x").replace("/", "-")
month_folder = date[:2] + "-20" + date[-2:]
year = date[-2:]
###############################
test = True
###############################

def rename(name, new_name, attach_list, i=2):
    this_name = os.path.realpath(name)
    this_new_name = os.path.realpath(new_name)
    this_attach_list = attach_list
    num = i
    try:
        os.rename(this_name, this_new_name)
        this_attach_list.append(this_new_name)
    except WindowsError:
        try:
            final_name = this_new_name[:-4] + " (" + str(num) + ")" + this_new_name[-4:]
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
        print 'Already a file with the name:'  + name +  'at location.'


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


def send_to(*args):
    addresses = ""
    for i in args:
        addresses = addresses + i + ";"
    return addresses


# region Email and Attachment Groups008/
amber =     "acook@sa.utah.edu"
brenda =    "bburke@sa.utah.edu"
carol =     "cbergstrom@sa.utah.edu"
chelsea =   "cspringer@sa.utah.edu"
emily =     "ehereth@sa.utah.edu"
hilerie =   "hilerie.harris@sa.utah.edu"
jen =       "jberry@sa.utah.edu"
jessica =   "jlykins@sa.utah.edu"
john =      "karen.henriquez@utah.edu"
jonathan =  "jleon@sa.utah.edu"
karen =     "karen.henriquez@utah.edu"
kayla =     "kmccloyn@sa.utah.edu"
krista =    "kburton@sa.utah.edu"
leila =     ""
leo =       "lgaray@sa.utah.edu"
linh =      "lly@sa.utah.edu"
lisa =      "lisa.zaelit@admin.utah.edu"
marc =      "mgangwer@sa.utah.edu; TDespain@sa.utah.edu"
mary =      "msnow@sa.utah.edu"
mat =       "mmason@sa.utah.edu"
raenetta =  "rking@sa.utah.edu"
ryan =      "rchristensen@sa.utah.edu"
scott =     ""
shana =     "syem@sa.utah.edu"
shanon =    "SYBrown@sa.utah.edu"
sheryl =    "shansen@sa.utah.edu"
steffany =  "steffany.forrest@income.utah.edu"
veronica =  "vchristensen@sa.utah.edu"
anne =       john + ";" + emily + ";" + mary

athletics = chelsea + ";" + kayla
loans =     krista + ";" + jessica
prof =     shana
systems =   mat + ";" + leo + ";" + jen + ";" + veronica

a_mail =        send_to(anne)
ac_mail =       send_to(amber, carol)
acl_mail =      send_to(amber, carol, leila)
aca_mail =      send_to(amber, carol, ryan)
acj_mail =      send_to(amber, carol, jonathan)
acrkm_mail =    send_to(amber, carol, ryan, karen, marc)
acvj_mail =     send_to(amber, carol, veronica, jen)
acvjm_mail =    send_to(amber, carol, veronica, jen, mat)
aj_mail =       send_to(amber, jonathan)
ak_mail =       send_to(ryan, karen)
aka_mail =      send_to(ryan, karen, anne)
akc_mail =      send_to(ryan, karen, carol)
akcal_mail =    send_to(ryan, karen, carol, amber, leila)
akl_mail =      send_to(ryan, karen, linh)
akv_mail =      send_to(ryan, karen, veronica)
alt_mail =      send_to(loans, ryan, systems)
amber_k_mail =  send_to(amber, karen)
amber_mail =    send_to(amber)
athletics_mail =send_to(athletics, karen)
disb_mail =     send_to(loans, ryan, karen, marc, steffany, chelsea, amber, carol, kayla, lisa, systems)
dl_mail =       send_to(loans, ryan, karen, systems)
jen_mail =      send_to(jen)
krms_mail =     send_to(krista, ryan, mary, shanon)
leo_mail =      send_to(leo)
loans_ak_mail = send_to(loans, ryan, karen)
loans_kr_mail = send_to(loans, ryan, karen)
loans_krv_mail =send_to(loans, ryan, karen, veronica)
loans_r_mail =  send_to(loans, ryan)
loans_rc_mail = send_to(prof, ryan, carol)
mat_mail =      send_to(mat)
ms_mail =       send_to(mat, scott)
mary_s_mail =   send_to(mary, shanon)
null_mail =     ""
prof_mail =     send_to(prof, ryan)
prof_rkm_mail = send_to(ryan, karen, prof, marc)
rac_mail =      send_to(raenetta, ryan, carol)
rkam_mail =     send_to(ryan, karen, anne, marc)
rkjm_mail =     send_to(ryan, karen, john, marc)
rkm_mail =      send_to(ryan, karen, marc)
rkmv_mail =     send_to(ryan, karen, marc, veronica)
scott_mail =    send_to(scott)
ss_mail =       send_to(amber, carol, jonathan, mary, sheryl, systems)
ssj_mail =      send_to(amber, carol, jonathan, mary, shanon, chelsea, brenda, hilerie, raenetta)
sys_mail =      send_to(systems)
v_mail =        send_to(veronica)
vm_mail =       send_to(veronica, mat)
vs_mail =       send_to(veronica, scott)

a_attachment_list = []
ac_attachment_list = []
acl_attachment_list = []
aca_attachment_list = []
acrkm_attachment_list = []
acj_attachment_list = []
ac_attachment_list = []
acvj_attachment_list = []
acvjm_attachment_list = []
aj_attachment_list = []
ak_attachment_list = []
aka_attachment_list = []
rkam_attachment_list = []
akc_attachment_list = []
akcal_attachment_list = []
akl_attachment_list = []
rkm_attachment_list = []
prof_rkm_attachment_list = []
rkjm_attachment_list = []
rkmv_attachment_list = []
akv_attachment_list = []
alt_attachment_list = []
amber_attachment_list = []
amber_k_attachment_list = []
athletics_attachment_list = []
disb_attachment_list = []
dl_attachment_list = []
jen_attachment_list = []
krms_attachment_list = []
leo_attachment_list = []
loans_r_attachment_list = []
loans_ak_attachment_list = []
loans_kr_attachment_list = []
loans_krv_attachment_list = []
loans_rc_attachment_list = []
mat_attachment_list = []
ms_attachment_list = []
mary_s_attachment_list = []
null_attachment_list = []
prof_attachment_list = []
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
    year = date[:2]
    for query_name in os.listdir("."):
        if query_name.startswith("FA_IL_CMT_CODE_OVR_AGR_") | query_name.startswith("UUFA_IL_CMT_CDE_OVR_AGR"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Daily', aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Daily', aid_year, month_folder))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_IL_ATHLETE_OVERAWARD_"):
            do_query(query, date + " Athlete Aid Overaward " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_IL_CMT_CDE_OVR_AGR_LMT_") and (year in query[:-10]) :
            do_query(query, date + " Comment Code Over Aggregate 20" + year + ".xls", directory,
                     loans_kr_attachment_list)

        if query.startswith("UUFA_IL_COMMENT_CODE_298_") and (year in query[:-10]) :
            do_query(query, date + " IASG - Pell Eligible 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FDEG_FFBD_FBLK_FFBC_") and (year in query[:-10]) :
            do_query(query, date + " Complete FDEG FFBD FBLK FFBC " + year + ".xls", directory,
                     prof_rkm_attachment_list)

        if query.startswith("FA_IL_COMPLETE_FDEG_") and (year in query[:-10]) :
            do_query(query, date + " FDEG Update 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_CORR_NOT_MARK_SENT_") and (year in query[:-10]) :
            do_query(query, date + " Corrections not Marked to Sent 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_CORR_SENT_RJCT_CD1_") and (year in query[:-10]) :
            do_query(query, date + " Correction Sent Reject Code 1 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_IL_ENRL_GR_DATE_ERRORS_") and (year in query[:-10]) :
            do_query(query, date + " Place FDIP" + year + " Checklist 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FBKP") and (year in query[:-10]) :
            do_query(query, date + " FBKP" + year + " Checklist Initiated.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FOUT") and (year in query[:-10]) :
            do_query(query, date + " Outside Resources 20" + year + ".xls", directory,
                     rac_attachment_list)

        if query.startswith("FA_IL_FP1B") and (year in query[:-10]) :
            do_query(query, date + " FP1B" + year + " Checklist " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_IL_FP2B") and (year in query[:-10]) :
            do_query(query, date + " FP2B" + year + " Checklist " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_IL_FP1N") and (year in query[:-10]) :
            do_query(query, date + " FP1N" + year + " Checklist " + year + ".xls", directory,
                     rac_attachment_list)

        if query.startswith("FA_IL_FP2N") and (year in query[:-10]) :
            do_query(query, date + " FP2N" + year + " Checklist " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FPJ") and (year in query[:-10]) :
            do_query(query, date + " FPJ" + year + " Checklist 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_ISIR_02_IND_UP_DWN_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Service IND UP Down 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FEHU") and (year in query[:-10]) :
            do_query(query, date + " Initiated FEHU" + year + " Checklist.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_IL_IRS_DRT_02_20") and (year in query[:-10]) :
            do_query(query, date + " IRS Data Retrieval Equal to 02 " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_IL_ISIR_CMT_CODE_359_360_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Comment Code 359 or 360 " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_IL_ISIR_GRD_I_UG_FATRM_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Graduate Independent UG FATERM 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_ISIR_PRMARY_EFC_DIF_") and (year in query[:-10]) :
            do_query(query, date + " Primary EFC Difference 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_ISIR_LOADED_NOT_PKG_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Loaded Not Packaged 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_LN_EXP_GRAD_DATE") and (year in query[:-10]) :
            do_query(query, date + " Loan Awd No FLRP Grad Date 20" + year + ".xls", directory,
                     loans_kr_attachment_list)

        if query.startswith("UUFA_IL_NURSING_LOANS_TILA_") and (year in query[:-10]) :
            do_query(query, date + " Nursing Loans 20" + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("FA_IL_OTHER_ATB_20") and (year in query[:-10]) :
            do_query(query, date + " ISIR Other ATB 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_IL_OTHER_ATTND_") and (year in query[:-10]) :
            do_query(query, date + " Attend Other Institution 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_PELL_LEU_C_") and (year in query[:-10]) :
            do_query(query, date + " Pell LEU Limit Flag C " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_PELL_LEU_E_") and (year in query[:-10]) :
            do_query(query, date + " Pell LEU Limit Flag E " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_PELL_MAX_ELIG_") and (year in query[:-10]) :
            do_query(query, date + " Pell Max Eligibility " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_IL_SF_RFND_AWD_NO_POST_") and (year in query[:-10]) :
            do_query(query, date + " Refund Post Third Party 20" + year + ".xls", directory,
                     rac_attachment_list)

        if query.startswith("UUFA_IL_SUB_ISIR_NO_PACKAGE_") and (year in query[:-10]) :
            do_query(query, date + " Subsequent ISIR Not Package Not Verified 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_VET_ACTV_DUTY_STAT_") and (year in query[:-10]) :
            do_query(query, date + " Veteran Active Duty Status 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_VER_I_SUB_SUSP_ISIR_") and (year in query[:-10]) :
            do_query(query, date + " FAVR Initiated Susp ISIR Psbl DRT " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_V5_VER_AFTR_OTH_VER_") and (year in query[:-10]) :
            do_query(query, date + " Selected for V5 after other Ver " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_UPDATED_SEC_") and (year in query[:-10]) :
            do_query(query, date + " New ISIR Updated ATB " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FATN_INITIATED_") and (year in query[:-10]) :
            do_query(query, date + " Review FATN Checklist 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FED_AID_OVERAWARD_") and (year in query[:-10]) :
            do_query(query, date + " Federal Aid Overaward " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FHST_I_HST_COMPLETE_") and (year in query[:-10]) :
            do_query(query, date + " HS Transcript 'C' FHST" + year + " I.xls", directory,
                     rkm_attachment_list)

        if year == "15" and query.startswith("FA_IL_SW_THESIS_HOURS"):
            do_query(query, date + " SW Thesis Hours.xls", directory,
                     ak_attachment_list)

        if query.startswith("ussf0034"):
            do_query(query, date + " " + query, directory,
                     akcal_attachment_list)

        if query.startswith("UUFA_IL_PKG_SCH_EXP_GRAD_FA_") and (year in query[:-10]) :
            do_query(query, date + " Scholarship Aid Grad Date Fall 20" + year + ".xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_IL_PKG_FED_EXP_GRAD_FA_") and (year in query[:-10]) :
            do_query(query, date + " Accepted Federal Aid Grad Date Fall 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FPEL") and (year in query[:-10]) :
            do_query(query, date + " FPEL" + year + " No Database Match.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PACKAGING_C_NO_FED_AID_") and (year in query[:-10]) :
            do_query(query, date + " Packaging C 71%-72% No Fed Aid " + year + ".xls", directory,
                     sys_attachment_list)

        if query.startswith("UUFA_PERKINS_NOT_DISBURSED"):
            do_query(query, date + " Perkins Not Disbursed " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_ITM_TYPE_NOT_IN_NEXT_AY"):
            do_query(query, date + " Item Types in " + str(int(year)-1) + "not in " + year + ".xls", directory,
                     mat_attachment_list)

        if query.startswith("UUFA_IL_FREV_CHECKLIST_INT_") and (year in query[:-10]) :
            do_query(query, date + " FREV Checklist Initiated " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_IL_FREV_DOES_NOT_EXIST_") and (year in query[:-10]) :
            do_query(query, date + " FREV Does Not Exist 20" + year + ".xls", directory,
                     rkm_attachment_list)

    if ac_attachment_list:
        mailer("", aid_year + " Daily Queries", ac_mail, "", ac_attachment_list)
        del ac_attachment_list[:]
    if ak_attachment_list:
        mailer("", aid_year + " Daily Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if akc_attachment_list:
        mailer("", aid_year + " Daily Queries", akc_mail, "", akc_attachment_list)
        del akc_attachment_list[:]
    if akcal_attachment_list:
        mailer("", aid_year + " Daily Queries", akcal_mail, "", akcal_attachment_list)
        del akcal_attachment_list[:]
    if rkm_attachment_list:
        mailer("", aid_year + " Daily Queries", prof_rkm_mail, "", prof_rkm_attachment_list)
        del prof_rkm_attachment_list[:]
    if rkm_attachment_list:
        mailer("", aid_year + " Daily Queries", rkm_mail, "", rkm_attachment_list)
        del rkm_attachment_list[:]
    if athletics_attachment_list:
        mailer("", aid_year + " Daily Queries", athletics_mail, "", athletics_attachment_list)
        del athletics_attachment_list[:]
    if loans_r_attachment_list:
        mailer("", aid_year + " Daily Queries", loans_r_mail, "", loans_r_attachment_list)
        del loans_r_attachment_list[:]
    if loans_kr_attachment_list:
        mailer("", aid_year + " Daily Queries", loans_kr_mail, "", loans_kr_attachment_list)
        del loans_kr_attachment_list[:]
    if mat_attachment_list:
        mailer("", aid_year + " Daily Queries", mat_mail, "", mat_attachment_list)
        del mat_attachment_list[:]
    if rac_attachment_list:
        mailer("", aid_year + " Daily Queries", rac_mail, "", rac_attachment_list)
        del rac_attachment_list[:]
    if sys_attachment_list:
        mailer("", aid_year + " Daily Queries", sys_mail, "", sys_attachment_list)
        del sys_attachment_list[:]


def do_monday_weeklies():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_WR_AID_DISB_NO_ENR_ATH_"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Monday Weekly', aid_year, month_folder))
        directory2014 = os.path.realpath(os.path.join('C:\Testing Bob\Monday Weekly', "2013-2014", month_folder))
        directory2015 = os.path.realpath(os.path.join('C:\Testing Bob\Monday Weekly', "2014-2015", month_folder))
        packaging_directory = os.path.realpath(os.path.join('C:\Testing Bob/Packaging', aid_year, month_folder))
        disb_failure_directory = os.path.realpath(os.path.join('C:\Testing Bob/Disb Failure ' + aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monday Weekly', aid_year, month_folder))
        directory2014 = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monday Weekly', "2013-2014", month_folder))
        directory2015 = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monday Weekly', "2014-2015", month_folder))
        packaging_directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Packaging', aid_year, month_folder))
        disb_failure_directory = os.path.realpath(os.path.join('O:/Disbursement Failure/Disb Failure ' + aid_year))

    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(directory2014):
        os.makedirs(directory2014)
    if not os.path.isdir(directory2015):
        os.makedirs(directory2015)
    if not os.path.isdir(packaging_directory):
        os.makedirs(packaging_directory)
    if not os.path.isdir(disb_failure_directory):
        os.makedirs(disb_failure_directory)

    # Change File_Name to be query ac it is received and _new_file_name to what the new query should be.Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_WR_AGG_CK_MLT_YR_AWDED_" ) and (year in query[:-10]) :
            do_query(query, date + " Student Pkgd for " + str(int(year) - 1) + " after " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_AID_DISB_NO_ENR_ATH_" ) and (year in query[:-10]) :
            do_query(query, date + " Athlete Disb Not Enrolled " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_WR_AID_DISB_NO_ENR_FED_" ) and (year in query[:-10]) :
            do_query(query, date + " Federal Disb Not Enrolled " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_AID_DISB_NO_ENR_SCH_" ) and (year in query[:-10]) :
            do_query(query, date + " T 53 Sch Disb Not Enrolled " + year + ".xls", directory,
                     amber_attachment_list)

        if query.startswith("UUFA_WR_AMERICORP_AWD_POST_" ) and (year in query[:-10]) :
            do_query(query, date + " Americorp Awards " + year + ".xls", directory,
                     amber_k_attachment_list)

        if query.startswith("UUFA_WR_ATHLETE_NOT_DISB_" ) and (year in query[:-10]) :
            do_query(query, date + " Athlete Not Disbursed " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_WR_ATH_HRS_AFTR_CENSUS_" ) and (year in query[:-10]) :
            do_query(query, date + " Ath Hours After Census " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_WR_ATH_SF_TERM_BALANCE_" ) and (year in query[:-10]) :
            do_query(query, date + " Athlete Tuition Fee Balance " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_WR_AUDIT_CLSS_AID_DISB_" ) and (year in query[:-10]) :
            do_query(query, date + " Audit Class Aid Disbursed " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_WR_AWD_UG_NOW_GRAD_ATH_" ) and (year in query[:-10]) :
            do_query(query, date + " Ath Awards past Grad Term " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_WR_AWD_UG_NOW_GRAD_FC_" ) and (year in query[:-10]) :
            do_query(query, date + " Federal Awards past Grad Term " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_AWD_UG_NOW_GRAD_SV_" ) and (year in query[:-10]) :
            do_query(query, date + " Scholar Awards past Grad Term " + year + ".xls", directory,
                     amber_attachment_list)

        if query.startswith("UUFA_WR_CHKLST_STATUS_ERROR_" ) and (year in query[:-10]) :
            do_query(query, date + " Checklist Status Error " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_DISB_ATH_FAILURE_" ) and (year in query[:-10]) :
            do_query(query, date + " Authorization Failure 20" + year + ".xls", disb_failure_directory,
                     null_attachment_list)

        if query.startswith("UUFA_WR_DL_DISBURSED_LTHT_" ) and (year in query[:-10]) :
            do_query(query, date + " DL Disbursed LTHT " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_WR_DL_EC_SUSPENDED_" ) and (year in query[:-10]) :
            do_query(query, date + " DL Entrance Counseling Suspense " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_DL_ORIG_TRNS_PEND_" ) and (year in query[:-10]) :
            do_query(query, date + " DL Orig Trans Pending " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_EFT_CONSENT_VERIF"):
            do_query(query, date + " EFT Consent Verification 20" + year + ".xls", directory,
                     rkjm_attachment_list)

        if query.startswith("UUFA_WR_FAFSA_CKLST_INCMP_" ) and (year in query[:-10]) :
            do_query(query, date + " PLUS FAFSA Incomplete " + year + ".xls", directory,
                     rkm_attachment_list)

        if  query.startswith("UUFA_WR_FALL_TOTAL_WDRN_DRP_" ) and (year in query[:-10]) :
            do_query(query, date + " Fall Disb Total Withdrawn Drop " + year + ".xls", directory,
                     rkm_attachment_list)

        if  query.startswith("UUFA_WR_SPR_TOTAL_WDRN_DRP_" ) and (year in query[:-10]) :
            do_query(query, date + " Spring Disb Total Withdrawn Drop " + year + ".xls", directory,
                     rkm_attachment_list)

        if  query.startswith("UUFA_WR_SUM_TOTAL_WDRN_DRP_" ) and (year in query[:-10]) :
            do_query(query, date + " Summer Disb Total Withdrawn Drop " + year + " .xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_FARC_CHECKLIST_" ) and (year in query[:-10]) :
            do_query(query, date + " FARC 30 Day Review " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_FARC_CMNT_CODES_" ) and (year in query[:-10]) :
            do_query(query, date + " Initiated FARC w ISIR Cmnt Codes " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_WR_FED_AID_OVERAWARD_" ) and (year in query[:-10]) :
            do_query(query, date + " Federal Aid Overaward " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_FGED_ISIR_DEGREE_" ) and (year in query[:-10]) :
            do_query(query, date + " FGED ISIR Degree 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_FPEL" + year + "_INITIATED_AWDED"):
            do_query(query, date + " FPEL" + year + " Initiated Pell.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_WR_GENDER_") and (year in query[:-10]) :
            do_query(query, date + " Gender Discrepancies 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_HEDU_PARAMEDIC_") and (year in query[:-10]) :
            do_query(query, date + " HEDU Paramedic Class 20" + year + "F.xls", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_HOME_SCHOOLED_") and (year in query[:-10]) :
            do_query(query, date + " Home Schooled Check " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_HRS_DECREASE_ATH_") and (year in query[:-10]) :
            do_query(query, date + " Hours Decrease Athlete " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_WR_HRS_DECREASE_FC_") and (year in query[:-10]) :
            do_query(query, date + " Hours Decrease FC " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_HRS_DECREASE_SV_") and (year in query[:-10]) :
            do_query(query, date + " Hours Decrease SV " + year + ".xls", directory,
                     amber_attachment_list)

        if query.startswith("UUFA_WR_FHST_I_HST_COMPLETE_") and (year in query[:-10]) :
            do_query(query, date + " HS Transcript 'C' FHST" + year + " 'I'.xls", directory,
                     rkm_attachment_list)
        
        if query.startswith("UUFA_WR_FMFB") and (year in query[:-10]) :
            do_query(query, date + " Football Athlete with Pell Award " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_WR_ISIR_AS_EFC_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Assumption EFC 20" + year + ".xls", directory,
                     mat_attachment_list)

        if query.startswith("UUFA_WR_ISIR_COR_ASSESSMENT_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Correction Assessment " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_ISIR_CORR_REJECT_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Correction Rejected " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_ISIR_DGR_ANSW_CHNG_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Degree Answer Change " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_ISIR_DEP_STAT_PRB_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Dependency 20" + year + ".xls", directory,
                     mat_attachment_list)

        if query.startswith("UUFA_WR_ISIR_REJECTED_CORR_") and (year in query[:-10]) :
            do_query(query, date + " ISIR Rejected Corrections 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_ISIR_REJECT_CODES_" + year) \
                | query.startswith("UUFA_WR_ISIR_REJECT_CODES_20") and (year in query[:-10]) :
            do_query(query, date + " Rejected ISIR's 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_ISIR_SS_MCH_NOT_CON_") and (year in query[:-10]) :
            do_query(query, date + " SS Match Not Confirmed 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_ISIR_SUSPENSE_20") and (year in query[:-10]) :
            do_query(query, date + " ISIR Suspense " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_LEGAL_ALIEN_WORK_") and (year in query[:-10]) :
            do_query(query, date + " Legal Alien Work 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_LN_ACCPT_STAF_31_32_") and (year in query[:-10]) :
            do_query(query, date + " Stafford Accept Offer " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_LOAN_CENSUS_DATE_") and (year in query[:-10]) :
            do_query(query, date + " Loans Census Date 20" + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_LN_FA907_1_REVISED_") and (year in query[:-10]) :
            do_query(query, date + " Loan Disbursed Report " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_LN_FA907_2_REVISED_") and (year in query[:-10]) :
            do_query(query, date + " Loan Not Disbursed Report " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("FA_WR_LOAN_ORIG_DEPT_REVIEW_") and (year in query[:-10]) :
            do_query(query, date + " Loan ORIG DEPT RVW " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_LN_SENT_NO_RESPONSE_") and (year in query[:-10]) :
            do_query(query, date + " Loan Sent No Response " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_LOAN_TRANSMIT_HOLD_") and (year in query[:-10]) :
            do_query(query, date + " Loan Transmit Hold " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_LW_MD_DN_AW_NO_DISB_") and (year in query[:-10]) :
            do_query(query, date + " LW MD DN Awards Not Disbursed " + year + ".xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_WR_MNTGMR_AMCORP_OVRAW_") and (year in query[:-10]) :
            do_query(query, date + " Montgomery Americorp Overaward " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_MULTIPLE_EMPLIDS_") and (year in query[:-10]) :
            do_query(query, date + " Multiple EMPLIDS 20" + year + ".xls", directory,
                     mat_attachment_list)

        if query.startswith("UUFA_WR_NO_COMMENT_CODE_") and (year in query[:-10]) :
            do_query(query, date + " Sub ISIR Checklist No ISIR Comment Code 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_NSLDS_LOAN_DATA_") and (year in query[:-10]) :
            do_query(query, date + " NSLDS Loan Data .xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("UUFA_WR_OVRD_ACAD_LVL_") and (year in query[:-10]) :
            do_query(query, date + " FA Term Override Acad Level " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PA_EXPECT") and (year in query[:-10]) :
            do_query(query, date + " PA MPS FDEG Checklist " + year + ".xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_WR_PA_FDEG_CHECKLIST"):
            do_query(query, date + " PA Expected Grad Date Blank 20" + year + "U.xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_WR_PELL_AWRD_LOCK_") and (year in query[:-10]) :
            do_query(query, date + " Pell Award Lock No FPEL" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PELL_OVERPAYMENT_") and (year in query[:-10]) :
            do_query(query, date + " Pell Ovpy Check NSLDS 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PELL_SUMMER_NO_PELL_") and (year in query[:-10]) :
            do_query(query, date + " Pell Summer No Pell 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PELL_TERM_FT_") and (year in query[:-10]) :
            do_query(query, date + " Term Pell Awards FT 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PELL_TERM_HT_") and (year in query[:-10]) :
            do_query(query, date + " Term Pell Awards HT 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PELL_TERM_LH_") and (year in query[:-10]) :
            do_query(query, date + " Term Pell Awards LH 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PELL_TERM_NL_") and (year in query[:-10]) :
            do_query(query, date + " Term Pell Awards NL 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PELL_TERM_TQ_") and (year in query[:-10]) :
            do_query(query, date + " Term Pell Awards TQ 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PERK_SPLIT_MISMATCH_") and (year in query[:-10]) :
            do_query(query, date + " Perkins Plan Split Mismatch " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_PRK_LN_ACAD_LVL_CHG_") and (year in query[:-10]) :
            do_query(query, date + " Perkins Awd With Acad Lvl Change " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_QUALITY_ASSURANCE_") and (year in query[:-10]) :
            do_query(query, date + " QA Students Complete Verification 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_WR_RT4_DROPPED_CLASSES_") and (year in query[:-10]) :
            do_query(query, date + " RT4 Dropped Classes 20" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_WR_SCH_NOT_DISB_") and (year in query[:-10]) :
            do_query(query, date + " Cash Non-Cash Sch Not Disb " + year + ".xls", directory,
                     amber_attachment_list)

        if query.startswith("UUFA_WR_SSR_MATCH_NOT_CNFRM_") and (year in query[:-10]) :
            do_query(query, date + " SSR Not Confirmed 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SSR_NOT_CNFRMD_VTRN_") and (year in query[:-10]) :
            do_query(query, date + " VA Match SSR DB Override " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SS_DB_OVERRIDE_") and (year in query[:-10]) :
            do_query(query, date + " SS DB Override " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SUB_ISIR_PACKAGED_") and (year in query[:-10]) :
            do_query(query, date + " Subsequent ISIR Packaged 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SUB_ISIR_REAWD_AID_") and (year in query[:-10]) :
            do_query(query, date + " Canceled FCOR Complete " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SUB_ISIR_SYSG_20") and (year in query[:-10]) :
            do_query(query, date + " Subsequent ISIR System Generated 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SUB_ISIR_VERIFIED_") and (year in query[:-10]) :
            do_query(query, date + " Subsequent ISIR Verified 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SUMMER_NO_DL_") and (year in query[:-10]) :
            do_query(query, date + " Summer Enroll No DL " + year + ".xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("UUFA_WR_SSP_DOB_PRB_APPLCNT_") and (year in query[:-10]) :
            do_query(query, date + " Suspense Applicant DOB Problem " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SSP_NAME_PRB_APLCNT_") and (year in query[:-10]) :
            do_query(query, date + " Suspense Applicant Name Problem " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_SSP_SSN_PRB_APLCNT_") and (year in query[:-10]) :
            do_query(query, date + " Suspense Applicant SSN Problem " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_TERM_NSLDS_LOAN_YR_") and (year in query[:-10]) :
            do_query(query, date + " NSLDS Loan Year Blank " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_TITLE_VII_MED_LOANS_") and (year in query[:-10]) :
            do_query(query, date + " Title VII Medical Loans TILA 20" + year + ".xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_WR_TRANSFER_ENT_CNS_") and (year in query[:-10]) :
            do_query(query, date + " Transfer Students Entrance Counseling 20" + year + ".xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("UUFA_WR_TRANSFER_STU_FA_SP_") and (year in query[:-10]) :
            do_query(query, date + " Transfer Students Fall-Spring 20" + year + ".xls", directory,
                     leo_attachment_list)

        if query.startswith("UUFA_WR_UG_GR_DIR_LN_GR_TRM_") and (year in query[:-10]) :
            do_query(query, date + " UG-GR Direct Ln Grad Term " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_UG_GR_PLUS_GR_TERM_") and (year in query[:-10]) :
            do_query(query, date + " UG-GR PLUS Grad Term " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_UAC_SNGDO_20") and (year in query[:-10]) :
            do_query(query, date + " UAC SNGDO Campus 20" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_UNDOCUMENTED_STUDENTS_") and (year in query[:-10]) :
            do_query(query, date + " Undocumented Student Awards " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_WR_VERI_CHKLST_MISSING_") and (year in query[:-10]) :
            do_query(query, date + " Verification Checklist Missing 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_VERI_INCOME_ADJ_20") and (year in query[:-10]) :
            do_query(query, date + " Income Adjustments 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_VER_NOT_CONSL_20") and (year in query[:-10]) :
            do_query(query, date + " Verification Not Consolidated 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_VETERAN_ACTIVE_DUTY_") and (year in query[:-10]) :
            do_query(query, date + " Veteran Active Duty 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_VETERAN_NO_QUALIFY_") and (year in query[:-10]) :
            do_query(query, date + " Veteran No Qualify 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_WEEKS_OF_INSTR_FIX_") and (year in query[:-10]) :
            do_query(query, date + " Weeks of Instruction 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_WR_DL_AY_SP_CANCELED_") and (year in query[:-10]) :
            do_query(query, date + " DL AY SP Cancelled " + year + ".xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("FA_WR_LOAN_TRANSMIT_HOLD_13"):
            do_query(query, date + " Loan Transmit Hold 13.xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_REJECT_CODE_ON_ISIR_") and (year in query[:-10]) :
            do_query(query, date + " Rejected ISIR's " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_RT4_FA_DROP_CLASSES_") and (year in query[:-10]) :
            do_query(query, date + " RT4 Fall Drop Classes 20" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_WR_RT4_SP_DROP_CLASSES_") and (year in query[:-10]) :
            do_query(query, date + " RT4 Spring Drop Classes 20" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_WR_RT4_SU_DROP_CLASSES_") and (year in query[:-10]) :
            do_query(query, date + " RT4 Summer Drop Classes 20" + year + ".xls", directory,
                     ak_attachment_list)

        # Manually run Queries
        if query.startswith("UUFA_WR_LOAN_EFT_DETAIL_ERROR"):
            do_query(query, date + " Loan EFT Detail Error.xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("UUFA_WR_NSL_PROMISSORY_NOTE_") and (year in query[:-10]) :
            do_query(query, date + " NSL Promissory Note " + year + ".xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_WR_SNGDO_CAMPUS") and (year in query[:-10]) :
            do_query(query, date + " Asian-SNGDO Campus " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("FA_WR_FALL_TOTAL_WTHDRN_DRP_13"):
            do_query(query, date + " Fall Total Withdrawn Drop 13.xls", directory2013,
                     rkm_attachment_list)

        if query.startswith("FA_WR_SPR_TOTAL_WTHDRN_DRP_13"):
            do_query(query, date + " Spring Total Withdrawn Drop 13.xls", directory2013,
                     rkm_attachment_list)

        if query.startswith("FA_WR_SUM_TOTAL_WTHDRN_DRP_13"):
            do_query(query, date + " Summer Total Withdrawn Drop 13.xls", directory2013,
                     rkm_attachment_list)

        if query.startswith("FA_WR_SUMMER_NO_DL_13"):
            do_query(query, date + " Summer Enroll No DL " + year + ".xls", directory2013,
                     ak_attachment_list)
            
        if query.startswith("FA_WR_FALL_TOTAL_WTHDRN_DRP_14"):
            do_query(query, date + " Fall Total Withdrawn Drop 14.xls", directory2014,
                     rkm_attachment_list)

        if query.startswith("FA_WR_SPR_TOTAL_WTHDRN_DRP_14"):
            do_query(query, date + " Spring Total Withdrawn Drop 14.xls", directory2014,
                     rkm_attachment_list)

        if query.startswith("FA_WR_SUM_TOTAL_WTHDRN_DRP_14"):
            do_query(query, date + " Summer Total Withdrawn Drop 14.xls", directory2014,
                     rkm_attachment_list)

        if query.startswith("FA_WR_FALL_TOTAL_WDRN_DRP_15"):
            do_query(query, date + " Fall Total Withdrawn Drop 15.xls", directory2015,
                     rkm_attachment_list)

        if query.startswith("FA_WR_SPR_TOTAL_WDRN_DRP_15"):
            do_query(query, date + " Spring Total Withdrawn Drop 15.xls", directory2015,
                     rkm_attachment_list)

        if query.startswith("FA_WR_SUM_TOTAL_WDRN_DRP_15"):
            do_query(query, date + " Summer Total Withdrawn Drop 15.xls", directory2015,
                     rkm_attachment_list)

        if query.startswith("FA_WR_LOAN_TRANSMIT_HOLD_13"):
            do_query(query, date + " Loan Transmit Hold 13.xls", directory2013,
                     loans_r_attachment_list)

        # Packaging queries that are being manually run.
        if query.startswith("FA_PRT_ATH_ACCEPT_FED_AID_") and (year in query[:-10]) :
            do_query(query, date + " Athlete Accepted Federal Aid " + year + ".xls", packaging_directory,
                     athletics_attachment_list)

        if query.startswith("FA_PRT_ATH_AWD_CBA_GRANT_") and (year in query[:-10]) :
            do_query(query, date + " Athlete Awarded CBA Grant " + year + ".xls", packaging_directory,
                     athletics_attachment_list)

        if query.startswith("FA_PRT_ATHLETE_GRAD_DATE_") and (year in query[:-10]) :
            do_query(query, date + " Athlete Accepted Grad Date " + year + ".xls", packaging_directory,
                     athletics_attachment_list)

        if query.startswith("FA_PRT_ATH_OFFERED_FED_AID_") and (year in query[:-10]) :
            do_query(query, date + " Athlete Offered Federal Aid " + year + ".xls", packaging_directory,
                     athletics_attachment_list)

    if aj_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", aj_mail, "", aj_attachment_list)
        del aj_attachment_list[:]
    if ak_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if rkm_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", rkm_mail, "", rkm_attachment_list)
        del rkm_attachment_list[:]
    if rkjm_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", rkjm_mail, "", rkjm_attachment_list)
        del rkjm_attachment_list[:]
    if amber_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", amber_mail, "", amber_attachment_list)
        del amber_attachment_list[:]
    if amber_k_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", amber_k_mail, "", amber_k_attachment_list)
        del amber_k_attachment_list[:]
    if athletics_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", athletics_mail, "", athletics_attachment_list)
        del athletics_attachment_list[:]
    if leo_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", leo_mail, "", leo_attachment_list)
        del leo_attachment_list[:]
    if loans_r_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", loans_r_mail, "", loans_r_attachment_list)
        del loans_r_attachment_list[:]
    if loans_ak_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", loans_ak_mail, "", loans_ak_attachment_list)
        del loans_ak_attachment_list[:]
    if mat_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", mat_mail, "", mat_attachment_list)
        del mat_attachment_list[:]
    if prof_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", prof_mail, "", prof_attachment_list)
        del prof_attachment_list[:]
    if v_attachment_list:
        mailer("", aid_year + " Monday Weekly Queries", v_mail, "", v_attachment_list)
        del v_attachment_list[:]


def do_budget_queries():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_BR_ACAD_LVLS_OUT_SYNC"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Budgets', aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Budgets', aid_year, month_folder))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be query AC it is received and _new_file_name to what the new query should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_BR_ACAD_LVLS_OUT_SYNC_") and (year in query[:-10]) :
            do_query(query, date + " GR Academic Levels Out of Sync " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_BR_ATH_TUIT_INCR_NR_") and (year in query[:-10]) :
            do_query(query, date + " Athlete Tuition Increase " + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_BR_ATH_TUITION_INCRS_") and (year in query[:-10]) :
            do_query(query, date + " Athlete Tuition Increase Non Resident 20" + year + ".xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_BR_BDGT_DOUBLE_BUDGETS_") and (year in query[:-10]) :
            do_query(query, date + " Double Budget " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_BR_COA_LESS_HT_") and (year in query[:-10]) :
            do_query(query, date + " PELL COA Less Than Half Time Enrollment " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_BR_COA_TUIT_ZERO_") and (year in query[:-10]) :
            do_query(query, date + " COA Tuition Amount Zero " + year + ".xls", directory,
                     akv_attachment_list)

        if query.startswith("UUFA_BR_DN_LW_MD_AID_ATRB_") and (year in query[:-10]) :
            do_query(query, date + " DN-LW-MD Student Aid Career " + year + ".xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_BR_FT_CLASS_OVERRIDES_") and (year in query[:-10]) :
            do_query(query, date + " Class Overrides " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_BR_ISIR_SCHOLARSHIP_") and (year in query[:-10]) :
            do_query(query, date + " Scholarship ISIR Received 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_BR_NO_BUDGET_ATTEND_") and (year in query[:-10]) :
            do_query(query, date + " NO Budget Attend 20" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_BR_NSLDS_NO_MCH_DB_FLG_") and (year in query[:-10]) :
            do_query(query, date + " NSLDS No Match DB Flag " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_BR_OMBAMBA_") and (year in query[:-10]) :
            do_query(query, date + " Academic Plan OMBAMBA " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_BR_PELL_COA_BLANK_") and (year in query[:-10]) :
            do_query(query, date + " PELL COA Blank " + year + ".xls", directory,
                     aka_attachment_list)

        if query.startswith("UUFA_BR_PELL_COA_DOUBLE_") and (year in query[:-10]) :
            do_query(query, date + " PELL COA Double " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("FA_BR_PELL_COA_LESS_HT_20") and (year in query[:-10]) :
            do_query(query, date + " PELL COA Less HT Enrollment " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_BR_PROC_STAT_RVW_STAT_") and (year in query[:-10]) :
            do_query(query, date + " Reset Processing Status to 1 " + year + ".xls", directory,
                     ms_attachment_list)

        if query.startswith("UUFA_BR_RES_NON_RES_BDGT_") and (year in query[:-10]) :
            do_query(query, date + " Resident - Non-Resident Budget " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_BR_SCH_TUITION_FEES_NR_") and (year in query[:-10]) :
            do_query(query, date + " Waiver-Scholar Tuition Fees NR " + year + ".xls", directory,
                     amber_attachment_list)

        if query.startswith("UUFA_BR_SCH_TUITION_ONLY_NR_") and (year in query[:-10]) :
            do_query(query, date + " Waiver-Scholar Tuition Only NR " + year + ".xls", directory,
                     amber_attachment_list)

        if query.startswith("UUFA_BR_SCHOLAR_TUIT_FEES_") and (year in query[:-10]) :
            do_query(query, date + " Waiver-Scholar Tuition Fees Res " + year + ".xls", directory,
                     amber_attachment_list)

        if query.startswith("UUFA_BR_SCHOLAR_TUIT_ONLY_") and (year in query[:-10]) :
            do_query(query, date + " Waiver-Scholar Tuition Only Res " + year + ".xls", directory,
                     amber_attachment_list)

        if query.startswith("FA_BR_UFORM_CHANGE_BUD_DUR_") and (year in query[:-10]) :
            do_query(query, date + " Correct Budget Duration " + year + ".xls", directory,
                     rkm_attachment_list)

    if ak_attachment_list:
        mailer("", aid_year + " Budget Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if aka_attachment_list:
        mailer("", aid_year + " Budget Queries", aka_mail, "", aka_attachment_list)
        del aka_attachment_list[:]
    if amber_attachment_list:
        mailer("", aid_year + " Budget Queries", amber_mail, "", amber_attachment_list)
        del amber_attachment_list[:]
    if rkm_attachment_list:
        mailer("", aid_year + " Budget Queries", rkm_mail, "", rkm_attachment_list)
        del rkm_attachment_list[:]
    if athletics_attachment_list:
        mailer("", aid_year + " Budget Queries", athletics_mail, "", athletics_attachment_list)
        del athletics_attachment_list[:]
    if akv_attachment_list:
        mailer("", aid_year + " Budget Queries", akv_mail, "", akv_attachment_list)
        del akv_attachment_list[:]
    if ms_attachment_list:
        mailer("", aid_year + " Budget Queries", ms_mail, "", ms_attachment_list)
        del ms_attachment_list[:]
    if prof_attachment_list:
        mailer("", aid_year + " Budget Queries", prof_mail, "", prof_attachment_list)
        del prof_attachment_list[:] 
    

def do_budget_test_queries():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_BUDGET_"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year

    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Budgets', aid_year, month_folder,"Wrong Budget Queries"))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Budgets', aid_year, month_folder, "Wrong Budget Queries"))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be query AC it is received and _new_file_name to what the new query should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_BUDGET_20" + year + "_DN1"):
            do_query(query, date + " Wrong Budget - Dental 1.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_DN2"):
            do_query(query, date + " Wrong Budget - Dental 2.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_DN3"):
            do_query(query, date + " Wrong Budget - Dental 3.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_DN4"):
            do_query(query, date + " Wrong Budget - Dental 4.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_ACCTMAC"):
            do_query(query, date + " Wrong Budget - Accounting Masters.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_ARCHMAR"):
            do_query(query, date + " Wrong Budget - Architect Masters.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_BUSINESS"):
            do_query(query, date + " Wrong Budget - Grad Business.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_COMDIS"):
            do_query(query, date + " Wrong Budget - COMDIS.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_EAEMS"):
            do_query(query, date + " Wrong Budget - EAE.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_EDUCATION"):
            do_query(query, date + " Wrong Budget - Education.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_ED_PSYCH"):
            do_query(query, date + " Wrong Budget - ED Psychology.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_ENGINERING"):
            do_query(query, date + " Wrong Budget - Grad Engineering.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_FINE_ARTS"):
            do_query(query, date + " Wrong Budget - Fine Arts .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_GENERAL_GR"):
            do_query(query, date + " Wrong Budget - General Grad .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_GENETICS"):
            do_query(query, date + " Wrong Budget - Genetics.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_HEALTH"):
            do_query(query, date + " Wrong Budget - Health.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PRO_HEALTH"):
            do_query(query, date + " Wrong Budget - Health Promotion .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_HUMANITIES"):
            do_query(query, date + " Wrong Budget - Humanities.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_MBA_BUADMB"):
            do_query(query, date + " Wrong Budget - MBA - BUADMBA .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_MD_SCIENCE"):
            do_query(query, date + " Wrong Budget - Medical Science .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_MEDICAL"):
            do_query(query, date + " Wrong Budget - MD Graduate .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_NURSING"):
            do_query(query, date + " Wrong Budget - Graduate Nursing .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PA"):
            do_query(query, date + " Wrong Budget - Physician Assistant .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PHARMACY"):
            do_query(query, date + " Wrong Budget - Pharmacy .xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PLANNING"):
            do_query(query, date + " Wrong Budget - Planning.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PMBAMBA"):
            do_query(query, date + " Wrong Budget - PMBAMBA.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PUBLICPOLI"):
            do_query(query, date + " Wrong Budget - Public Policy.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PUBLIC_ADM"):
            do_query(query, date + " Wrong Budget - Public Administration.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_SCIENCE"):
            do_query(query, date + " Wrong Budget - Science.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_SOC_BEHAV"):
            do_query(query, date + " Wrong Budget - Social and Behavioral.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_SW"):
            do_query(query, date + " Wrong Budget - Social Work.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_XMBAMBA"):
            do_query(query, date + " Wrong Budget - XMBAMBA.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_LW1"):
            do_query(query, date + " Wrong Budget - Law 1.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_LW2"):
            do_query(query, date + " Wrong Budget - Law 2.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_LW3"):
            do_query(query, date + " Wrong Budget - Law 3.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_MD1"):
            do_query(query, date + " Wrong Budget - Med 1.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_MD2"):
            do_query(query, date + " Wrong Budget - Med 2.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_MD3"):
            do_query(query, date + " Wrong Budget - Med 3.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_MD4"):
            do_query(query, date + " Wrong Budget - Med 4.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_UNDERGRAD"):
            do_query(query, date + " Wrong Budget - Undergraduate.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_BUSINESS"):
            do_query(query, date + " Wrong Budget - Undergraduate Business.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_BUS_LTHT"):
            do_query(query, date + " Wrong Budget - Undergraduate Business LTHT.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_ENGINERING"):
            do_query(query, date + " Wrong Budget - Undergraduate Engineering.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_ENG_LTHT"):
            do_query(query, date + " Wrong Budget - Undergraduate Engineering LTHT.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_LTHT"):
            do_query(query, date + " Wrong Budget - Undergraduate LTHT.xls", directory,
                     null_attachment_list)
            
        if query.startswith("UUFA_BUDGET_20" + year + "_UG_NURSE_LTHT"):
            do_query(query, date + " Wrong Budget - Undergraduate Nursing LTHT.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_NURSING"):
            do_query(query, date + " Wrong Budget - Undergraduate Nursing.xls", directory,
                     null_attachment_list)


def do_packaging_queries():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_PRT_ACAD_PROG_REVIEW"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Packaging', aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Packaging', aid_year, month_folder))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be query AC it is received and _new_file_name to what the new query should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_PRT_ACAD_LVLS_OUT_SYNC"):
            do_query(query, date + " UG Acad Levels Out of Sync.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_ACAD_PROG_REVIEW_") and (year in query[:-10]) :
            do_query(query, date + " Academic Progress Review.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_ATH_ACCEPT_FED_AID"):
            do_query(query, date + " Athlete Accepted Federal Aid.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_PRT_ATH_AWD_CBA_GRANT"):
            do_query(query, date + " Athlete Awarded CBA Grant.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_PRT_ATH_GRAD_DATE"):
            do_query(query, date + " Athlete Expected Grad Date.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_PRT_ATH_OFFRD_FED_AID_"):
            do_query(query, date + " Athlete Offered Federal Aid.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_PRT_ATH_OFFR_ACCPT_AID_") and (year in query[:-10]) :
            do_query(query, date + " ATH Fed State Inst O-A.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_PRT_AWARD_DATE_HAD_SAT"):
            do_query(query, date + " SAT Hold Date Award Review.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_AWARD_TERM_HAD_SAT"):
            do_query(query, date + " SAT Hold Term Award Review.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_AWARDS_OTHER_INST"):
            do_query(query, date + " Checklist FAOI" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_AWD_CMB_OVR_AG_RVW_"):
            do_query(query, date + " Award Combined Over Aggregate.xls", directory,
                     loans_r_attachment_list)

        if query.startswith("FA_PRT_AWD_MASS_P_NO_AWARDS"):
            do_query(query, date + " Award Mass Packaging No Awards.xls", directory,
                     ms_attachment_list)

        if query.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_") and (year in query[:-10]) :
            do_query(query, date + " PELL ELIGIBLE NO PELL 20" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_AWD_SUB_OVR_AG_RVW"):
            do_query(query, date + " SUB Over Aggregate.xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("UUFA_PRT_CTZN_IND_AWD_NO_LN"):
            do_query(query, date + " LA-wk eligible - Award No Loans.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PRT_DEFR_ENROLLMENT_") and (year in query[:-10]) :
            do_query(query, date + " DEFER Enrollment " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PRT_DEP_PRNT_SSN_RVW"):
            do_query(query, date + " Parent SSN Review.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PRT_DIAG_AWD_PELL_TERM_") and (year in query[:-10]) :
            do_query(query, date + " Term Pell Awards 20" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("FA_PRT_DISB_PLAN_SPLT_CODE_") and (year in query[:-10]) :
            do_query(query, date + " Disb Plan FY Split Code XX.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_DL_DPAY_SCSP_") and (year in query[:-10]) :
            do_query(query, date + " Disb Plan AY-Split Code SP.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_DSB_PLN_SPLT_CD_FD_") and (year in query[:-10]) :
            do_query(query, date + " Federal Disb Plan FY Split Code XX " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_DSB_PLN_SPLT_CD_SC_") and (year in query[:-10]) :
            do_query(query, date + " Scholarship Disb Plan FY Split Code XX " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_EXPECT_GRAD_TERM_11"):
            do_query(query, date + " Expected Grad Term 1" + str(int(year) - 1) + "8.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PRT_DL_GRAD_TERM_FALL_") and (year in query[:-10]) :
            do_query(query, date + " DL Expected Grad Term Fall " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_GRAD_TRM_FALL_") and (year in query[:-10]) :
            do_query(query, date + " Loan Proration Grad Term Fall " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_GRAD_TRM_SPRING_") and (year in query[:-10]) :
            do_query(query, date + " Loan Proration Grad Term Spring " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_GRNT_UG_5_YR_2BCH"):
            do_query(query, date + " UG 5th YR 2ND Bachelor.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_HEAL_"):
            do_query(query, date + " Heal 20 " + year + ".xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("UUFA_PRT_LEU_C_FSEOG_") and (year in query[:-10]) :
            do_query(query, date + " LEU C Flag Awarded FSEOG " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_LEU_C_PELL_AWARD_") and (year in query[:-10]) :
            do_query(query, date + " LEU C Flag Awarded Pell " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_LEU_E_FSEOG_") and (year in query[:-10]) :
            do_query(query, date + " LEU E Flag Awarded FSEOG " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_LN_CBA_AWD_NO_ELIG"):
            do_query(query, date + " Loan CBA Review Eligible.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PRT_NO_FALL_11" + str(int(year) - 1) + "8"):
            do_query(query, date + " Packaging No Fall 1" + str(int(year) - 1) + "8 (2).xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PRT_NSL_LOAN_RPT_VERI_") and (year in query[:-10]) :
            do_query(query, date + " NSL Loan Need Verification.xls", directory,
                     loans_krv_attachment_list)

        if query.startswith("UUFA_PRT_NURSING_LOAN_RPT_") and (year in query[:-10]) :
            do_query(query, date + " NSL Needs NSL P-N Checklist " + year + ".xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_PRT_ON_LINE_PACKAGING"):
            do_query(query, date + " Manual Packaging Counselors.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_PELL_COMMENT_037_") and (year in query[:-10]) :
            do_query(query, date + " Pell Comment Code 037 20" + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_PELL_ELG_NO_PELL_") and (year in query[:-10]) :
            do_query(query, date + " PELL ELIGIBLE NO PELL " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_PLL_EL_CTZN_NOT_INDCT"):
            do_query(query, date + " Pell Eligible Citizenship Not Indicated.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PRT_PELL_FPEL" + year + "_INITIATED"):
            do_query(query, date + " Pell FPEL" + year + " Initiated.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_PELL_UG_5TH_YR_2ND_BA"):
            do_query(query, date + " Pell UG 5th YR.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_PHARM_NO_HEAL"):
            do_query(query, date + " Pharmacy students with NO HEAL.xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("UUFA_PRT_PKG_AWD_NO_BDGT"):
            do_query(query, date + " Award NO Budget for Term.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_PKG_SCH_AWD_NO_BGT_") and (year in query[:-10]) :
            do_query(query, date + " Scholarship Award NO Budget for Term.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_PRT_PRIOR_TERM_STFFRD_OFR"):
            do_query(query, date + " Cancel Prior Term Stafford Offer " + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_PRT_READY_PKG_" + year + "_ACTIVE"):
            do_query(query, date + " Manual Awd Pkg Active 20" + year + ".xls", directory,
                     ms_attachment_list)

        if query.startswith("UUFA_PRT_SCHOL_GRAD_DATE"):
            do_query(query, date + " Scholarship-Expected Grad Date.xls", directory,
                     acl_attachment_list)

        if query.startswith("UUFA_PRT_SUB_UNSUB_SP_SP"):
            do_query(query, date + " Loan Offrd Disb Plan SP Split Code SP.xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_PRT_SET_HEAL_FLAG_") and (year in query[:-10]) :
            do_query(query, date + " MD - Pharmacy Heal Eligible Flag.xls", directory,
                     loans_kr_attachment_list)

        if query.startswith("UUFA_PRT_STATE_OF_RES_FM_MH_PW"):
            do_query(query, date + " State of Residence FM MH PW.xls", directory,
                     rkm_attachment_list)

        if query.startswith("FA_PRT_STILL_UNPRCD_AFTER_PKG"):
            do_query(query, date + " Students Not Packaged (old).xls", directory,
                     ms_attachment_list)

        if query.startswith("UUFA_PRT_STDNT_NOT_PACKAGED_") and (year in query[:-10]) :
            do_query(query, date + " Students Not Packaged " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PRT_TEACH_CREDENTIAL_") and (year in query[:-10]) :
            do_query(query, date + " Teach Credential 20" + year + ".xls", directory,
                     ak_attachment_list)

    if ac_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", ac_mail, "", ac_attachment_list)
        del ac_attachment_list[:]
    if acl_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", acl_mail, "", acl_attachment_list)
        del acl_attachment_list[:]
    if ak_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if rkm_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", rkm_mail, "", rkm_attachment_list)
        del rkm_attachment_list[:]
    if akv_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", akv_mail, "", akv_attachment_list)
        del akv_attachment_list[:]
    if athletics_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", athletics_mail, "", athletics_attachment_list)
        del athletics_attachment_list[:]
    if loans_r_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", loans_r_mail, "", loans_r_attachment_list)
        del loans_r_attachment_list[:]
    if loans_ak_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", loans_ak_mail, "", loans_ak_attachment_list)
        del loans_ak_attachment_list[:]
    if loans_krv_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", loans_krv_mail, "", loans_krv_attachment_list)
        del loans_krv_attachment_list[:]
    if ms_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", ms_mail, "", ms_attachment_list)
        del ms_attachment_list[:]
    if prof_attachment_list:
        mailer("", aid_year + " Batch Packaging Queries", prof_mail, "", prof_attachment_list)
        del prof_attachment_list[:]


def do_monthlies():
    global aid_year
    for query_name in os.listdir("."):
        if "MR_ATHLETE_RESIDENCY" in query_name:
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    # Create FOLDER variables to be used in Move() operation and establishes the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Monthly', aid_year, month_folder))
        dl_directory = os.path.realpath(os.path.join('C:/Testing Bob/Direct Loans', aid_year, 'Monthly'))
        acct_directory = os.path.realpath(os.path.join('C:\Testing Bob/acct/Chartfields', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monthly', aid_year, month_folder))
        dl_directory = os.path.realpath(os.path.join('O:/Systems/Direct Loans', aid_year, 'Monthly'))
        acct_directory = os.path.realpath(os.path.join('O:/acct/Chartfields', aid_year))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(dl_directory):
        os.makedirs(dl_directory)
    if not os.path.isdir(acct_directory):
        os.makedirs(acct_directory)

    # Change File_Name to be file ac it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if "MR_COMMENT_CODE_298_" in query and year in query[:-10]:
            do_query(query, date + " IASG - Pell Eligible 20 " + year + ".xls", directory,
                     rkm_attachment_list)

        if "MR_3RD_PARTY_CROSSWALK" in query and year in query[:-10]:
            do_query(query, date + " Third Party Crosswalk " + year + ".xls", directory,
                     rac_attachment_list)

        if "MR_3RD_PRT_MNTR_IA_ALL" in query and year in query[:-10]:
            do_query(query, date + " Third Party Monitor " + year + ".xls", directory,
                     rac_attachment_list)

        if "MR_ACAD_LVLS_NOT_SYNC" in query and year in query[:-10]:
            do_query(query, date + " Academic Levels out of SYNC " + year + ".xls", directory,
                     rkm_attachment_list)

        if "MR_ADM_DEFERRAL" in query and year in query[:-10]:
            do_query(query, date + " FA Admission Deferral " + year + ".xls", directory,
                     aj_attachment_list)

        if "MR_ALT_LN_TRNSMIT_HOLD" in query and year in query[:-10]:
            do_query(query, date + " Alt Loan Transmit Hold " + year + ".xls", directory,
                     loans_r_attachment_list)

        if "MR_ATHLETE_RESIDENCY" in query and year in query[:-10]:
            do_query(query, date + " Residency for Athlete Student " + year + ".xls", directory,
                     athletics_attachment_list)

        if "MR_ATHLETE_T53_AWARDS" in query and year in query[:-10]:
            do_query(query, date + " Athlete T53 Awards Accepted " + year + ".xls", directory,
                     athletics_attachment_list)

        if "MR_COD_DL" in query and year in query[:-10]:
            do_query(query, date + " COD DL FATB" + year + " FCRD" + year + " FHMS" + year + ".xls", directory,
                     aka_attachment_list)

        if "MR_COD_PELL_TEACH_IASG" in query and year in query[:-10]:
            do_query(query, date + " COD Grant FCRD" + year + "-FHMS" + year + " Report.xls", directory,
                     aka_attachment_list)

        if "MR_DIR_LN_TRNSMIT_HOLD" in query and year in query[:-10]:
            do_query(query, date + " COD Direct Loan Transmit Hold " + year + ".xls", directory,
                     loans_r_attachment_list)

        if "MR_DISB_ATH_AWD_NOPOST" in query and year in query[:-10]:
            do_query(query, date + " Athlete Waiver Disbursed Not Posted " + year + ".xls", directory,
                     athletics_attachment_list)

        if "MR_DL_DISB_FAILED" in query and year in query[:-10]:
            do_query(query, date + " DL Disbursement Failed " + year + ".xls", dl_directory,
                     krms_attachment_list)

        if "MR_DL_ORIG_AWARD" in query and year in query[:-10]:
            do_query(query, date + " DL ORIG Award " + year + ".xls", directory,
                     athletics_attachment_list)

        if "MR_DSB_CASH_AWD_NOPOST" in query and year in query[:-10]:
            do_query(query, date + " Cash Disbursed Not Posted " + year + ".xls", directory,
                     aca_attachment_list)

        if "MR_DSB_WAVR_AWD_NOPOST" in query and year in query[:-10]:
            do_query(query, date + " Waiver-Scholarship Disbursed Not Posted " + year + ".xls", directory,
                     ac_attachment_list)

        if "MR_DN_INC_CHECKLISTS" in query and year in query[:-10]:
            do_query(query, date + " Dental Students with I Checklists " + year + ".xls", directory,
                     prof_attachment_list)

        if "MR_FWS_WITH_NSI_HOLD" in query and year in query[:-10]:
            do_query(query, date + " FWS with NSI Holds.xls", directory,
                     a_attachment_list)

        if "MR_GRAD_TERM_PRB" in query and year in query[:-10]:
            do_query(query, date + " Grad Term Wrong " + year + ".xls", directory,
                     rkm_attachment_list)

        if "MR_ITEM_CHARTFLD_SETUP" in query and year in query[:-10]:
            do_query(query, date + " Item Chartfield Setup.xls", acct_directory,
                     ac_attachment_list)

        if "MR_ITEM_TYPE_DISB_RULE" in query and year in query[:-10]:
            do_query(query, date + " Item Type Career - Match Disb Rule Career " + year + ".xls", directory,
                     acvjm_attachment_list)

        if "MR_LAW_INC_CHECKLISTS" in query and year in query[:-10]:
            do_query(query, date + " Law Students with I Checklists " + year + ".xls", directory,
                     prof_attachment_list)

        if "MR_LOAN_AWD_PARTL_DISB" in query and year in query[:-10]:
            do_query(query, date + " Loan Awards Partial Disbursed " + year + ".xls", directory,
                     loans_r_attachment_list)

        if "MR_MED_INC_CHECKLISTS" in query and year in query[:-10]:
            do_query(query, date + " Med Students with I Checklists " + year + ".xls", directory,
                     prof_attachment_list)

        if "MR_MED_LAW_LVL_REVIEW" in query or "_MR_DN_LW_MD_LVL_RVW" in query and year in query[:-10]:
            do_query(query, date + " MED-LAW Academic Level Review " + year + ".xls", directory,
                     prof_attachment_list)

        if "MR_PART_TW_OTHER_SCH" in query and year in query[:-10]:
            do_query(query, date + " Partial TW Other Scholarship " + year + ".xls", directory,
                     ac_attachment_list)

        if "MR_PELL_AWD_ADJUSTMENT" in query and year in query[:-10]:
            do_query(query, date + " Pell Award Adjust " + year + ".xls", directory,
                     rkm_attachment_list)

        if "MR_PELL_ONLY" in query and year in query[:-10]:
            do_query(query, date + " Pell Awd  Zero Grants Loans " + year + ".xls", directory,
                     rkm_attachment_list)

        if "MR_PELL_SSN_MISMATCH" in query and year in query[:-10]:
            do_query(query, date + " SSN Mismatch " + year + ".xls", directory,
                     rkm_attachment_list)

        if "MR_PERKINS_CLASS_LIMIT" in query and year in query[:-10]:
            do_query(query, date + " Perkins Class Limits " + year + ".xls", directory,
                     rkm_attachment_list)

        if "MR_PERK_MISC_LN_CNCLD" in query and year in query[:-10]:
            do_query(query, date + " Perkins - Misc Loans Cancelled " + year + ".xls", directory,
                     loans_r_attachment_list)

        if "MR_PERK_MISC_LOAN_DISB" in query and year in query[:-10]:
            do_query(query, date + " Perkins - Misc Loans Disbursed " + year + ".xls", directory,
                     loans_r_attachment_list)

        if "MR_SCH" in query and "LOA" in query and year in query[:-10]:
            do_query(query, date + " Scholarship LOA " + year + ".xls", directory,
                     acj_attachment_list)

        if "SCHOLAR_REINSTATE" in query and year in query[:-10]:
            do_query(query, date + " Scholarship Reinstate " + year + ".xls", directory,
                     acj_attachment_list)

        if "MR_SF_DIS_AWD_PT_ER_FC" in query and year in query[:-10]:
            do_query(query, date + " Federal Award Disb Post Error " + year + ".xls", directory,
                     loans_ak_attachment_list)

        if "MR_SF_DIS_AWD_PT_ER_SV" in query and year in query[:-10]:
            do_query(query, date + " SCHOL-ATH Award Disb Post Error " + year + ".xls", directory,
                     ac_attachment_list)

        if "MR_STATE_FM_MH_PW" in query and year in query[:-10]:
            do_query(query, date + " Palau - Micronesia - Marshall Islands Students " + year + ".xls", directory,
                     rkam_attachment_list)

        if "MR_SUSPEND_RC2" in query and year in query[:-10]:
            do_query(query, date + " ISIR Suspended Reason Code 2 " + year + ".xls", directory,
                     rkmv_attachment_list)

        if "MR_UFORM_GRAD_TERM_PRB" in query and year in query[:-10]:
            do_query(query, date + " Grad Term Wrong " + year + ".xls", directory,
                     rac_attachment_list)

        if "MR_UNDS_OFFER_SCHOLAR" in query and year in query[:-10]:
            do_query(query, date + " Scholarship Awards UNDS Career " + year + ".xls", directory,
                     ac_attachment_list)

        if "MR_UNDS_OFRD_AMT_FDRL" in query and year in query[:-10]:
            do_query(query, date + " Athlete Awards UNDS Career " + year + ".xls", directory,
                     athletics_attachment_list)

        if "MR_UNDS_OFRD_AMT_ATH" in query and year in query[:-10]:
            do_query(query, date + " Federal Awards UNDS Career " + year + ".xls", directory,
                     rkm_attachment_list)

        if "MR_VERIFY_DEP_OVERRIDE" in query and year in query[:-10]:
            do_query(query, date + " Verification Dependency Override " + year + ".xls", directory,
                     rkm_attachment_list)

        if "SEFA_DL_TOTAL_AWARDS" in query and year in query[:-10]:
            do_query(query, date + " SEFA DL Amounts 20" + year + ".xls", directory,
                     mary_s_attachment_list)

        if "SEFA_TOTAL_STUDENTS" in query and year in query[:-10]:
            do_query(query, date + " SEFA DL Total Students 20" + year + ".xls", directory,
                     mary_s_attachment_list)

    if a_attachment_list:
        mailer("", aid_year + " Monthly Queries", a_mail, "", a_attachment_list)
        del a_attachment_list[:]
    if ac_attachment_list:
        mailer("", aid_year + " Monthly Queries", ac_mail, "", ac_attachment_list)
        del ac_attachment_list[:]
    if aca_attachment_list:
        mailer("", aid_year + " Monthly Queries", aca_mail, "", aca_attachment_list)
        del aca_attachment_list[:]
    if acj_attachment_list:
        mailer("", aid_year + " Monthly Queries", acj_mail, "", acj_attachment_list)
        del acj_attachment_list[:]
    if acvj_attachment_list:
        mailer("", aid_year + " Monthly Queries", acvj_mail, "", acvj_attachment_list)
        del acvj_attachment_list[:]
    if acvjm_attachment_list:
        mailer("", aid_year + " Monthly Queries", acvjm_mail, "", acvjm_attachment_list)
        del acvjm_attachment_list[:]
    if aka_attachment_list:
        mailer("", aid_year + " Monthly Queries", aka_mail, "", aka_attachment_list)
        del aka_attachment_list[:]
    if aj_attachment_list:
        mailer("", aid_year + " Monthly Queries", aj_mail, "", aj_attachment_list)
        del aj_attachment_list[:]
    if krms_attachment_list:
        mailer("", aid_year + " Monthly Queries", krms_mail, "", krms_attachment_list)
        del krms_attachment_list[:]
    if rkam_attachment_list:
        mailer("", aid_year + " Monthly Queries", rkam_mail, "", rkam_attachment_list)
        del rkam_attachment_list[:]
    if rkm_attachment_list:
        mailer("", aid_year + " Monthly Queries", rkm_mail, "", rkm_attachment_list)
        del rkm_attachment_list[:]
    if rkmv_attachment_list:
        mailer("", aid_year + " Monthly Queries", rkmv_mail, "", rkmv_attachment_list)
        del rkmv_attachment_list[:]
    if athletics_attachment_list:
        mailer("", aid_year + " Monthly Queries", athletics_mail, "", athletics_attachment_list)
        del athletics_attachment_list[:]
    if loans_r_attachment_list:
        mailer("", aid_year + " Monthly Queries", loans_r_mail, "", loans_r_attachment_list)
        del loans_r_attachment_list[:]
    if loans_ak_attachment_list:
        mailer("", aid_year + " Monthly Queries", loans_ak_mail, "", loans_ak_attachment_list)
        del loans_ak_attachment_list[:]
    if mary_s_attachment_list:
        mailer("", aid_year + " Monthly Queries", mary_s_mail, "", mary_s_attachment_list)
        del mary_s_attachment_list[:]
    if prof_attachment_list:
        mailer("", aid_year + " Monthly Queries", prof_mail, "", prof_attachment_list)
        del prof_attachment_list[:]
    if loans_rc_attachment_list:
        mailer("", aid_year + " Monthly Queries", loans_rc_mail, "", loans_rc_attachment_list)
        del loans_rc_attachment_list[:]
    if rac_attachment_list:
        mailer("", aid_year + " Monthly Queries", rac_mail, "", rac_attachment_list)
        del rac_attachment_list[:]
    

def do_end_of_term_queries():
    global aid_year
    term = 'F'
    input = "Error"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_ACADEMIC_PLAN_RVW_FRAP"):
            prompt = "Enter Term: (e.g. 2016U, 2017F, or 2032S):"
            while True:
                input = str.upper(raw_input(prompt))
                aid_year = input[2:4]
                term = input[4]
                if term == 'S' or term == 'U' or term == 'F':
                    break

    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/SAT/', "20" + year + term))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems\QUERIES/SAT/', "20" + year + term))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        if query.startswith("UUFA_ACADEMIC_PLAN_RVW_FRAP"):
            do_query(query, date + " Academic plan RVW FRAP " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_ATHLETE_ACAD_PROG_REVIEW"):
            do_query(query, date + " Athlete Academic Progress Review.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_ATHLETE_CUR_PART_ALL_AWRD"):
            do_query(query, date + " Athlete Participation Report.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_CBA_ACAD_PROG_BELOW_FT"):
            do_query(query, date + " CBA Acad Prog below FT.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_CBA_UNDISBURSED"):
            do_query(query, date + " CBA UNDISBURSED .xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_EOT_ALT_LOAN_SAT"):
            do_query(query, date + " Alt Loans Awards with SAT Holds.xls", directory,
                     loans_ak_attachment_list)

        if query.startswith("UUFA_LOANS_ORIG_FAILED_PENDING"):
            do_query(query, date + " Loans Originated Failed Pending.xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_LOAN_ACAD_PROG_BLW_HT_UND"):
            do_query(query, date + " Loan Acad Prog below HT Undisbursed.xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_LOAN_ACAD_PROG_BLW_HT_SUB"):
            do_query(query, date + " Loan Acad Prog below HT Subsq Disb.xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_PRO_STDNTS_SAT_WARNING"):
            do_query(query, date + " SAT Warning Professional Students.xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_LW1_LW2_LW3_SAT_WARNING"):
            do_query(query, date + " SAT Warning LW1 LW2 LW3.xls", directory,
                     loans_r_attachment_list)

        if query.startswith("UUFA_MED_SAT"):
            do_query(query, date + " Medical SAT Review.xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_MR_FWS_WITH_NSI_HOLD_") and (year in query[:-10]) :
            do_query(query, date + " FWS with NSI Holds.xls", directory,
                     a_attachment_list)

        if query.startswith("UUFA_PELL_ACAD_PROG_LES_THN_") and (year in query[:-10]) :
            do_query(query, date + " Pell Less Than " + year + ".xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_PELL_OFFERED_NOT_DIS"):
            do_query(query, date + " Pell Offered Not Disbursed.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_THESIS_STUDENTS_NONRES"):
            do_query(query, date + " Thesis Students Non-Resident .xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_SCH_PROB_ACAD_PROG_RVW"):
            do_query(query, date + " U Tradition Review.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SCH_U_TRAD_ACAD_PROG_REN"):
            do_query(query, date + " U Tradition Renewal.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SCH_U_TRAD_ACAD_PROG_RVW"):
            do_query(query, date + " Scholarship Probation Review.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SCHOLAR_ACAD_PROG_REVIEW"):
            do_query(query, date + " Scholarship Academic Progress Review Basic.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SAT_AGGCP_LAW"):
            do_query(query, date + " SAT Aggregate Law Career.xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_SAT_AGGCP_MED"):
            do_query(query, date + " SAT Aggregate Med Career.xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_SAT_AGGCP_DN"):
            do_query(query, date + " SAT Aggregate Dental Career.xls", directory,
                     prof_attachment_list)

        if query.startswith("UUFA_EU_FALL_GRADE"):
            do_query(query, date + " EU Grade Fall " + str(int(year) - 1) + ".xls", directory,
                     akl_attachment_list)

        if query.startswith("UUFA_EU_SPRING_GRADE"):
            do_query(query, date + " EU Grade Spring " + year + ".xls", directory,
                     akl_attachment_list)

        if query.startswith("UUFA_EU_SUMMER_GRADE"):
            do_query(query, date + " EU Grade Summer " + year + ".xls", directory,
                     akl_attachment_list)

        if query.startswith("UUFA_PELL_ELIG_ENROLL_NO_AWARD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_SAT_FSAP"):
            do_query(query, date + " FSAP Students.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_SCHOLAR_ACAD_PROG_RVW_TT"):
            do_query(query, date + " Scholarship Academic Progress Review Top Ten.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SCH_ACAD_PROG_REVIEW_79"):
            do_query(query, date + " All 70000792 Item types Academic Progress Review.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SCHOLAR_ALUMNI_CGPA"):
            do_query(query, date + " Scholarship Alumni CGPA.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SCHOLAR_CASH_NO_AWARD"):
            do_query(query, date + " Scholarship Cash NO Award.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SCHOLAR_LEADER_CGPA"):
            do_query(query, date + " Scholarship CGPA Leadership.xls", directory,
                     ac_attachment_list)

        if query.startswith("FA_SU_SCHOLAR_ACAD_PROG_REVIEW"):
            do_query(query, date + " Scholarship Academic Progress Review Summer.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SU_SCHLR_ACAD_PROG_RVW_TT"):
            do_query(query, date + " Scholarship Academic Progress Review Top 10 Summer.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SU_SCHOLAR_ALUMNI_CGPA"):
            do_query(query, date + " Scholarship CGPA Alumni - Summer.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_SU_SCHOLAR_LEADER_CGPA"):
            do_query(query, date + " Scholarship CGPA Leader - Summer.xls", directory,
                     ac_attachment_list)

    if a_attachment_list:
        mailer("", "End of Term Queries", a_mail, "", a_attachment_list)
        del a_attachment_list[:]
    if ac_attachment_list:
        mailer("", "End of Term Queries", ac_mail, "", ac_attachment_list)
        del ac_attachment_list[:]
    if ak_attachment_list:
        mailer("", "End of Term Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if loans_r_attachment_list:
        mailer("", "End of Term Queries", loans_r_mail, "", loans_r_attachment_list)
        del loans_r_attachment_list[:]
    if prof_attachment_list:
        mailer("", "End of Term Queries", prof_mail, "", prof_attachment_list)
        del prof_attachment_list[:]
    if rkm_attachment_list:
        mailer("", "End of Term Queries", rkm_mail, "", rkm_attachment_list)
        del rkm_attachment_list[:]
    if akl_attachment_list:
        mailer("", "End of Term Queries", akl_mail, "", akl_attachment_list)
        del akl_attachment_list[:]
    if loans_r_attachment_list:
        mailer("", "End of Term Queries", loans_r_mail, "", loans_r_attachment_list)
        del loans_r_attachment_list[:]
    if prof_attachment_list:
        mailer("", "End of Term Queries", prof_mail, "", prof_attachment_list)
        del prof_attachment_list[:]
    if athletics_attachment_list:
        mailer("", "End of Term Queries", athletics_mail, "", athletics_attachment_list)
        del athletics_attachment_list[:]


def do_disb_queries():
    global aid_year
    year = "17"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    disb_date = str(raw_input("Enter Date the Disbursement ran in 'MM-DD-YY' format:"))
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Disbursement',
                                                  aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems\QUERIES\Disbursement', aid_year, month_folder))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        if query.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB_") and (year in query[:-10]) :
            do_query(query, disb_date + " Item Types Authorized Not Disbursed " + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("FA_DQ_ATHLETE_RM_BD_") and (year in query[:-10]) :
            do_query(query, disb_date + " Athlete Room and Board " + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("FA_DQ_ATH_OFF_SCHED_RM_BD_") and (year in query[:-10]) :
            do_query(query, disb_date + " Athlete Off Schedule R&B " + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_CASH_DISB_TOTALS_20") and (year in query[:-10]) :
            do_query(query, disb_date + " Cash Disbursement Totals " + year + ".xls", directory,
                     sys_attachment_list)

        if query.startswith("UUFA_DQ_FALL_" + year) :
            do_query(query, disb_date + " DL Fall Awards 20" + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_DQ_FALL_SPRING_") and (year in query[:-10]) :
            do_query(query, disb_date + " DL Fall Spring Awards 20" + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_DQ_SPRING_") and (year in query[:-10]) :
            do_query(query, disb_date + " DL Spring Awards 20" + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_DQ_UG_PLUS_REFUND_IA_") and (year in query[:-10]) :
            do_query(query, disb_date + " DL UG PLUS Refund Borrower " + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_MISC_DISB_TOTALS_20") and (year in query[:-10]) :
            do_query(query, disb_date + " Misc Disbursement Totals " + year + ".xls", directory,
                     sys_attachment_list)

        if query.startswith("UUFA_DQ_MISC_RESOURCE_DISB_") and (year in query[:-10]) :
            do_query(query, disb_date + " Misc Resources Disbursement " + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_NONCASH_DISB_TOTALS_20") and (year in query[:-10]) :
            do_query(query, disb_date + " Non-cash Disbursement Totals " + year + ".xls", directory,
                     sys_attachment_list)

        if query.startswith("UUFA_DQ_PELL_ACPT_GR8_DISB_") and (year in query[:-10]) :
            do_query(query, disb_date + " Pell Accepted Awards Greater Disb " + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_DQ_SF_ITEM_TYPE_ERROR"):
            do_query(query, disb_date + " FA SF Item Type Error " + year + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_DQ_TEACH_GRANT_" + str(int(year) - 1)):
            do_query(query, disb_date + " Teach Grant Recipients 20" + str(int(year) - 1) + ".xls", directory,
                     disb_attachment_list)

        if query.startswith("UUFA_DQ_TEACH_GRANT_") and (year in query[:-10]) :
            do_query(query, disb_date + " Teach Grant Recipients 20" + year + ".xls", directory,
                     disb_attachment_list)

    if disb_attachment_list:
        mailer("", aid_year + " Disbursement Queries", disb_mail, "", disb_attachment_list)
        del disb_attachment_list[:]
    if sys_attachment_list:
        mailer("", aid_year + " Disbursement Queries", sys_mail, "", sys_attachment_list)
        del sys_attachment_list[:]


def do_2nd_ldr():
    global aid_year
    year = 17
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_PRT_PELL_ELG_NO_PELL_"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + str(year)
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Term', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Term', aid_year))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        if query.startswith("UUFA_CBA_UNDISBURSED"):
            do_query(query, date + " CBA Undisbursed.xls", directory,
                     ak_attachment_list)

        if query.startswith("UUFA_HRS_DECREASE_ATH"):
            do_query(query, date + " Hours Decrease Athlete.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_HRS_DECREASE_FC"):
            do_query(query, date + " Hours Decrease FC.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_HRS_DECREASE_SV"):
            do_query(query, date + " Hours Decrease SV.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_PELL_ELIG_ENROLL_NO_AWARD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PELL_OFFERED_NOT_DISB"):
            do_query(query, date + " Pell Offered Not Disbursed.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PELL_DL_MATH990"):
            do_query(query, date + " Pell-DL Math 990.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PELL_DL_ELI575_ELI685"):
            do_query(query, date + " Pell-DL ELI0075 ELI0085.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_SF_DISB_ATH_AWD_NOPOST"):
            do_query(query, date + " Athlete Awd Disb not Posted.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_SF_DISB_WAIVER_AWD_NOPOST"):
            do_query(query, date + " Waiver Awd Disb not Posted.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_PRT_PELL_ELG_NO_PELL_") and (year in query[:-10]) :
            do_query(query, date + " Pell Eligible No Pell 20" + year + ".xls", directory,
                     rkm_attachment_list)

    if ac_attachment_list:
        mailer("", "Second Session LDR Queries", ac_mail, "", ac_attachment_list)
        del ac_attachment_list[:]
    if ak_attachment_list:
        mailer("", "Second Session LDR Queries", ak_mail, "", ak_attachment_list)
        del ak_attachment_list[:]
    if rkm_attachment_list:
        mailer("", "Second Session LDR Queries", rkm_mail, "", rkm_attachment_list)
        del rkm_attachment_list[:]
    if athletics_attachment_list:
        mailer("", "Second Session LDR Queries", athletics_mail, "", athletics_attachment_list)
        del athletics_attachment_list[:]


def do_day_after_ldr():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_PRT_PELL_ELG_NO_PELL_"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/LDR', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/LDR', aid_year))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_LDR_MIN_ENROLLMENT_ATH"):
            do_query(query, date + " Minimum Enrollment Athlete.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_LDR_MINIMUM_ENROLLMENT_FC"):
            do_query(query, date + " Minimum Enrollment FC (Federal & Campus Based Aid).xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_LDR_MINIMUM_ENROLLMENT_SV"):
            do_query(query, date + " Minimum Enrollment SV (Scholarships & Waivers).xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_LDR_PELL_AWARDS"):
            do_query(query, date + " Pell Awards.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_ATHLETE_AWARD_DISBURSED"):
            do_query(query, date + " Athlete Award Disbursed.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_CBA_UNDISBURSED"):
            do_query(query, date + " CBA Undisbursed.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PELL_DL_MATH990"):
            do_query(query, date + " Pell-DL MATH 990.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PELL_DL_ELI575_ELI685"):
            do_query(query, date + " Pell-DL ELI575 ELI685.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PELL_ELIG_ENROLL_NO_AWARD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PELL_OFFERED_NOT_DISB"):
            do_query(query, date + " Pell Offered Not Disbursed.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PELL_SUMMER_ENROLLMENT"):
            do_query(query, date + " Pell Summer Enrollment Check.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_THESIS_STUDENTS_NONRES"):
            do_query(query, date + " Thesis Students Non-Resident.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_REGISTERED_CENSUS_DATE"):
            do_query(query, date + " LDR FA Load Check.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_SF_DISB_ATH_AWD_NOPOST"):
            do_query(query, date + " Athlete Awd Disb not Posted.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_SF_DISB_WAIVER_AWD_NOPOST"):
            do_query(query, date + " Waiver Awd Disb Not Posted.xls", directory,
                     ac_attachment_list)

        if query.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_") and (year in query[:-10]) :
            do_query(query, date + " Pell Eligible No Pell 20" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_NO_MTRC_STU_ATH_BAL_OWING"):
            do_query(query, date + " Non-Matric Stu Athlete Balance Owing.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_FATERM_SOURCE_N_AWD_ATH"):
            do_query(query, date + " NonTerm Source N Cancel Award Rvw - Athlete.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_FATERM_SOURCE_N_AWD_SV"):
            do_query(query, date + " Term Source N Cancel Award Rvw - Scholarships.xls", directory,
                     athletics_attachment_list)

        if query.startswith("UUFA_PRT_PELL_ELG_NO_PELL_"):
            do_query(query, date + " Pell Eligible No Pell " + year + ".xls", directory,
                     athletics_attachment_list)

    if ac_attachment_list:
        mailer("", "Day After LDR Queries", ac_mail, "", ac_attachment_list)
        del ac_attachment_list[:]
    if rkm_attachment_list:
        mailer("", "Day After LDR Queries", rkm_mail, "", rkm_attachment_list)
        del rkm_attachment_list[:]
    if athletics_attachment_list:
        mailer("", "Day After LDR Queries", athletics_mail, "", athletics_attachment_list)
        del athletics_attachment_list[:]


def dl_pre_outbound():
    global aid_year
    year = "16"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_DLR_NOT_DISBURSED"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    orig_file_doc = date + " DL ORIG 20" + year + ".doc"
    orig_file_doc_2 = date + " DL ORIG 20" + year + " (2).doc"
    orig_file_docx = date + " DL ORIG 20" + year + ".docx"
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Direct Loans', aid_year, 'DL Pre-Outbound'))
        orig_doc = os.path.realpath(os.path.join('C:\Testing Bob\Direct Loans', aid_year, 'Origination', orig_file_doc))
        orig_docx = os.path.realpath(os.path.join('C:\Testing Bob\Direct Loans', aid_year, 'Origination', orig_file_docx))

    else:
        directory = os.path.realpath(os.path.join('O:/Systems/Direct Loans', aid_year, 'DL Pre-Outbound'))
        orig_doc = os.path.realpath(os.path.join('O:/Systems/Direct Loans', aid_year, 'Origination', orig_file_doc))
        orig_doc_2 = os.path.realpath(os.path.join('O:/Systems/Direct Loans', aid_year, 'Origination', orig_file_doc_2))
        orig_docx = os.path.realpath(os.path.join('O:/Systems/Direct Loans', aid_year, 'Origination', orig_file_docx))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_DLR_ORIG_TRNS_PEND_") and (year in query[:-10]) :
            do_query(query, date + " DL Orig-Trans Pending " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_ENTRANCE_COUNSEL_") and (year in query[:-10]) :
            do_query(query, date + " DL Entrance Counseling I  " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LOAN_EFT_DT_LNDR_ERR"):
            do_query(query, date + " Loan EFT Date Lender Error.xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LOAN_ORIG_FA_LOAD_"):
            do_query(query, date + " DL ORIG FA Load " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LOAN_NO_NSLDS_") and (year in query[:-10]) :
            do_query(query, date + " Loan No NSLDS " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LOAN_ORIG_ACAD_LVL_") and (year in query[:-10]) :
            do_query(query, date + " Loans Academic Level " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LOAN_ORIG_EDIT_ERR"):
            do_query(query, date + " Loan Originate Edit Errors.xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LOAN_ORIG_SPLT_CDS_") and (year in query[:-10]) :
            do_query(query, date + " Loan Split Codes " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LOAN_ORIG_VLOAN_RSN"):
            do_query(query, date + " Loan ORIG VLOAN Reasons.xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LOAN_SPC_NEED_OVWD_") and (year in query[:-10]) :
            do_query(query, date + " Loan Overaward Special Need " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_LN_ACPT_STAF_31_32_") and (year in query[:-10]) :
            do_query(query, date + " Stafford Accept Offer " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_NOT_DISB_20") and (year in query[:-10]) :
            do_query(query, date + " DL Disb Errors " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_NOT_DISBURSED_") and (year in query[:-10]) :
            do_query(query, date + " DL Disbursement Errors " + year + "(Old).xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_NOT_DISBURSED_20") and (year in query[:-10]) :
            do_query(query, date + " DL Disbursement Errors " + year + ".xls", directory,
                     dl_attachment_list)

        if query.startswith("UUFA_DLR_UG_PLUS_REFND_IND"):
            do_query(query, date + " DL UG PLUS Refund Indicator.xls", directory,
                     dl_attachment_list)

    if not test:
        while True:
            if os.path.isfile(orig_doc):
                dl_attachment_list.append(orig_doc)
                if os.path.isfile(orig_doc_2):
                    dl_attachment_list.append(orig_doc_2)
                break
            if os.path.isfile(orig_docx):
                dl_attachment_list.append(orig_docx)
                if os.path.isfile(orig_doc_2):
                    dl_attachment_list.append(orig_doc_2)
                break
            else:
                raw_input("\nCould not locate DL ORIG 20" + year + ".doc\nMake sure it is located in O:/Systems/Direct Loans/" + aid_year +
                          "/Origination\n\nPress Enter when ready.")

    mailer("", aid_year + " Pre-Outbound Queries", dl_mail, "", dl_attachment_list)
    del dl_attachment_list[:]


def al_pre_outbound():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_ALR_LOAN_ORG_LND_NT_CK_"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    skip = "n"
    orig_file_doc = date + " ALT Loan ORIG 20" + year + ".doc"
    orig_file_docx = date + " ALT Loan ORIG 20" + year + ".docx"

    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/ALT Loans/', aid_year))
        orig_doc = os.path.realpath(os.path.join('C:\Testing Bob/ALT Loans/', aid_year, orig_file_doc))
        orig_docx = os.path.realpath(os.path.join('C:\Testing Bob/ALT Loans/', aid_year, orig_file_docx))
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
        if query.startswith("UUFA_ALR_110_CHNG_PNDNG_TRANS"):
            do_query(query, date + " Loan Pending Change Transactions.xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_CL_APP_RESPONSE_ERRS"):
            do_query(query, date + " CL Response Load Errors.xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LN_FA907_1_REVISE_"):
            do_query(query, date + " Loan Disbursed Report " + year + ".xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LN_FA907_2_REVISE_"):
            do_query(query, date + " CLoan Not Disbursed Report " + year + ".xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LN_SENT_NO_RESP_"):
            do_query(query, date + " CLoan Sent No Response " + year + ".xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_EFT_DETAIL_ERROR"):
            do_query(query, date + " CLoans EFT Detail Error.xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_EFT_DT_LNDR_ERR"):
            do_query(query, date + " Loan EFT Date Lender Errors.xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_ORIG_ACAD_LVL_") and (year in query[:-10]) :
            do_query(query, date + " Loans Academic Level 20" + year + ".xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_ORIG_EDIT_ERRORS"):
            do_query(query, date + " Loan Originate Edit Errors.xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_ORIG_FA_LOAD_") and (year in query[:-10]) :
            do_query(query, date + " Loan ORIG FA Load 20" + year + ".xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_ORG_LND_NT_CK_") and (year in query[:-10]) :
            do_query(query, date + " Loan ORIG Lender Note 20" + year + ".xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_ORIG_SPLT_CDS_") and (year in query[:-10]) :
            do_query(query, date + " Loan ORIG Split Codes 20" + year + ".xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_ORIG_VLOAN_RSN"):
            do_query(query, date + " Loan ORIG VLOAN Reasons.xls", directory,
                     alt_attachment_list)

        if query.startswith("UUFA_ALR_LOAN_SPC_NEED_OVWD_") and (year in query[:-10]) :
            do_query(query, date + " Loan Overaward Special Need " + year + ".xls", directory,
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
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Pell Repackaging', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Pell Repackaging', aid_year))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        # Pre-Pell Repackaging Queries
        if query.startswith("UUFA_PP_RPKG_AGGREGATE_LIMITS"):
            do_query(query, date + " Pell AGG Limits Awards Reduced.xls", directory,
                     v_attachment_list)

        if query.startswith("UUFA_PP_RPKG_AWD_AY_NO_BDGT"):
            do_query(query, date + " Pell Award AY One STRM Budget.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PP_RPKG_AWD_STRM_INACTIVE"):
            do_query(query, date + " Pell Award STRM Inactive.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PP_RPKG_AWRD_LOCK"):
            do_query(query, date + " Pell Award Lock No FPEL.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PP_RPKG_COA_DOUBLE"):
            do_query(query, date + " Pell COA Double.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PP_RPKG_LTHT_PELL_COA"):
            do_query(query, date + " Pell COA LTHT.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PP_RPKG_RPKG_NO_BUDGET"):
            do_query(query, date + " Pell Repackaging No Budget.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_PP_RPKG_RPKG_SNAPSHOT"):
            do_query(query, date + " Pell Repackage Snapshot.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_PP_RPKG_TOTAL_WDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop.xls", directory,
                     rkm_attachment_list)

    if rkm_attachment_list:
        mailer("", "Pre-Pell Only Repackaging Queries", rkm_mail, sys_mail, rkm_attachment_list)
        del rkm_attachment_list[:]
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
        directory = os.path.realpath(os.path.join('C:/QueryRunnerProj/Testing/Test/Pell Repackaging', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Pell Repackaging', aid_year))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    year = aid_year[7:]

    for query in os.listdir("."):
        if query.startswith("UUFA_MP_RPKG_AID_PROC_STATUS_4"):
            do_query(query, date + " Aid Processing Status 4 Repackage.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_MP_RPKG_AWD_STRM_INACTIVE"):
            do_query(query, date + " Pell Award STRM Inactive.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_MP_RPKG_FCIT"):
            do_query(query, date + " Pell Repackage FCIT" + year + " FDEG" + year + ".xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_MP_RPKG_FDR"):
            do_query(query, date + " Pell Repackage FDR" + year + " Initiated.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_MP_RPKG_FOVP_FARC_I"):
            do_query(query, date + " Pell REPKG FOVP FACR Initiated.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_MP_RPKG_ISIR_CMT_346_347"):
            do_query(query, date + " Pell Repackage ISIR CMT 346 347.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_MP_RPKG_LEAVE"):
            do_query(query, date + " Pell Repackaging Leave Absense.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_MP_RPKG_TOTAL_WDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop.xls", directory,
                     sys_attachment_list)

        if query.startswith("UUFA_MP_RPKG_TOTAL_WTHDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop (old).xls", directory,
                     sys_attachment_list)

        if query.startswith("UUFA_MP_RPKG_VAR_1_2"):
            do_query(query, date + " Pell Repackage Flags.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_MP_RPKG_VER_UNFLAG"):
            do_query(query, date + " Pell Verification Flag Unchecked.xls", directory,
                     null_attachment_list)

    if rkm_attachment_list:
        mailer("", "Pell Repackaging Mid-Packaging Queries", rkm_mail, sys_mail, rkm_attachment_list)
        del rkm_attachment_list[:]
        del null_attachment_list[:]
    if sys_attachment_list:
        mailer("", "Pell Repackaging Mid-Repackaging Check Total Withdrawn Drop", sys_mail,"", sys_attachment_list)
        del sys_attachment_list[:]


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
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Pell Repackaging', aid_year))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Pell Repackaging', aid_year))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    year = aid_year[7:]

    for query in os.listdir("."):
        # Pell Repackaging Queries
        if query.startswith("UUFA_AP_RPKG_5TH_YR_2ND_BACH"):
            do_query(query, date + " UG 5th YR 2ND Bachelor.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_ACTN"):
            do_query(query, date + " Pell Award Activity.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_AWACT_C"):
            do_query(query, date + " Pell Awards Cancelled.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_AW_ACT"):
            do_query(query, date + " Pell Repackage Activity.xls", directory,
                     v_attachment_list)

        if query.startswith("UUFA_AP_RPKG_FPEL_AWARD_LCK"):
            do_query(query, date + " Pell Award Lock FPEL" + year + ".xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_PLAN_ID_BLANK"):
            do_query(query, date + " Pell Repackaging Plan ID Blank.xls", directory,
                     v_attachment_list)

        if query.startswith("UUFA_AP_RPKG_RPKG_SNAPSHOT"):
            do_query(query, date + " Pell Repackage Snapshot.xls", directory,
                     null_attachment_list)

        if query.startswith("UUFA_AP_RPKG_SAT_HOLD_DELETED"):
            do_query(query, date + " Pell SAT Holds.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_SKIP"):
            do_query(query, date + " Pell Repackage Skip.xls", directory,
                     ms_attachment_list)

        if query.startswith("UUFA_AP_RPKG_TERM_FT"):
            do_query(query, date + " Term Pell Awards FT.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_TERM_HT"):
            do_query(query, date + " Term Pell Awards HT.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_TERM_LH"):
            do_query(query, date + " Term Pell Awards LH.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_TERM_NL"):
            do_query(query, date + " Term Pell Awards NL.xls", directory,
                     rkm_attachment_list)

        if query.startswith("UUFA_AP_RPKG_TERM_TQ"):
            do_query(query, date + " Term Pell Awards TQ.xls", directory,
                     rkm_attachment_list)

    if rkm_attachment_list:
        mailer("", "Pell Only Repackaging Queries", rkm_mail, sys_mail, rkm_attachment_list)
        del rkm_attachment_list[:]
    if ms_attachment_list:
        mailer("", "Pell Only Repackaging Queries", ms_mail, sys_mail, ms_attachment_list)
        del ms_attachment_list[:]
    if v_attachment_list:
        mailer("", "Pell Only Repackaging Queries", v_mail, sys_mail, v_attachment_list)
        del v_attachment_list[:]


def do_daily_scholarships():
    global aid_year
    year = date[:2]
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_SCHOLAR_DISB_ZERO"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year

    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob', aid_year + ' Scholar\Queries'))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems', aid_year + ' Scholar\Queries'))

    # the list 'my_path' should be populated  with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_SCHOLAR_DISB_ZERO_") and (year in query[:-10]) :
            do_query(query, date + " Scholarships awarded not disbursed " + year + ".xls", directory,
                     ss_attachment_list)

        if query.startswith("UUFA_SCHOLAR_TWO_CAREERS_") and (year in query[:-10]) :
            do_query(query, date + " Scholarship Award with Two Careers " + year + ".xls", directory,
                     jen_attachment_list)

        if query.startswith("UUFA_SCHOLAR_AUTH_NOT_DISB_") and (year in query[:-10]) :
            do_query(query, date + " Scholar Authorized Not Disbursed " + year + ".xls", directory,
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
        if ("BOOKS_NOPOST" in query_name):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year

    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/', aid_year + ' Scholar\Queries'))
        directory_save = os.path.realpath(os.path.join('C:\Testing Bob/Save', aid_year ))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems', aid_year + ' Scholar\Queries'))
        directory_save = os.path.realpath(os.path.join('O:\Systems\Queries\Save', aid_year ))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(directory_save):
        os.makedirs(directory_save)

    # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if ("DEPT_POST_WRNG_ITEM" in query) and (year in query[:-10]) :
            do_query(query, date + " Depts posting to the wrong item type  " + year + ".xls", directory,
                     ssj_attachment_list)

        if ("MISC_TOTAL" in query) and (year in query[:-10]) :
            do_query(query, date + " IT Dept Misc Awards Total (10001) " + year + ".xls", directory_save,
                     null_attachment_list)

        if ("MISC_NOPOST" in query) and (year in query[:-10]) :
            do_query(query, date + " 7880013 Dept Misc Awards No Post " + year + ".xls", directory_save,
                     ssj_attachment_list)

        if ("BOOKS_TOTAL" in query) and (year in query[:-10]) :
            do_query(query, date + " 788 IT Dept Books Awards Total (10002) " + year + ".xls", directory_save,
                     null_attachment_list)

        if ("BOOKS_NOPOST" in query) and (year in query[:-10]) :
            do_query(query, date + " 7880015 Dept Books No Post " + year + ".xls", directory_save,
                     ssj_attachment_list)

        if ("ROOMBOARD_TOTAL" in query) and (year in query[:-10]) :
            do_query(query, date + " 788 IT Dept Room & Board Awards Total (10003) " + year + ".xls", directory_save,
                     null_attachment_list)

        if ("ROOMBOARD_NOPOST" in query) and (year in query[:-10]) :
            do_query(query, date + " 7880029 Dept Room & Board No Post " + year + ".xls", directory_save,
                     ssj_attachment_list)

        if ("TRAVEL_TOTAL" in query) and (year in query[:-10]) :
            do_query(query, date + " 788 IT Dept Travel Awards Total (10004) " + year + ".xls", directory_save,
                     null_attachment_list)

        if ("TRAVEL_NOPOST" in query) and (year in query[:-10]) :
            do_query(query, date + " 7880033 Dept Travel No Post " + year + ".xls", directory_save,
                     ssj_attachment_list)

        if (("TRAINEESHIP" in query) and ("TOTAL" in query)) and (year in query[:-10]) :
            do_query(query, date + " 788 IT Dept Traineeship Awards Total (10005) " + year + ".xls", directory_save,
                     null_attachment_list)

        if ("TRAINEESHIP_NOPOST" in query) and (year in query[:-10]) :
            do_query(query, date + " 7880034 Dept Traineeship No Post " + year + ".xls", directory_save,
                     ssj_attachment_list)

        if ("SCH_ALL_NEED" in query) and (year in query[:-10]) :
            do_query(query, date + " All Scholarships Need Based " + year + ".xls", directory,
                     ssj_attachment_list)

        if ("SCH_ALL_NRFRESH" in query) and (year in query[:-10]) :
            do_query(query, date + " All Scholarships Non Res Freshman " + year + ".xls", directory,
                     ssj_attachment_list)

        if ("SCH_ALL_NRTRAN" in query) and (year in query[:-10]) :
            do_query(query, date + " All Scholarships Non Res Transfer " + year + ".xls", directory,
                     ssj_attachment_list)

        if ("SCH_ALL_RESFRESH" in query) and (year in query[:-10]) :
            do_query(query, date + " All Scholarships Res Freshman " + year + ".xls", directory,
                     ssj_attachment_list)

        if ("SCH_ALL_RESTRAN" in query) and (year in query[:-10]) :
            do_query(query, date + " All Scholarships Res Transfer " + year + ".xls", directory,
                     ssj_attachment_list)


    if ssj_attachment_list:
        mailer("", aid_year + " Weekly Scholarship Queries", ssj_mail, "", ssj_attachment_list)
        del ssj_attachment_list[:]


for filename in os.listdir("."):
    # Daily Queries
    if filename.startswith("UUFA_IL_CMT_CDE_OVR_AGR"):
        do_dailies()
    # Monday Weekly Queries
    if filename.startswith("UUFA_WR_AID_DISB_NO_ENR_ATH_"):
        do_monday_weeklies()
    # Budget Queries
    if filename.startswith("UUFA_BR_ACAD_LVLS_OUT_SYNC"):
        do_budget_queries()
    # Packaging Queries
    if filename.startswith("UUFA_PRT_ACAD_PROG_REVIEW"):
        do_packaging_queries()
    # Monthly Queries
    if filename.startswith("UUFA_MR_ATHLETE_RESIDENCY_") | filename.startswith("FA_MR_ATHLETE_RESIDENCY_"):
        do_monthlies()
    # Disbursement Queries
    if filename.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB"):
        do_disb_queries()
    #2nd LDR Queries
    if filename.startswith("UUFA_PRT_PELL_ELG_NO_PELL"):
      do_2nd_ldr()
    # End of Term Queries
    if filename.startswith("UUFA_ACADEMIC_PLAN_RVW_FRAP_"):
        do_end_of_term_queries()
    # Day After LDR Queries
    if filename.startswith("UUFA_LDR_MIN_ENROLLMENT_ATH"):
        do_day_after_ldr()
    # Direct Loans Pre-Outbound Queries
    if filename.startswith("UUFA_DLR_NOT_DISBURSED"):
        dl_pre_outbound()
    # Alternative Loan Pre-Outbound Queries
    if filename.startswith("UUFA_ALR_LOAN_ORG_LND_NT_CK_"):
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
    if filename.startswith("UUFA_SCHOLAR_DISB_ZERO"):
        do_daily_scholarships()
    # Weekly Scholarships Queries
    if ("BOOKS_NOPOST" in filename):
        do_weekly_scholarships()
    # Budget Testing Queries
    if filename.startswith("UUFA_BUDGET_20"):
        do_budget_test_queries()

        # TEMPLATE
        # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
        # will be added.
        # for query in os.listdir("."):
        # if query.startswith("____________________"):
        #        do_query(query, date + " ________________" + year + ".xls", directory,
        #                 lkj_attachment_list)

        # if ak_attachment_list:
        #    mailer("", aid_year + " _____________", ak_mail, "", ak_attachment_list)
        #   del ak_attachment_list[:]