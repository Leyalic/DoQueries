__author__ = 'mmason'
#Version 4.30
import os
import datetime
import time
import calendar
import shutil
import re

# date becomes the current date and is then placed in MM-DD-YY format
date = time.strftime("%x").replace("/", "-")
now = datetime.datetime.now()
last_month = now.month - 1 if now.month > 1 else 12
last_months_year = now.year - 1 if now.month == 12 else now.year
month_folder = date[:2] + "-20" + date[-2:]
###############################
test = False 
###############################
class MailGroup(object):
    name = ""
    recipients = ""
    attachments = []

    def __init__(self, recipients):
        self.attachments = []
        self.recipients = recipients
        

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
        print ("Already a file with the name:" + name + "at location.")
    except IOError as e:
        print(e)
        pass


def mailer(text, subject, recipient, cc, attachments):
    import win32com.client as win32

    list(attachments)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.cc = cc
    mail.Subject = "PHI " + subject
    mail.HtmlBody = text
    for each in attachments:
        mail.Attachments.Add(Source=each)
    #TODO: add protection for too many email windows open
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


def make_recipients(*args):
    addresses = ""
    for i in args:
        addresses = addresses + i + ";"
    return addresses


# region Email and Attachment Groups/
mat                 = "mat.mason@utah.edu"
atir                = ""
ashley              = "ashley.stevenson@utah.edu;Samantha.Delacerda@utah.edu"
anthony             = "anthony.p.jones@utah.edu"
chelsea             = "chelsea.springer@utah.edu; lamishia.lin@utah.edu"
chelsea_wells       = "chelsea.wells@utah.edu"
emilie              = "emilie.hereth@utah.edu"
heather             = "hhansen@sa.utah.edu"
hilerie             = "hilerie.harris@sa.utah.edu"
maisa               = ""
jenny               = "jenny.ryan@utah.edu"
russ                = ""
karen               = ""
kayla               = "Kayla.LaMont@utah.edu"
kelly               = "kelly.tomlinson@utah.edu"
devon               = "devon.arellano@utah.edu"
buttars             = "bryan.buttars@utah.edu"
leo                 = buttars
linh                = "linh.ly@utah.edu"
lisa                = "steffany.forrest@income.utah.edu"
courtney            = "courtney.l.young@utah.edu"
mateo               = "mateo.remsburg@utah.edu"
justin              = "c.justin.walker@utah.edu"
plooster            = ""
raenetta            = ""
counselor_manager   = ""
evans	            = "walter.hills@utah.edu"
sam                 = "Samantha.Delacerda@utah.edu"
shana               = "Shana.Yem@utah.edu"
sheryl              = ""
steffany            = "steffany.forrest@income.utah.edu"
stephen             = ""
chelsea             = chelsea
walter              = "walter.hills@utah.edu"
dax                 = mat
dallas              = "dallas.nelson@utah.edu"
cory                = "cory.dawson@utah.edu"
lamishia            = "lamishia.lin@utah.edu"

accounting = emilie + ";" + justin + ";" + evans
athletics  = chelsea + ";" + kayla + ";" + dallas
loans      = cory + ";" + heather + ";" + devon
prof       = shana
systems    = mat + ";" + leo + ";" + dax + ";" + buttars + ";" + maisa
schol      = russ + ";" + sheryl + ";" + plooster + ";" + raenetta + ";" + ashley + ";" + jenny + ";" + walter
schol2     = schol + ";" + mateo + ";" + kelly + ";" + chelsea_wells
schol_mngr = ashley

query_groups = []

ckkls_mail      = MailGroup(make_recipients(chelsea, courtney, karen, counselor_manager, stephen)); query_groups.append(ckkls_mail)
acnt_kkmt_mail  = MailGroup(make_recipients(accounting, karen, counselor_manager, courtney, chelsea)); query_groups.append(acnt_kkmt_mail)
acnt_kkt_mail   = MailGroup(make_recipients(accounting, karen, counselor_manager, chelsea)); query_groups.append(acnt_kkt_mail)
acnt_mail       = MailGroup(make_recipients(accounting)); query_groups.append(acnt_mail)
maisa_mail      = MailGroup(make_recipients(maisa)); query_groups.append(maisa_mail)
aehjkkt_mail    = MailGroup(make_recipients(atir, evans, ashley, jenny, karen, counselor_manager, chelsea)); query_groups.append(aehjkkt_mail)
ahjmv_mail      = MailGroup(make_recipients(maisa, ashley, jenny, mat, dax)); query_groups.append(ahjmv_mail)
ahjv_mail       = MailGroup(make_recipients(maisa, ashley, jenny, dax)); query_groups.append(ahjv_mail)
akj_mail        = MailGroup(make_recipients(karen, counselor_manager, jenny, chelsea));query_groups.append(akj_mail)
alt_mail        = MailGroup(make_recipients(loans, counselor_manager, systems)); query_groups.append(alt_mail)
ath_kkmt_mail   = MailGroup(make_recipients(athletics, karen, counselor_manager, courtney, chelsea)); query_groups.append(ath_kkmt_mail)
ath_mail        = MailGroup(make_recipients(athletics, karen)); query_groups.append(ath_mail)
ben_mail        = MailGroup(make_recipients(emilie, justin)); query_groups.append(ben_mail)
bhjjns_mail     = MailGroup(make_recipients(ashley, russ, jenny, justin, sheryl, evans)); query_groups.append(bhjjns_mail)
bekkt_mail      = MailGroup(make_recipients(emilie, karen, counselor_manager, justin, chelsea, lamishia, cory)); query_groups.append(bekkt_mail)
bhjkn_mail      = MailGroup(make_recipients(ashley, jenny, karen, justin)); query_groups.append(bhjkn_mail)
bjnrs_mail      = MailGroup(make_recipients(russ, justin, raenetta, sheryl)); query_groups.append(bjnrs_mail)
bjns_mail       = MailGroup(make_recipients(russ, justin, sheryl)); query_groups.append(bjns_mail)
bk_mail         = MailGroup(make_recipients(anthony, karen));query_groups.append(bk_mail)
bkk_mail        = MailGroup(make_recipients(anthony, karen, counselor_manager, chelsea));query_groups.append(bkk_mail)
disb_mail       = MailGroup(make_recipients(loans, ashley, jenny, karen, counselor_manager, courtney, justin, chelsea, systems)); query_groups.append(disb_mail)
disb_tot_mail   = MailGroup(make_recipients(systems, counselor_manager, chelsea, karen, courtney, anthony)); query_groups.append(disb_tot_mail)
dl_mail         = MailGroup(make_recipients(loans, karen, counselor_manager, chelsea, systems)); query_groups.append(dl_mail)
ashley_mail     = MailGroup(make_recipients(ashley)); query_groups.append(ashley_mail)
hj_mail         = MailGroup(make_recipients(ashley, jenny)); query_groups.append(hj_mail)
hjj_mail        = MailGroup(make_recipients(ashley, jenny, russ)); query_groups.append(hjj_mail)
hjjk_mail       = MailGroup(make_recipients(ashley, jenny, russ, counselor_manager,ashley)); query_groups.append(hjjk_mail)
hjjr_mail       = MailGroup(make_recipients(ashley, russ, jenny, raenetta)); query_groups.append(hjjr_mail)
hjkkmt_mail     = MailGroup(make_recipients(ashley, jenny, karen, counselor_manager, courtney, chelsea)); query_groups.append(hjkkmt_mail)
hjkt_mail       = MailGroup(make_recipients(ashley, jenny, counselor_manager, chelsea, ashley)); query_groups.append(hjkt_mail)
hjr_mail        = MailGroup(make_recipients(ashley, jenny, raenetta)); query_groups.append(hjr_mail)
hjs_mail        = MailGroup(make_recipients(ashley, jenny, sheryl)); query_groups.append(hjs_mail)
hk_mail         = MailGroup(make_recipients(ashley, karen)); query_groups.append(hk_mail)
jj_mail         = MailGroup(make_recipients(russ, jenny)); query_groups.append(jj_mail)
jja_mail        = MailGroup(make_recipients(russ, ashley)); query_groups.append(jja_mail)
kklt_mail       = MailGroup(make_recipients(karen, counselor_manager, chelsea)); query_groups.append(kklt_mail)
kkmt_mail       = MailGroup(make_recipients(karen, counselor_manager, courtney, chelsea)); query_groups.append(kkmt_mail)
kkmtv_mail      = MailGroup(make_recipients(karen, counselor_manager, courtney, chelsea, dax)); query_groups.append(kkmtv_mail)
kkt_mail        = MailGroup(make_recipients(karen, counselor_manager, chelsea)); query_groups.append(kkt_mail)
kktv_mail       = MailGroup(make_recipients(karen, counselor_manager, chelsea, dax)); query_groups.append(kktv_mail)
kmmt_mail       = MailGroup(make_recipients(counselor_manager, courtney, shana, chelsea, ashley)); query_groups.append(kmmt_mail)
kmt_mail        = MailGroup(make_recipients(counselor_manager, courtney, chelsea)); query_groups.append(kmt_mail)
krms_mail       = MailGroup(make_recipients(devon, counselor_manager, chelsea, justin)); query_groups.append(krms_mail)
kt_mail         = MailGroup(make_recipients(counselor_manager, chelsea)); query_groups.append(kt_mail)
leo_mail        = MailGroup(make_recipients(leo)); query_groups.append(leo_mail)
loans_kkt_mail  = MailGroup(make_recipients(loans, karen, counselor_manager, chelsea)); query_groups.append(loans_kkt_mail)
loans_kktv_mail = MailGroup(make_recipients(loans, karen, counselor_manager, chelsea, dax)); query_groups.append(loans_kktv_mail)
loans_kt_mail   = MailGroup(make_recipients(loans, counselor_manager, chelsea)); query_groups.append(loans_kt_mail)
loans_c_mail    = MailGroup(make_recipients(loans, chelsea));query_groups.append(loans_c_mail)
loans_mail      = MailGroup(make_recipients(loans)); query_groups.append(loans_mail)
buttars_mail     = MailGroup(make_recipients(buttars)); query_groups.append(buttars_mail)
mat_mail        = MailGroup(make_recipients(mat)); query_groups.append(mat_mail)
ml_mail         = MailGroup(make_recipients(mat, leo)); query_groups.append(ml_mail)
monthly_dl_mail = MailGroup(make_recipients(devon, heather, cory, counselor_manager, chelsea, karen, anthony, justin, evans, walter)); query_groups.append(monthly_dl_mail)
natalie_mail    = MailGroup(make_recipients(justin)); query_groups.append(natalie_mail)
null_mail       = MailGroup(make_recipients("")) #do not add this to list of mail_groups, don't want to send email.
prof_jkt_mail   = MailGroup(make_recipients(prof, jenny, counselor_manager, chelsea)); query_groups.append(prof_jkt_mail)
prof_k_mail     = MailGroup(make_recipients(prof, karen)); query_groups.append(prof_k_mail)
prof_kkt_mail   = MailGroup(make_recipients(prof, karen, counselor_manager, chelsea)); query_groups.append(prof_kkt_mail)
prof_kktm_mail  = MailGroup(make_recipients(prof, karen, counselor_manager, chelsea, courtney)); query_groups.append(prof_kktm_mail)
prof_kmt_mail   = MailGroup(make_recipients(prof, counselor_manager, courtney, chelsea)); query_groups.append(prof_kmt_mail)
prof_mail       = MailGroup(make_recipients(prof, chelsea, counselor_manager,)); query_groups.append(prof_mail)
prof_rk_mail    = MailGroup(make_recipients(prof, chelsea, counselor_manager, karen)); query_groups.append(prof_rk_mail)
schol_mail      = MailGroup(make_recipients(schol)); query_groups.append(schol_mail)
schol2_mail     = MailGroup(make_recipients(schol2)); query_groups.append(schol2_mail)
sheryl_mail     = MailGroup(make_recipients(sheryl)); query_groups.append(sheryl_mail)
sl_mail         = MailGroup(make_recipients(steffany, lisa)); query_groups.append(sl_mail)
ss_mail         = MailGroup(make_recipients(ashley, jenny, russ, justin, sheryl, systems)); query_groups.append(ss_mail)
sys_mail        = MailGroup(make_recipients(systems)); query_groups.append(sys_mail)
v_mail          = MailGroup(make_recipients(dax)); query_groups.append(v_mail)
vm_mail         = MailGroup(make_recipients(dax, mat)); query_groups.append(vm_mail)
russ_mail       = MailGroup(make_recipients(russ)); query_groups.append(russ_mail)
# endregion

#Daily Queries
def do_dailies():
    global aid_year
    year = date[:2]
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_IL_ISIR_SSN_MISMATCH"):
            year = str(int(re.search(r'\d+', query_name).group()))
            aid_year = "20" + str(int(year) - 1) + "-20" + year
            break
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Daily', aid_year, month_folder))
        royall_directory = os.path.realpath('C:\Testing Bob/Royall')
        pell_directory = os.path.realpath(os.path.join('C:\Testing Bob/QUERIES\Pell Repackaging', aid_year))
        disb_directory = os.path.realpath('C:\Testing Bob\QUERIES\Disbursement\Pre-Disbursement Queries')
        refund_directory = os.path.realpath(os.path.join('C:\Testing Bob/QUERIES\Refund Credit Holds', month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Daily', aid_year, month_folder))
        royall_directory = os.path.realpath('O:/Systems/Royall')
        pell_directory = os.path.realpath(os.path.join('O:\Systems\QUERIES\Pell Repackaging', aid_year))
        disb_directory = os.path.realpath("O:\Systems\QUERIES\Disbursement\Pre-Disbursement Queries")
        refund_directory = os.path.realpath(os.path.join('O:\Systems\QUERIES\Refund Credit Holds', month_folder))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(royall_directory):
        os.makedirs(royall_directory)
    if not os.path.isdir(pell_directory):
        os.makedirs(pell_directory)
    if not os.path.isdir(disb_directory ):
        os.makedirs(disb_directory)
    if not os.path.isdir(refund_directory):
        os.makedirs(refund_directory)

    # Change File_Name to be file as it is received and _new_file_name to what
    # the new file should be.  Prefix date will be added.
    for query in os.listdir("."):
        if "IL_ATHLETE_OVERAWARD_" in query and year in query[:-8]:
            do_query(query, date + " Athlete Aid Overaward " + year + ".xls", directory,
                     ath_mail.attachments) 

        if "IL_CMT_CDE_OVR_AGR_LMT_" in query and year in query[:-8]:
            do_query(query, date + " Comment Code Over Aggregate " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "IL_COMMENT_CODE_298_" in query and (year in query[:-8]) :
            do_query(query, date + " IASG - Pell Eligible 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if "IL_FDEG_FFBD_FBLK_FFBC_" in query and (year in query[:-8]) :
            do_query(query, date + " Complete FDEG FFBD FBLK FFBC " + year + ".xls", directory,
                     prof_kmt_mail.attachments)

        if "IL_COMPLETE_FDEG_" in query and (year in query[:-8]) :
            do_query(query, date + " FDEG Update 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if "IL_CORR_NOT_MARK_SENT_" in query and (year in query[:-8]) :
            do_query(query, date + " Corrections not Marked to Sent 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if "UUFA_IL_CORR_SENT_RJCT_CD1_" in query and (year in query[:-8]) :
            do_query(query, date + " Correction Sent Reject Code 1 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "IL_ATHLETE_RESIDENCY_" in query and (year in query[:-8]) :
            do_query(query, date + " ATH Res Change " + year + ".xls", directory,
                     ath_mail.attachments)
					 
        if query.startswith("FA_IL_ENRL_GR_DATE_ERRORS_") and (year in query[:-8]) :
            do_query(query, date + " Place FDIP" + year + " Checklist 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_FBKP") and (year in query[:-8]) :
            do_query(query, date + " FBKP" + year + " Checklist Initiated.xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_FOUT") and (year in query[:-8]) :
            do_query(query, date + " Outside Resources 20" + year + ".xls", directory,
                     hjr_mail.attachments)

        if query.startswith("FA_IL_FP1B") and (year in query[:-8]) :
            do_query(query, date + " FP1B" + year + " Checklist " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("FA_IL_FP2B") and (year in query[:-8]) :
            do_query(query, date + " FP2B" + year + " Checklist " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("FA_IL_FP1N") and (year in query[:-8]) :
            do_query(query, date + " FP1N" + year + " Checklist " + year + ".xls", directory,
                     hjr_mail.attachments)

        if query.startswith("FA_IL_FP2N") and (year in query[:-8]) :
            do_query(query, date + " FP2N" + year + " Checklist " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_FPJ") and (year in query[:-8]) :
            do_query(query, date + " FPJ" + year + " Checklist 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_ISIR_02_IND_UP_DWN_") and (year in query[:-8]) :
            do_query(query, date + " ISIR Service IND UP Down 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_FEHU") and (year in query[:-8]) :
            do_query(query, date + " Initiated FEHU" + year + " Checklist.xls", directory,
                     kt_mail.attachments)

        if query.startswith("UUFA_IL_IRS_DRT_02") and (year in query[:-8]) :
            do_query(query, date + " IRS Data Retrieval Equal to 02 " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("FA_IL_ISIR_CMT_CODE_359_360_") and (year in query[:-8]) :
            do_query(query, date + " ISIR Comment Code 359 or 360 " + year + ".xls", directory,
                     kt_mail.attachments)

        if query.startswith("UUFA_IL_ISIR_GRD_I_UG_FATRM_") and (year in query[:-8]) :
            do_query(query, date + " ISIR Graduate Independent UG FATERM 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_ISIR_PRMARY_EFC_DIF_") and (year in query[:-8]) :
            do_query(query, date + " Primary EFC Difference 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_ISIR_LOADED_NOT_PKG_") and (year in query[:-8]) :
            do_query(query, date + " ISIR Loaded Not Packaged 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_LN_EXP_GRAD_DATE") and (year in query[:-8]) :
            do_query(query, date + " Loan Awd No FLPR Grad Date " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if (("STDT_WITH_FLPR_INIT") in query) and (year in query[:-8]) :
            do_query(query, date + " Students with FLPR Initiated.xls", directory,
                     loans_kt_mail.attachments)
            
        if query.startswith("UUFA_IL_NURSING_LOANS_TILA_") and (year in query[:-8]) :
            do_query(query, date + " Nursing Loans 20" + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if query.startswith("FA_IL_OTHER_ATB_20") and (year in query[:-8]) :
            do_query(query, date + " ISIR Other ATB 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("FA_IL_OTHER_ATTND_") and (year in query[:-8]) :
            do_query(query, date + " Attend Other Institution 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_PELL_LEU_C_") and (year in query[:-8]) :
            do_query(query, date + " Pell LEU Limit Flag C " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_PELL_LEU_E_") and (year in query[:-8]) :
            do_query(query, date + " Pell LEU Limit Flag E " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_PELL_MAX_ELIG_") and (year in query[:-8]) :
            do_query(query, date + " Pell Max Eligibility " + year + ".xls", directory,
                     kt_mail.attachments)

        if query.startswith("UUFA_IL_SF_RFND_AWD_NO_POST_") and (year in query[:-8]) :
            do_query(query, date + " Refund Post Third Party 20" + year + ".xls", directory,
                     hjr_mail.attachments)

        if query.startswith("UUFA_IL_SUB_ISIR_NO_PACKAGE_") and (year in query[:-8]) :
            do_query(query, date + " Subsequent ISIR Not Package Not Verified 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if "IL_SUB_ISIR_SYSG" in query and (year in query[:-8]) :
            do_query(query, date + " Subsequent ISIR System Generated 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_IL_VET_ACTV_DUTY_STAT_") and (year in query[:-8]) :
            do_query(query, date + " Veteran Active Duty Status 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_VER_I_SUB_SUSP_ISIR_") and (year in query[:-8]) :
            do_query(query, date + " FAVR Initiated Susp ISIR Psbl DRT " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_V5_VER_AFTR_OTH_VER_") and (year in query[:-8]) :
            do_query(query, date + " Selected for V5 after other Ver " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_UPDATED_SEC_") and (year in query[:-8]) :
            do_query(query, date + " New ISIR Updated ATB " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_FATN_INITIATED_") and (year in query[:-8]) :
            do_query(query, date + " Review FATN Checklist 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_FED_AID_OVERAWARD_") and (year in query[:-8]) :
            do_query(query, date + " Federal Aid Overaward " + year + ".xls", directory,
                     kmmt_mail.attachments)
        
        if query.startswith("UUFA_IL_ALL_ITEMS_OVERAWARD") and (year in query[:-8]) :
            do_query(query, date + " All Items Overaward " + year + ".xls", directory,
                     kmmt_mail.attachments)

        if query.startswith("UUFA_IL_FHST_I_HST_COMPLETE_") and (year in query[:-8]) :
            do_query(query, date + " HS Transcript 'C' FHST" + year + " I.xls", directory,
                     kmt_mail.attachments)

        # if query.startswith("ussf0034"):
        #     do_query(query, date + " " + query, refund_directory,
        #              aehjkkt_mail.attachments)

        if query.startswith("UUFA_REFUND_HOLDS"):
            do_query(query, date + " " + query, refund_directory,
                     null_mail.attachments)

        if query.startswith("UUFA_IL_PKG_SCH_EXP_GRAD_FA_") and (year in query[:-8]) :
            do_query(query, date + " Scholarship Aid Grad Date Fall 20" + year + ".xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_IL_PKG_FED_EXP_GRAD_FA_") and (year in query[:-8]) :
            do_query(query, date + " Accepted Federal Aid Grad Date Fall " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_FPEL") and (year in query[:-8]) :
            do_query(query, date + " FPEL" + year + " No Database Match.xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_PACKAGING_C_NO_FED_AID_") and (year in query[:-8]) :
            do_query(query, date + " Packaging C 71%-72% No Fed Aid " + year + ".xls", directory,
                     sys_mail.attachments)

        if query.startswith("UUFA_PERKINS_NOT_DISBURSED"):
            do_query(query, date + " Perkins Not Disbursed " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if query.startswith("UUFA_ITM_TYPE_NOT_IN_NEXT_AY"):
            do_query(query, date + " Item Types in " + str(int(year) - 1) + " not in " + year + ".xls", directory,
                     mat_mail.attachments)

        if query.startswith("UUFA_IL_FREV_CHECKLIST_INT_") and (year in query[:-8]) :
            do_query(query, date + " FREV Checklist Initiated " + year + ".xls", directory,
                     kmt_mail.attachments)

        if query.startswith("UUFA_IL_FREV_DOES_NOT_EXIST_") and (year in query[:-8]) :
            do_query(query, date + " FREV Does Not Exist 20" + year + ".xls", directory,
                     kmt_mail.attachments)

        if "IL_STU_WITH_06_CODE" in query and year in query[:-8]:
            do_query(query, date + " Students with Code 06 " + year + ".xls", directory,
                     kmt_mail.attachments)

        if "IL_GRNT_UG_5_YR_2BCH" in query and year in query[:-8]:
            do_query(query, date + " UG 5th YR 2ND Bachelor.xls", directory,
                     kt_mail.attachments)

        if "IL_NSLDS_NO_MCH_DB_FLG_" in query and (year in query[:-8]) :
            do_query(query, date + " NSLDS No Match DB Flag " + year + ".xls", directory,
                     kmt_mail.attachments)

        if "IL_FAVR_CHKLST_INT_" in query and (year in query[:-8]) :
            do_query(query, date + " FAVR Checklist Initiated " + year + ".xls", directory,
                     kmt_mail.attachments)

        if "QJ23029" in query :
            do_query(query, query, royall_directory,
                     null_mail.attachments)

        if "IL_RES_NON_RES_BDGT_" in query and (year in query[:-8]) :
            do_query(query, date + " Resident - Non-Resident Budget " + year + ".xls", directory,
                     kkt_mail.attachments)

        if "PELL_RPKG_VAR_FLAG_2" in query :
            do_query(query, date + " PELL RPKG Var Flag 2.xls", pell_directory,
                     null_mail.attachments)

        if "IL_3RD_PARTY_EXCEPT" in query and (year in query[:-8]) :
            do_query(query, date + " Third Party Change Exceptions " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "IL_3RD_PARTY_MAIN" in query and (year in query[:-8]) and "SU" not in query :
            do_query(query, date + " Third Party Main Report " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "IL_3RD_PARTY_NO_BDGT" in query and (year in query[:-8]) :
            do_query(query, date + " Third Party No Budget " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "IL_3RD_PARTY_MAIN_SU" in query and (year in query[:-8]) :
            do_query(query, date + " Third Party Main Rpt Summer " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "PDQ_SAP_HOLD_DEL" in query:
            do_query(query, date + " Offers with Previous SAP " + year + ".xls", disb_directory,
                     null_mail.attachments)

        if "SHOPPING_SHEET_UG_" in query:
            do_query(query, date + " Shopping Sheet Group List " + year + ".xls", disb_directory,
                     null_mail.attachments)

        if "_IL_REPEAT_COURSES_FALL_" in query and (year in query[:-8]):
            do_query(query, date + " Repeat Courses Fall " + year + ".xls", disb_directory,
                     bkk_mail.attachments)

        if "_IL_REPEAT_COURSES_SPR_" in query and (year in query[:-8]):
            do_query(query, date + " Repeat Courses Spring " + year + ".xls", disb_directory,
                     bkk_mail.attachments)

        if "_IL_REPEAT_COURSES_SUM_" in query and (year in query[:-8]):
            do_query(query, date + " Repeat Courses Summer " + year + ".xls", disb_directory,
                     bkk_mail.attachments)

        if "_INCOMPLETE_CHECKLIST_" in query:
            do_query(query, date + " famil2 " + year + ".xls", disb_directory,
                     null_mail.attachments)

        if "IL_SUMMER_AWARDING_" in query and (year in query[:-8]):
            do_query(query, date + " Summer Awarding " + year + ".xls", disb_directory,
                     ckkls_mail.attachments)
            
        if "IL_ISIR_SSN_MISMATCH" in query and (year in query[:-8]):
            do_query(query, date + " ISIR SSN Mismatch " + year + ".xls", disb_directory,
                     kmt_mail.attachments)


    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Daily Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Monday Weekly Queries
def do_monday_weeklies():
    global aid_year
    global year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_WR_"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    # Create FOLDER variables to be used in Move() operation and establishes
    # the directory to save the files
    # using the date.
    if test:
        directory               = os.path.realpath(os.path.join('C:/Testing Bob/Monday Weekly', aid_year, month_folder))
        packaging_directory     = os.path.realpath(os.path.join('C:/Testing Bob/Packaging', aid_year, month_folder))
        disb_failure_directory  = os.path.realpath(os.path.join('C:/Testing Bob/Disb Failure ' + aid_year))
        save_directory          = os.path.realpath(os.path.join('C:/Testing Bob/SAVE', aid_year))
    else:
        directory               = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monday Weekly', aid_year, month_folder))
        packaging_directory     = os.path.realpath(os.path.join('O:/Systems/QUERIES/Packaging', aid_year, month_folder))
        disb_failure_directory  = os.path.realpath(os.path.join('O:/Disbursement Failure/Disb Failure ' + aid_year))
        save_directory          = os.path.realpath(os.path.join('O:/Systems/QUERIES/SAVE', aid_year))

    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(packaging_directory):
        os.makedirs(packaging_directory)
    if not os.path.isdir(disb_failure_directory):
        os.makedirs(disb_failure_directory)
    if not os.path.isdir(save_directory):
        os.makedirs(save_directory)

    # Change File_Name to be query ac it is received and _new_file_name to what
    # the new query should be.Prefix date
    # will be added.
    for query in os.listdir("."):
        if "_WR_ACAD_LVLS_OUT_SYNC_" in query and (year in query[:-8]) :
            do_query(query, date + " UG Academic Levels Out of Sync " + year + ".xls", directory,
                     kkt_mail.attachments)

        if "_WR_ADV_FSOI_INITIATED_" in query and (year in query[:-8]) :
            do_query(query, date + " FSOI Initiated  " + year + ".xls", directory,
                     kmt_mail.attachments)

        if "_WR_AGG_CK_MLT_YR_AWDED_" in query and (year in query[:-8]) :
            do_query(query, date + " Student Pkgd for " + str(int(year) - 1) + " after " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_AID_DISB_NO_ENR_ATH_" in query and (year in query[:-8]) :
            do_query(query, date + " Athlete Disb Not Enrolled " + year + ".xls", directory,
                     ath_mail.attachments)

        if "_WR_AID_DISB_NO_ENR_FED_" in query and (year in query[:-8]) :
            do_query(query, date + " Federal Disb Not Enrolled " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_AID_DISB_NO_ENR_SCH_" in query and (year in query[:-8]) :
            do_query(query, date + " T 53 Sch Disb Not Enrolled " + year + ".xls", directory,
                     hj_mail.attachments)

        if "_WR_ALL_V4_V5_VER_" in query and (year in query[:-8]) :
            do_query(query, date + " All Verification V4-V5 " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_AMERICORP_AWD_POST_" in query and (year in query[:-8]) :
            do_query(query, date + " Americorp Awards " + year + ".xls", directory,
                     hk_mail.attachments)

        if "_WR_ATHLETE_NOT_DISB_" in query and (year in query[:-8]) :
            do_query(query, date + " Athlete Not Disbursed " + year + ".xls", directory,
                     ath_mail.attachments)

        if "_WR_ATH_HRS_AFTR_CENSUS_" in query and (year in query[:-8]) :
            do_query(query, date + " Ath Hours After Census " + year + ".xls", directory,
                     ath_mail.attachments)

        if "_WR_ATH_SF_TERM_BALANCE" in query and (year in query[:-8]) :
            do_query(query, date + " Athlete Tuition Fee Balance " + year + ".xls", directory,
                     ath_mail.attachments)

        if "_WR_AUDIT_CLSS_AID_DISB_" in query and (year in query[:-8]) :
            do_query(query, date + " Audit Class Aid Disbursed " + year + ".xls", directory,
                     kkt_mail.attachments)

        if "_WR_ISA_PER_TERM" in query and (year in query[:-8]) :
            do_query(query, date + " ISA Per Term FISA" + year + ".xls", directory,
                     kkt_mail.attachments)  

        if "_WR_AWD_UG_NOW_GRAD_ATH_" in query and (year in query[:-8]) :
            do_query(query, date + " Ath Awards past Grad Term " + year + ".xls", directory,
                     ath_mail.attachments)

        if "_WR_AWD_UG_NOW_GRAD_FC_" in query and (year in query[:-8]) :
            do_query(query, date + " Federal Awards past Grad Term " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_AWD_UG_NOW_GRAD_SV_" in query and (year in query[:-8]) :
            do_query(query, date + " Scholar Awards past Grad Term " + year + ".xls", directory,
                     hj_mail.attachments)

        if "_WR_CHKLST_STATUS_ERROR_" in query and (year in query[:-8]) :
            do_query(query, date + " Checklist Status Error " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "WR_CMT_CDE_O_AGR_LMT_2" in query and year in query[:-8]:
            do_query(query, date + " Comment Code Over Aggregate No FATERM Req " + year + ".xls", directory,
                     loans_kkt_mail.attachments)

        if "_WR_DISB_ATH_FAILURE_" in query and (year in query[:-8]) :
            do_query(query, date + " Authorization Failure 20" + year + ".xls", disb_failure_directory,
                     null_mail.attachments)

        if "_WR_DL_DISBURSED_LTHT_" in query and (year in query[:-8]) :
            do_query(query, date + " DL Disbursed LTHT " + year + ".xls", directory,
                     kkt_mail.attachments)

        if "_WR_DL_EC_SUSPENDED_" in query and (year in query[:-8]) :
            do_query(query, date + " DL Entrance Counseling Suspense " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_DL_ORIG_TRNS_PEND_" in query and (year in query[:-8]) :
            do_query(query, date + " DL Orig Trans Pending " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_EFT_CONSENT_VERIF" in query and (year in query[:-8]) :
            do_query(query, date + " EFT Consent Verification 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_FAFSA_CKLST_INCMP_" in query and (year in query[:-8]) :
            do_query(query, date + " PLUS FAFSA Incomplete " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if  "_WR_FALL_TOTAL_WDRN_DRP_" in query and (year in query[:-8]) :
            do_query(query, date + " Fall Disb Total Withdrawn Drop " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if  "_WR_SPR_TOTAL_WDRN_DRP_" in query and (year in query[:-8]) :
            do_query(query, date + " Spring Disb Total Withdrawn Drop " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if  "WR_SNGDO_CAMPUS" in query and year in query[:-8] :
            do_query(query, date + " Asian-SNGDO Campus " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if  "_WR_SUM_TOTAL_WDRN_DRP_" in query and (year in query[:-8]) :
            do_query(query, date + " Summer Disb Total Withdrawn Drop " + year + " .xls", directory,
                     kkmt_mail.attachments)

        if "_WR_FARC_CHECKLIST_" in query and (year in query[:-8]) :
            do_query(query, date + " FARC 30 Day Review " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_FARC_CMNT_CODES_" in query and (year in query[:-8]) :
            do_query(query, date + " Initiated FARC w ISIR Cmnt Codes " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_FED_AID_OVERAWARD_" in query and (year in query[:-8]) :
            do_query(query, date + " Federal Aid Overaward " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_FGED_ISIR_DEGREE_" in query and (year in query[:-8]) :
            do_query(query, date + " FGED ISIR Degree 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if ("_WR_FPEL" + year + "_INITIATED_AWDED") in query:
            do_query(query, date + " FPEL" + year + " Initiated Pell.xls", directory,
                     kkt_mail.attachments)

        if "WR_FREV_GR_WS" in query and year in query:
            do_query(query, date + " Grad FREV with Work Study" + year + ".xls", directory,
                     kmt_mail.attachments)

        if "_WR_GENDER_" in query and (year in query[:-8]) :
            do_query(query, date + " Gender Discrepancies 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_HEDU_PARAMEDIC_" in query and (year in query[:-8]) :
            do_query(query, date + " HEDU Paramedic Class 20" + year + "F.xls", directory,
                     kkt_mail.attachments)

        if "R_HOME_SCHOOLED_" in query and (year in query[:-8]) :
            do_query(query, date + " Home Schooled Check " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_HRS_DECREASE_ATH_" in query and (year in query[:-8]) :
            do_query(query, date + " Hours Decrease Athlete " + year + ".xls", directory,
                     ath_mail.attachments)

        if "_WR_HRS_DECREASE_FC_" in query and (year in query[:-8]) :
            do_query(query, date + " Hours Decrease FC " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_HRS_DECREASE_SV_" in query and (year in query[:-8]) :
            do_query(query, date + " Hours Decrease SV " + year + ".xls", directory,
                     ashley_mail.attachments)

        if "_WR_FHST_I_HST_COMPLETE_" in query and (year in query[:-8]) :
            do_query(query, date + " HS Transcript 'C' FHST" + year + " 'I'.xls", directory,
                     kkmt_mail.attachments)

        if "_WR_ISIR_AS_EFC_" in query and (year in query[:-8]) :
            do_query(query, date + " ISIR Assumption EFC 20" + year + ".xls", directory,
                     mat_mail.attachments)

        if "_WR_ISIR_COR_ASSESSMENT_" in query and (year in query[:-8]) :
            do_query(query, date + " ISIR Correction Assessment " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_ISIR_CORR_REJECT_" in query and (year in query[:-8]) :
            do_query(query, date + " ISIR Correction Rejected " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_ISIR_DGR_ANSW_CHNG_" in query and (year in query[:-8]) :
            do_query(query, date + " ISIR Degree Answer Change " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_ISIR_DEP_STAT_PRB_" in query and (year in query[:-8]) :
            do_query(query, date + " ISIR Dependency 20" + year + ".xls", directory,
                     mat_mail.attachments)

        if "_WR_ISIR_REJECTED_CORR_" in query and (year in query[:-8]) :
            do_query(query, date + " ISIR Rejected Corrections 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if ("UUFA_WR_ISIR_REJECT_CODES_" in query and year in query[:-8]) \
                | (("UUFA_WR_ISIR_REJECT_CODES_20") in query and (year in query[:-8])) :
            do_query(query, date + " Rejected ISIRs 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_ISIR_SS_MCH_NOT_CON_" in query and (year in query[:-8]) :
            do_query(query, date + " SS Match Not Confirmed 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_ISIR_SUSPENSE_" in query and (year in query[:-8]) :
            do_query(query, date + " ISIR Suspense " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_LEGAL_ALIEN_WORK_" in query and (year in query[:-8]) :
            do_query(query, date + " Legal Alien Work 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_LN_ACCPT_STAF_31_32_" in query and (year in query[:-8]) :
            do_query(query, date + " Stafford Accept Offer " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_LOAN_CENSUS_DATE_" in query and (year in query[:-8]) :
            do_query(query, date + " Loans Census Date 20" + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_LN_FA907_1_REVISED_" in query and (year in query[:-8]) :
            do_query(query, date + " Loan Disbursed Report " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_LN_FA907_2_REVISED_" in query and (year in query[:-8]) :
            do_query(query, date + " Loan Not Disbursed Report " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "R_LOAN_ORIG_DEPT_REVIEW_" in query and (year in query[:-8]) :
            do_query(query, date + " Loan ORIG DEPT RVW " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_LN_SENT_NO_RESPONSE_" in query and (year in query[:-8]) :
            do_query(query, date + " Loan Sent No Response " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_LOAN_TRANSMIT_HOLD_" in query and (year in query[:-8]) :
            do_query(query, date + " Loan Transmit Hold " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_LW_MD_DN_AW_NO_DISB_" in query and (year in query[:-8]) :
            do_query(query, date + " LW MD DN Awards Not Disbursed " + year + ".xls", directory,
                     prof_kkt_mail.attachments)

        if "_WR_MNTGMR_AMCORP_OVRAW_" in query and (year in query[:-8]) :
            do_query(query, date + " Montgomery Americorp Overaward " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_MULTIPLE_EMPLIDS_" in query and (year in query[:-8]) :
            do_query(query, date + " Multiple EMPLIDS 20" + year + ".xls", directory,
                     sys_mail.attachments)

        if "_WR_NO_COMMENT_CODE_" in query and (year in query[:-8]) :
            do_query(query, date + " Sub ISIR Checklist No ISIR Comment Code 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_NSLDS_LOAN_DATA_" in query and (year in query[:-8]) :
            do_query(query, date + " NSLDS Loan Data .xls", directory,
                     loans_kkt_mail.attachments)

        if "_WR_OVRD_ACAD_LVL_" in query and (year in query[:-8]) :
            do_query(query, date + " FA Term Override Acad Level " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PA_EXPECT" in query and (year in query[:-8]) :
            do_query(query, date + " PA Expected Grad Date Blank " + year + ".xls", directory,
                     prof_kkt_mail.attachments)

        if "_WR_PA_FDEG_CHECKLIST" in query:
            do_query(query, date + " PA MPS FDEG Checklist" + year + "U.xls", directory,
                     prof_mail.attachments)

        if "_WR_PELL_AWRD_LOCK_" in query and (year in query[:-8]) :
            do_query(query, date + " Pell Award Lock No FPEL" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PELL_OVERPAYMENT_" in query and (year in query[:-8]) :
            do_query(query, date + " Pell Ovpy Check NSLDS 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PELL_SUMMER_NO_PELL_" in query and (year in query[:-8]) :
            do_query(query, date + " Pell Summer No Pell 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PELL_TERM_FT_" in query and (year in query[:-8]) :
            do_query(query, date + " Term Pell Awards FT 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PELL_TERM_HT_" in query and (year in query[:-8]) :
            do_query(query, date + " Term Pell Awards HT 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PELL_TERM_LH_" in query and (year in query[:-8]) :
            do_query(query, date + " Term Pell Awards LH 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PELL_TERM_NL_" in query and (year in query[:-8]) :
            do_query(query, date + " Term Pell Awards NL 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PELL_TERM_TQ_" in query and (year in query[:-8]) :
            do_query(query, date + " Term Pell Awards TQ 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PERK_SPLIT_MISMATCH_" in query and (year in query[:-8]) :
            do_query(query, date + " Perkins Plan Split Mismatch " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_PRK_LN_ACAD_LVL_CHG_" in query and (year in query[:-8]) :
            do_query(query, date + " Perkins Awd With Acad Lvl Change " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_QUALITY_ASSURANCE_" in query and (year in query[:-8]) :
            do_query(query, date + " QA Students Complete Verification 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "R_RT4_DROPPED_CLASSES_" in query and (year in query[:-8]) :
            do_query(query, date + " RT4 Dropped Classes 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if "_WR_SCH_NOT_DISB_" in query and (year in query[:-8]) :
            do_query(query, date + " Cash Non-Cash Sch Not Disb " + year + ".xls", directory,
                     hjjr_mail.attachments)

        if "_WR_SCHOLAR_TBP_NO_AWRD" in query and year in query[:-8]:
            do_query(query, date + " TBP NO Award " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "_WR_GRAD_FELLOW" in query and year in query[:-8]:
            do_query(query, date + " Grad Fellowship " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "_WR_SSR_MATCH_NOT_CNFRM_" in query and (year in query[:-8]) :
            do_query(query, date + " SSR Not Confirmed 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SSR_NOT_CNFRMD_VTRN_" in query and (year in query[:-8]) :
            do_query(query, date + " VA Match SSR DB Override " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SS_DB_OVERRIDE_" in query and (year in query[:-8]) :
            do_query(query, date + " SS DB Override " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SUB_ISIR_PACKAGED_" in query and (year in query[:-8]) :
            do_query(query, date + " Subsequent ISIR Packaged 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SUB_ISIR_REAWD_AID_" in query and (year in query[:-8]) :
            do_query(query, date + " Canceled FCOR Complete " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SUB_ISIR_SYSG_" in query and (year in query[:-8]) :
            do_query(query, date + " Subsequent ISIR System Generated 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SUB_ISIR_VERIFIED_" in query and (year in query[:-8]) :
            do_query(query, date + " Subsequent ISIR Verified 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SUMMER_NO_DL_" in query and (year in query[:-8]) :
            do_query(query, date + " Summer Enroll No DL " + year + ".xls", directory,
                     loans_kkt_mail.attachments)

        if "_WR_SUMMER_PELL_LTHT" in query and (year in query[:-8]) :
            do_query(query, date + " Summer Pell LTHT " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SSP_DOB_PRB_APPLCNT_" in query and (year in query[:-8]) :
            do_query(query, date + " Suspense Applicant DOB Problem " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SSP_NAME_PRB_APLCNT_" in query and (year in query[:-8]) :
            do_query(query, date + " Suspense Applicant Name Problem " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_SSP_SSN_PRB_APLCNT_" in query and (year in query[:-8]) :
            do_query(query, date + " Suspense Applicant SSN Problem " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_TERM_NSLDS_LOAN_YR_" in query and (year in query[:-8]) :
            do_query(query, date + " NSLDS Loan Year Blank " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_TITLE_VII_MED_LOANS_" in query and (year in query[:-8]) :
            do_query(query, date + " Title VII Medical Loans TILA 20" + year + ".xls", directory,
                     prof_kkt_mail.attachments)

        if "_WR_TRANSFER_ENT_CNS_" in query and (year in query[:-8]) :
            do_query(query, date + " Transfer Students Entrance Counseling 20" + year + ".xls", directory,
                     loans_kkt_mail.attachments)

        if "_WR_TRANSFER_STU_FA_SP_" in query and (year in query[:-8]) :
            do_query(query, date + " Transfer Students Fall-Spring 20" + year + ".xls", directory,
                     buttars_mail.attachments)

        if "_WR_UG_GR_DIR_LN_GR_TRM_" in query and (year in query[:-8]) :
            do_query(query, date + " UG-GR Direct Ln Grad Term " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_UG_GR_PLUS_GR_TERM_" in query and (year in query[:-8]) :
            do_query(query, date + " UG-GR PLUS Grad Term " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "WR_UAC_FASI_STATUS_" in query and (year in query[:-8]) :
            do_query(query, date + " UAC FASI Status " + year + ".xls", directory,
                     kkt_mail.attachments)

        if "_WR_UAC_SNGDO_" in query and (year in query[:-8]) :
            do_query(query, date + " UAC SNGDO Campus 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_UNDOCUMENTED_STUDENTS_" in query and (year in query[:-8]) :
            do_query(query, date + " Undocumented Student Awards " + year + ".xls", directory,
                     kkt_mail.attachments)

        if "_WR_VERI_CHKLST_MISSING_" in query and (year in query[:-8]) :
            do_query(query, date + " Verification Checklist Missing 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_VERI_INCOME_ADJ_" in query and (year in query[:-8]) :
            do_query(query, date + " Income Adjustments 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_VER_NOT_CONSL_" in query and (year in query[:-8]) :
            do_query(query, date + " Verification Not Consolidated 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_VETERAN_ACTIVE_DUTY_" in query and (year in query[:-8]) :
            do_query(query, date + " Veteran Active Duty 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_VETERAN_NO_QUALIFY_" in query and (year in query[:-8]) :
            do_query(query, date + " Veteran No Qualify 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_WEEKS_OF_INSTR_FIX_" in query and (year in query[:-8]) :
            do_query(query, date + " Weeks of Instruction 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "_WR_DL_AY_SP_CANCELED_" in query and (year in query[:-8]) :
            do_query(query, date + " DL AY SP Cancelled " + year + ".xls", directory,
                     loans_kkt_mail.attachments)

        if "_WR_LOAN_TRANSMIT_HOLD_13" in query:
            do_query(query, date + " Loan Transmit Hold 13.xls", directory,
                     loans_kt_mail.attachments)

        if "_WR_REJECT_CODE_ON_ISIR_" in query and (year in query[:-8]) :
            do_query(query, date + " Rejected ISIR's " + year + ".xls", directory,
                     kmt_mail.attachments)

        if "_WR_RT4_FA_DROP_CLASSES_" in query and (year in query[:-8]) :
            do_query(query, date + " RT4 Fall Drop Classes 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if "_WR_RT4_SP_DROP_CLASSES_" in query and (year in query[:-8]) :
            do_query(query, date + " RT4 Spring Drop Classes 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if "_WR_RT4_SU_DROP_CLASSES_" in query and (year in query[:-8]) :
            do_query(query, date + " RT4 Summer Drop Classes 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if "WR_AWARDS_OTHER_INST" in query and (year in query[:-8]) :
            do_query(query, date + " Checklist FAOI" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "WR_DEP_PRNT_SSN_RVW" in query and (year in query[:-8]) :
            do_query(query, date + " Parent SSN Review " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "WR_DN_LW_MD_AID_ATRB" in query and (year in query[:-8]) :
            do_query(query, date + " DN-LW-MD Student Aid Career " + year + ".xls", directory,
                     prof_kkt_mail.attachments)

        if "WR_FSEOG_NO_PELL" in query and (year in query[:-8]) :
            do_query(query, date + " DNFSEOG no Pell " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "WR_FT_CLASS_OVERRIDES" in query and (year in query[:-8]) :
            do_query(query, date + " Class Overrides " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "WR_GR_ACAD_LV_OUT_SYNC" in query and (year in query[:-8]) :
            do_query(query, date + " GR Academic Levels Out of Sync " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "WR_PELL_COA_DOUBLE" in query and (year in query[:-8]) :
            do_query(query, date + " PELL COA Double  " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "WR_PKG_AWD_NO_BDGT" in query and (year in query[:-8]) :
            do_query(query, date + " Award NO Budget for Term " + year + ".xls", directory,
                     kkt_mail.attachments)

        if "WR_SCH_TUITION_FEES" in query and (year in query[:-8]) :
            do_query(query, date + " Waiver-Scholar Tuition Fees " + year + ".xls", directory,
                     hj_mail.attachments)

        if "WR_SCHOL_GRAD_DATE" in query and (year in query[:-8]) :
            do_query(query, date + " Scholarship-Expected Grad Date  " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "_WR_STDNT_NOT_PACKAGED" in query and (year in query[:-8]) :
            do_query(query, date + " Students Not Packaged " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "WR_SAVE_CTZNSHIP_VER" in query and (year in query[:-8]) :
            do_query(query, date + " SAVE SB81 CTZNSHP VERI " + year + ".xls", save_directory,
                     bk_mail.attachments)

        if "WR_ALL_C_FVRA_I_NO_COR" in query and (year in query[:-8]) :
            do_query(query, date + " FVRA Ini with no Correction " + year + ".xls", save_directory,
                     kkmt_mail.attachments)
            
        if "WR_CORRECTION_NOT_SENT" in query and (year in query[:-8]) :
            do_query(query, date + " Corrections Sent still Pending " + year + ".xls", save_directory,
                     kkmt_mail.attachments)

        # Manually run Queries
        if "_WR_LOAN_EFT_DETAIL_ERROR" in query:
            do_query(query, date + " Loan EFT Detail Error.xls", directory,
                     loans_kkt_mail.attachments)

        if "_WR_NSL_PROMISSORY_NOTE_" in query and (year in query[:-8]) :
            do_query(query, date + " NSL Promissory Note " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "UUFA_AP_RPKG_FPEL_AWARD_LCK" in query:
            do_query(query, date + " Pell Award Lock FPEL" + year + ".xls", directory,
                     loans_kkt_mail.attachments)

        if "WR_FAVR_I_NO_FCOR_ALL_C" in query:
            do_query(query, date + " FAVR I No FCOR all check Completed.xls", directory,
                     loans_kkt_mail.attachments)

        # Packaging queries that are being manually run.
        if "PRT_ATH_ACCEPT_FED_AID_" in query and (year in query[:-8]) :
            do_query(query, date + " Athlete Accepted Federal Aid " + year + ".xls", packaging_directory,
                     ath_mail.attachments)

        if "PRT_ATH_AWD_CBA_GRANT_" in query and (year in query[:-8]) :
            do_query(query, date + " Athlete Awarded CBA Grant " + year + ".xls", packaging_directory,
                     ath_mail.attachments)

        if "PRT_ATHLETE_GRAD_DATE_" in query and (year in query[:-8]) :
            do_query(query, date + " Athlete Accepted Grad Date " + year + ".xls", packaging_directory,
                     ath_mail.attachments)

        if "PRT_ATH_OFFERED_FED_AID_" in query and (year in query[:-8]) :
            do_query(query, date + " Athlete Offered Federal Aid " + year + ".xls", packaging_directory,
                     ath_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Weekly Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Budget Queries
def do_budget_queries():
    global aid_year
    global year
    for query_name in os.listdir("."):
        if "BR_BDGT_" in query_name:
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    # Create FOLDER variables to be used in Move() operation and establishes
    # the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Budgets', aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Budgets', aid_year, month_folder))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be query AC it is received and _new_file_name to what
    # the new query should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_BR_ACAD_LVLS_OUT_SYNC_") and (year in query[:-8]) :
            do_query(query, date + " GR Academic Levels Out of Sync " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_BR_ATH_TUIT_INCR_NR_") and (year in query[:-8]) :
            do_query(query, date + " Athlete Tuition Increase NR " + year + ".xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_BR_ATH_TUITION_INCRS_") and (year in query[:-8]) :
            do_query(query, date + " Athlete Tuition Increase " + year + ".xls", directory,
                     ath_mail.attachments)

        if "_BR_BDGT_DOUBLE_BUDGETS_" in query and (year in query[:-8]) :
            do_query(query, date + " Double Budget " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_BR_COA_LESS_HT_") and (year in query[:-8]) :
            do_query(query, date + " PELL COA Less Than Half Time Enrollment " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_BR_COA_TUIT_ZERO_") and (year in query[:-8]) :
            do_query(query, date + " COA Tuition Amount Zero " + year + ".xls", directory,
                     kktv_mail.attachments)

        if query.startswith("UUFA_BR_DN_LW_MD_AID_ATRB_") and (year in query[:-8]) :
            do_query(query, date + " DN-LW-MD Student Aid Career " + year + ".xls", directory,
                     prof_mail.attachments)

        if query.startswith("UUFA_BR_FT_CLASS_OVERRIDES_") and (year in query[:-8]) :
            do_query(query, date + " Class Overrides " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_BR_COA_ISIR_BDGT_DIFF_") and (year in query[:-8]) :
            do_query(query, date + " COA ISIR Budget Mismatch " + year + ".xls", directory,
                     ath_kkmt_mail.attachments)

        if query.startswith("UUFA_BR_NO_BUDGET_ATTEND_") and (year in query[:-8]) :
            do_query(query, date + " NO Budget Attend 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_BR_PELL_COA_BLANK_") and (year in query[:-8]) :
            do_query(query, date + " PELL COA Blank " + year + ".xls", directory,
                     acnt_kkt_mail.attachments)

        if (("BR_PELL_COA_CHECK" in query) and (year in query[:-8])) :
            do_query(query, date + " PELL COA Fix - COD Review " + year + ".xls", directory,
                     acnt_kkt_mail.attachments)
        

        if query.startswith("UUFA_BR_PELL_COA_DOUBLE_") and (year in query[:-8]) :
            do_query(query, date + " PELL COA Double " + year + ".xls", directory,
                     kkt_mail.attachments)

        if "BR_PELL_COA_DBLD_WRNG" in query and (year in query[:-8]) :
            do_query(query, date + " PELL COA Double " + year + ".xls", directory,
                     null_mail.attachments)

        if query.startswith("FA_BR_PELL_COA_LESS_HT_20") and (year in query[:-8]) :
            do_query(query, date + " PELL COA Less HT Enrollment " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_BR_PROC_STAT_RVW_STAT_") and (year in query[:-8]) :
            do_query(query, date + " Reset Processing Status to 1 " + year + ".xls", directory,
                     ml_mail.attachments)

        if query.startswith("UUFA_BR_RES_NON_RES_BDGT_") and (year in query[:-8]) :
            do_query(query, date + " Resident - Non-Resident Budget " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_BR_SCH_TUITION_FEES_NR_") and (year in query[:-8]) :
            do_query(query, date + " Waiver-Scholar Tuition Fees NR " + year + ".xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_BR_SCH_TUITION_ONLY_NR_") and (year in query[:-8]) :
            do_query(query, date + " Waiver-Scholar Tuition Only NR " + year + ".xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_BR_SCHOLAR_TUIT_FEES_") and (year in query[:-8]) :
            do_query(query, date + " Waiver-Scholar Tuition Fees Res " + year + ".xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_BR_SCHOLAR_TUIT_ONLY_") and (year in query[:-8]) :
            do_query(query, date + " Waiver-Scholar Tuition Only Res " + year + ".xls", directory,
                     hj_mail.attachments)

        if query.startswith("FA_BR_UFORM_CHANGE_BUD_DUR_") and (year in query[:-8]) :
            do_query(query, date + " Correct Budget Duration " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "BR_ACAD_LVLS_NOT_SYNC" in query and (year in query[:-8]) :
            do_query(query, date + " TRIAL Acad Levels out of SYNC " + year + ".xls", directory,
                     kkmt_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Budget Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]
    
#Budget Testing Queries
def do_budget_test_queries():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_BUDGET_"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year

    # Create FOLDER variables to be used in Move() operation and establishes
    # the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Budgets', aid_year, month_folder,"Wrong Budget Queries"))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Budgets', aid_year, month_folder, "Wrong Budget Queries"))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be query AC it is received and _new_file_name to what
    # the new query should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_BUDGET_20" + year + "_DN1"):
            do_query(query, date + " Wrong Budget - Dental 1.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_DN2"):
            do_query(query, date + " Wrong Budget - Dental 2.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_DN3"):
            do_query(query, date + " Wrong Budget - Dental 3.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_DN4"):
            do_query(query, date + " Wrong Budget - Dental 4.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_ACCTMAC"):
            do_query(query, date + " Wrong Budget - Accounting Masters.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_ARCHMAR"):
            do_query(query, date + " Wrong Budget - Architect Masters.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_BUSINESS"):
            do_query(query, date + " Wrong Budget - Grad Business.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_COMDIS"):
            do_query(query, date + " Wrong Budget - COMDIS.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_EAEMS"):
            do_query(query, date + " Wrong Budget - EAE.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_EDUCATION"):
            do_query(query, date + " Wrong Budget - Education.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_ED_PSYCH"):
            do_query(query, date + " Wrong Budget - ED Psychology.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_ENGINERING"):
            do_query(query, date + " Wrong Budget - Grad Engineering.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_FINE_ARTS"):
            do_query(query, date + " Wrong Budget - Fine Arts .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_GENERAL_GR"):
            do_query(query, date + " Wrong Budget - General Grad .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_GENETICS"):
            do_query(query, date + " Wrong Budget - Genetics.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_HEALTH"):
            do_query(query, date + " Wrong Budget - Health.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PRO_HEALTH"):
            do_query(query, date + " Wrong Budget - Health Promotion .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_HUMANITIES"):
            do_query(query, date + " Wrong Budget - Humanities.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_MBA_BUADMB"):
            do_query(query, date + " Wrong Budget - MBA - BUADMBA .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_MD_SCIENCE"):
            do_query(query, date + " Wrong Budget - Medical Science .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_MEDICAL"):
            do_query(query, date + " Wrong Budget - MD Graduate .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_NURSING"):
            do_query(query, date + " Wrong Budget - Graduate Nursing .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PA"):
            do_query(query, date + " Wrong Budget - Physician Assistant .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PHARMACY"):
            do_query(query, date + " Wrong Budget - Pharmacy .xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PLANNING"):
            do_query(query, date + " Wrong Budget - Planning.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PMBAMBA"):
            do_query(query, date + " Wrong Budget - PMBAMBA.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PUBLICPOLI"):
            do_query(query, date + " Wrong Budget - Public Policy.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_PUBLIC_ADM"):
            do_query(query, date + " Wrong Budget - Public Administration.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_SCIENCE"):
            do_query(query, date + " Wrong Budget - Science.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_SOC_BEHAV"):
            do_query(query, date + " Wrong Budget - Social and Behavioral.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_SW"):
            do_query(query, date + " Wrong Budget - Social Work.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_GR_XMBAMBA"):
            do_query(query, date + " Wrong Budget - XMBAMBA.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_LW1"):
            do_query(query, date + " Wrong Budget - Law 1.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_LW2"):
            do_query(query, date + " Wrong Budget - Law 2.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_LW3"):
            do_query(query, date + " Wrong Budget - Law 3.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_MD1"):
            do_query(query, date + " Wrong Budget - Med 1.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_MD2"):
            do_query(query, date + " Wrong Budget - Med 2.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_MD3"):
            do_query(query, date + " Wrong Budget - Med 3.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_MD4"):
            do_query(query, date + " Wrong Budget - Med 4.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_UNDERGRAD"):
            do_query(query, date + " Wrong Budget - Undergraduate.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_BUSINESS"):
            do_query(query, date + " Wrong Budget - Undergraduate Business.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_BUS_LTHT"):
            do_query(query, date + " Wrong Budget - Undergraduate Business LTHT.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_ENGINERING"):
            do_query(query, date + " Wrong Budget - Undergraduate Engineering.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_ENG_LTHT"):
            do_query(query, date + " Wrong Budget - Undergraduate Engineering LTHT.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_LTHT"):
            do_query(query, date + " Wrong Budget - Undergraduate LTHT.xls", directory,
                     null_mail.attachments)
            
        if query.startswith("UUFA_BUDGET_20" + year + "_UG_NURSE_LTHT"):
            do_query(query, date + " Wrong Budget - Undergraduate Nursing LTHT.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_BUDGET_20" + year + "_UG_NURSING"):
            do_query(query, date + " Wrong Budget - Undergraduate Nursing.xls", directory,
                     null_mail.attachments)

#Packaging Queries
def do_packaging_queries():
    global aid_year
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_PRT_ACAD_PROG_REVIEW"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    # Create FOLDER variables to be used in Move() operation and establishes
    # the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Packaging', aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Packaging', aid_year, month_folder))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be query AC it is received and _new_file_name to what
    # the new query should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_PRT_ACAD_LVLS_OUT_SYNC"):
            do_query(query, date + " UG Acad Levels Out of Sync.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_ACAD_PROG_REVIEW_") and (year in query[:-8]) :
            do_query(query, date + " Academic Progress Review.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_ATH_ACCEPT_FED_AID"):
            do_query(query, date + " Athlete Accepted Federal Aid.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_PRT_ATH_AWD_CBA_GRANT"):
            do_query(query, date + " Athlete Awarded CBA Grant.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_PRT_ATH_GRAD_DATE"):
            do_query(query, date + " Athlete Expected Grad Date.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_PRT_ATH_OFFRD_FED_AID_"):
            do_query(query, date + " Athlete Offered Federal Aid.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_PRT_ATH_OFFR_ACCPT_AID_") and (year in query[:-8]) :
            do_query(query, date + " ATH Fed State Inst O-A.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_PRT_AWARD_TERM_HAD_SAP"):
            do_query(query, date + " SAP Hold Term Award Review.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_AWARDS_OTHER_INST"):
            do_query(query, date + " Checklist FAOI" + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_AWD_CMB_OVR_AG_RVW_"):
            do_query(query, date + " Award Combined Over Aggregate.xls", directory,
                     loans_kt_mail.attachments)

        if query.startswith("FA_PRT_AWD_MASS_P_NO_AWARDS"):
            do_query(query, date + " Award Mass Packaging No Awards.xls", directory,
                     ml_mail.attachments)

        if query.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_") and (year in query[:-8]) :
            do_query(query, date + " PELL ELIGIBLE NO PELL 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_AWD_SUB_OVR_AG_RVW"):
            do_query(query, date + " SUB Over Aggregate.xls", directory,
                     loans_kkt_mail.attachments)

        if query.startswith("UUFA_PRT_CTZN_IND_AWD_NO_LN"):
            do_query(query, date + " LA-wk eligible - Award No Loans.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PRT_DEFR_ENROLLMENT_") and (year in query[:-8]) :
            do_query(query, date + " DEFER Enrollment " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PRT_DEP_PRNT_SSN_RVW"):
            do_query(query, date + " Parent SSN Review.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PRT_DIAG_AWD_PELL_TERM_") and (year in query[:-8]) :
            do_query(query, date + " Term Pell Awards 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("FA_PRT_DISB_PLAN_SPLT_CODE_") and (year in query[:-8]) :
            do_query(query, date + " Disb Plan FY Split Code XX.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_DL_DPAY_SCSP_") and (year in query[:-8]) :
            do_query(query, date + " Disb Plan AY-Split Code SP.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_DSB_PLN_SPLT_CD_FD_") and (year in query[:-8]) :
            do_query(query, date + " Federal Disb Plan FY Split Code XX " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_DSB_PLN_SPLT_CD_SC_") and (year in query[:-8]) :
            do_query(query, date + " Scholarship Disb Plan FY Split Code XX " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_EXPECT_GRAD_TERM_11"):
            do_query(query, date + " Expected Grad Term 1" + str(int(year) - 1) + "8.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PRT_DL_GRAD_TERM_FALL_") and (year in query[:-8]) :
            do_query(query, date + " DL Expected Grad Term Fall " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_GRAD_TRM_FALL_") and (year in query[:-8]) :
            do_query(query, date + " Loan Proration Grad Term Fall " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_GRAD_TRM_SPRING_") and (year in query[:-8]) :
            do_query(query, date + " Loan Proration Grad Term Spring " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_HEAL_"):
            do_query(query, date + " Heal 20 " + year + ".xls", directory,
                     loans_kkt_mail.attachments)

        if query.startswith("UUFA_PRT_LEU_C_PELL_FSEOG_") and (year in query[:-8]) :
            do_query(query, date + " LEU C Flag Awarded FSEOG " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_LEU_C_PELL_AWARD_") and (year in query[:-8]) :
            do_query(query, date + " LEU C Flag Awarded Pell " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_LEU_E_PELL_FSEOG_") and (year in query[:-8]) :
            do_query(query, date + " LEU E Flag Awarded FSEOG " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_LN_CBA_AWD_NO_ELIG"):
            do_query(query, date + " Loan CBA Review Eligible.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PRT_NO_FALL_11" + str(int(year) - 1) + "8"):
            do_query(query, date + " Packaging No Fall 1" + str(int(year) - 1) + "8 (2).xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PRT_NSL_LOAN_RPT_VERI_") and (year in query[:-8]) :
            do_query(query, date + " NSL Loan Need Verification.xls", directory,
                     loans_c_mail.attachments)

        if query.startswith("UUFA_PRT_NURSING_LOAN_RPT_") and (year in query[:-8]) :
            do_query(query, date + " NSL Needs NSL P-N Checklist " + year + ".xls", directory,
                     prof_mail.attachments)

        if query.startswith("UUFA_PRT_ON_LINE_PACKAGING"):
            do_query(query, date + " Manual Packaging Counselors.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_PELL_COMMENT_037_") and (year in query[:-8]) :
            do_query(query, date + " Pell Comment Code 037 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_PELL_ELG_NO_PELL_") and (year in query[:-8]) :
            do_query(query, date + " PELL ELIGIBLE NO PELL " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_PLL_EL_CTZN_NOT_INDCT"):
            do_query(query, date + " Pell Eligible Citizenship Not Indicated.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PRT_PELL_FPEL" + year + "_INITIATED"):
            do_query(query, date + " Pell FPEL" + year + " Initiated.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_PELL_UG_5TH_YR_2ND_BA"):
            do_query(query, date + " Pell UG 5th YR.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_PHARM_NO_HEAL"):
            do_query(query, date + " Pharmacy students with NO HEAL.xls", directory,
                     loans_kkt_mail.attachments)

        if query.startswith("UUFA_PRT_PKG_AWD_NO_BDGT"):
            do_query(query, date + " Award NO Budget for Term.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PRT_PKG_SCH_AWD_NO_BGT_") and (year in query[:-8]) :
            do_query(query, date + " Scholarship Award NO Budget for Term.xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_PRT_PRIOR_TERM_STFFRD_OFR"):
            do_query(query, date + " Cancel Prior Term Stafford Offer " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_READY_PKG_20" + year):
            do_query(query, date + " Students Ready to Package " + year + ".xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_PRT_SCHOL_GRAD_DATE"):
            do_query(query, date + " Scholarship-Expected Grad Date.xls", directory,
                     jj_mail.attachments)

        if query.startswith("UUFA_PRT_SUB_UNSUB_SP_SP"):
            do_query(query, date + " Loan Offrd Disb Plan SP Split Code SP.xls", directory,
                     loans_kt_mail.attachments)

        if query.startswith("UUFA_PRT_SET_HEAL_FLAG_") and (year in query[:-8]) :
            do_query(query, date + " MD - Pharmacy Heal Eligible Flag.xls", directory,
                     loans_kkt_mail.attachments)

        if query.startswith("UUFA_PRT_STATE_OF_RES_FM_MH_PW"):
            do_query(query, date + " State of Residence FM MH PW.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("FA_PRT_STILL_UNPRCD_AFTER_PKG"):
            do_query(query, date + " Students Not Packaged (old).xls", directory,
                     ml_mail.attachments)

        if query.startswith("UUFA_PRT_STDNT_NOT_PACKAGED_") and (year in query[:-8]) :
            do_query(query, date + " Students Not Packaged " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_PRT_TEACH_CREDENTIAL_") and (year in query[:-8]) :
            do_query(query, date + " Teach Credential 20" + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("PKG_NOT_PKGD_NO_UNITS") and (year in query[:-8]) :
            do_query(query, date + " Not Packaged No Units " + year + ".xls", directory,
                     kkt_mail.attachments)

        if ("PRT_SUB_UNSUB_FA_FA") in query and (year in query[:-8]) :
            do_query(query, date + " Loan Offrd Disb Plan FA Split Code FA.xls" + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if ("PRT_PELL_PKG_LOAD_CHCK_") in query and (year in query[:-8]) :
            do_query(query, date + " Pell Package Load Check " + year + ".xls", directory,
                     kkt_mail.attachments)

        if ("PRT_COUNT_ITEM_TYPE") in query and (year in query[:-8]) :
            do_query(query, date + " Aid Awarded No Funds " + year + ".xls", directory,
                     bekkt_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Batch Packaging Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]
            
#Monthly Queries
def do_monthlies():
    global aid_year
    year = "18"
    for query_name in os.listdir("."):
        if "MR_PELL_SSN_MISMATCH" in query_name:
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    t_path = "Award Summary 20" + year + "/Award Summary " + calendar.month_name[last_month] + " " + str(last_months_year)

    # Create FOLDER variables to be used in Move() operation and establishes
    # the directory to save the files
    # using the date.
    # Create variables to be used in Move() operation.
    if test:
        directory = os.path.realpath(os.path.join('C:/Testing Bob/Monthly', month_folder))
        dl_directory = os.path.realpath(os.path.join('C:/Testing Bob/Direct Loans', 'Monthly'))
        acct_directory = os.path.realpath(os.path.join('C:/Testing Bob/ACCT/Chartfields'))
        t_directory = os.path.realpath(os.path.join('C:/Testing Bob/ACCT/Award Summary', t_path))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/Monthly', month_folder))
        dl_directory = os.path.realpath(os.path.join('O:/Systems/Direct Loans','Monthly'))
        acct_directory = os.path.realpath(os.path.join('O:/ACCT/Chartfields'))
        t_directory = os.path.realpath(os.path.join('O:/ACCT/Award Summary', t_path))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(dl_directory):
        os.makedirs(dl_directory)
    if not os.path.isdir(acct_directory):
        os.makedirs(acct_directory)
    if not os.path.isdir(t_directory):
        os.makedirs(t_directory)

    # Change File_Name to be file ac it is received and _new_file_name to what
    # the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if "MR_COMMENT_CODE_298_" in query and year in query[:-8]:
            do_query(query, date + " IASG - Pell Eligible " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_3RD_PARTY_CROSSWALK" in query and year in query[:-8]:
            do_query(query, date + " Third Party Crosswalk " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "MR_3RD_PRT_MNTR_IA_ALL" in query and year in query[:-8]:
            do_query(query, date + " Third Party Monitor " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "MR_ACAD_LVLS_NOT_SYNC" in query and year in query[:-8]:
            do_query(query, date + " Academic Levels out of SYNC " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_ADM_DEFERRAL" in query and year in query[:-8]:
            do_query(query, date + " FA Admission Deferral " + year + ".xls", directory,
                     hjj_mail.attachments)

        if "MR_ALT_LN_TRNSMIT_HOLD" in query and year in query[:-8]:
            do_query(query, date + " Alt Loan Transmit Hold " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "MR_ATHLETE_T53_AWARDS" in query and year in query[:-8]:
            do_query(query, date + " Athlete T53 Awards Accepted " + year + ".xls", directory,
                     ath_mail.attachments)

        if "MR_COD_DL" in query and year in query[:-8]:
            do_query(query, date + " COD DL FATB" + year + " FCRD" + year + " FHMS" + year + ".xls", directory,
                     acnt_kkt_mail.attachments)

        if "MR_COD_PELL_TEACH_IASG" in query and year in query[:-8]:
            do_query(query, date + " COD Grant FCRD" + year + "-FHMS" + year + " Report.xls", directory,
                     acnt_kkt_mail.attachments)

        if "MR_DIR_LN_TRNSMIT_HOLD" in query and year in query[:-8]:
            do_query(query, date + " COD Direct Loan Transmit Hold " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "MR_DISB_ATH_AWD_NOPOST" in query and year in query[:-8]:
            do_query(query, date + " Athlete Waiver Disbursed Not Posted " + year + ".xls", directory,
                     ath_mail.attachments)

        if "MR_DL_DISB_FAILED" in query and year in query[:-8]:
            do_query(query, date + " DL Disbursement Failed " + year + ".xls", dl_directory,
                     krms_mail.attachments)

        if "MR_DL_ORIG_AWARD" in query and year in query[:-8]:
            do_query(query, date + " DL ORIG Award " + year + ".xls", directory,
                     loans_c_mail.attachments)

        if "MR_DSB_CASH_AWD_NOPOST" in query and year in query[:-8]:
            do_query(query, date + " Cash Disbursed Not Posted " + year + ".xls", directory,
                     hjj_mail.attachments)

        if "MR_DSB_WAVR_AWD_NOPOST" in query and year in query[:-8]:
            do_query(query, date + " Waiver-Scholarship Disbursed Not Posted " + year + ".xls", directory,
                     hj_mail.attachments)

        if "DSB_AWD_NOPOST" in query and year in query[:-8]:
            do_query(query, date + " Award Disbursed Not Posted " + year + ".xls", directory,
                     hjjk_mail.attachments)

        if "MR_DN_INC_CHECKLISTS" in query and year in query[:-8]:
            do_query(query, date + " Dental Students with I Checklists " + year + ".xls", directory,
                     prof_mail.attachments)

        if "MR_M_L_D_INI_CHECKLIST" in query and year in query[:-8]:
            do_query(query, date + " L-M-D Students with I Checklists " + year + ".xls", directory,
                     prof_mail.attachments)

        if "MR_FWS_WITH_NSI_HOLD" in query and year in query[:-8]:
            do_query(query, date + " FWS with NSI Holds.xls", directory,
                     acnt_mail.attachments)

        if "MR_GRAD_TERM_PRB" in query and year in query[:-8]:
            do_query(query, date + " Grad Term Wrong " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_ITEM_CHARTFLD_SETUP" in query and year in query[:-8]:
            do_query(query, date + " Item Chartfield Setup.xls", acct_directory,
                     bhjkn_mail.attachments)

        if "MR_ITEM_TYPE_DISB_RULE" in query and year in query[:-8]:
            do_query(query, date + " Item Type Career - Match Disb Rule Career " + year + ".xls", directory,
                     ahjmv_mail.attachments)

        if "MR_LAW_INC_CHECKLISTS" in query and year in query[:-8]:
            do_query(query, date + " Law Students with I Checklists " + year + ".xls", directory,
                     prof_mail.attachments)

        if "MR_LOAN_AWD_PARTL_DISB" in query and year in query[:-8]:
            do_query(query, date + " Loan Awards Partial Disbursed " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "MR_MED_INC_CHECKLISTS" in query and year in query[:-8]:
            do_query(query, date + " Med Students with I Checklists " + year + ".xls", directory,
                     prof_mail.attachments)

        if "MR_MED_LAW_LVL_REVIEW" in query or "_MR_DN_LW_MD_LVL_RVW" in query and year in query[:-8]:
            do_query(query, date + " MED-LAW Academic Level Review " + year + ".xls", directory,
                     prof_mail.attachments)

        if "MR_PART_TW_OTHER_SCH" in query and year in query[:-8]:
            do_query(query, date + " Partial TW Other Scholarship " + year + ".xls", directory,
                     hj_mail.attachments)

        if "MR_PELL_AWD_ADJUSTMENT" in query and year in query[:-8]:
            do_query(query, date + " Pell Award Adjust " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_PELL_ONLY" in query and year in query[:-8]:
            do_query(query, date + " Pell Awd  Zero Grants Loans " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_PELL_SSN_MISMATCH" in query and year in query[:-8]:
            do_query(query, date + " SSN Mismatch " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_PERKINS_CLASS_LIMIT" in query and year in query[:-8]:
            do_query(query, date + " Perkins Class Limits " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_PERK_MISC_LN_CNCLD" in query and year in query[:-8]:
            do_query(query, date + " Perkins - Misc Loans Cancelled " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "MR_PERK_MISC_LOAN_DISB" in query and year in query[:-8]:
            do_query(query, date + " Perkins - Misc Loans Disbursed " + year + ".xls", directory,
                     loans_kt_mail.attachments)

        if "_MR_SCH_IT_RECON_" in query and year in query[:-8]:
            do_query(query, date + " Scholarship IT Recon " + year + ".xls", directory,
                     schol_mail.attachments)

        if "MR_SCH" in query and "LOA" in query and year in query[:-8]:
            do_query(query, date + " Scholarship LOA " + year + ".xls", directory,
                     hjj_mail.attachments)

        if "SCHOLAR_REINSTATE" in query and year in query[:-8]:
            do_query(query, date + " Scholarship Reinstate " + year + ".xls", directory,
                     hjj_mail.attachments)

        if "MR_SF_DIS_AWD_PT_ER_FC" in query and year in query[:-8]:
            do_query(query, date + " Federal Award Disb Post Error " + year + ".xls", directory,
                     loans_kkt_mail.attachments)

        if "MR_SF_DIS_AWD_PT_ER_SV" in query and year in query[:-8]:
            do_query(query, date + " SCHOL-ATH Award Disb Post Error " + year + ".xls", directory,
                     hj_mail.attachments)

        if "MR_STATE_FM_MH_PW" in query and year in query[:-8]:
            do_query(query, date + " Palau - Micronesia - Marshall Islands Students " + year + ".xls", directory,
                     acnt_kkmt_mail.attachments)

        if "MR_SUSPEND_RC2" in query and year in query[:-8]:
            do_query(query, date + " ISIR Suspended Reason Code 2 " + year + ".xls", directory,
                     kkmtv_mail.attachments)

        if "MR_UFORM_GRAD_TERM_PRB" in query and year in query[:-8]:
            do_query(query, date + " Grad Term Wrong " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "MR_UNDS_OFFER_SCHOLAR" in query and year in query[:-8]:
            do_query(query, date + " Scholarship Awards UNDS Career " + year + ".xls", directory,
                     hj_mail.attachments)

        if "MR_UNDS_OFRD_AMT_FDRL" in query and year in query[:-8]:
            do_query(query, date + " Federal Awards UNDS Career " + year + ".xls", directory,
                     ath_mail.attachments)

        if "MR_UNDS_OFRD_AMT_ATH" in query and year in query[:-8]:
            do_query(query, date + " Athlete Awards UNDS Career " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_VERIFY_DEP_OVERRIDE" in query and year in query[:-8]:
            do_query(query, date + " Verification Dependency Override " + year + ".xls", directory,
                     kkmt_mail.attachments)

        if "MR_GRBEN_EA_POST" in query and year in query[:-8]:
            do_query(query, date + " Grad Benefit EA Post " + year + ".xls", directory,
                     hjr_mail.attachments)

        if "SEFA_DL_TOTAL_AWARDS" in query and year in query[:-8]:
            do_query(query, date + " SEFA DL Amounts " + year + ".xls", directory,
                     monthly_dl_mail.attachments)

        if "SEFA_TOTAL_STUDENT" in query and year in query[:-8]:
            do_query(query, date + " SEFA DL Total Students " + year + ".xls", directory,
                     monthly_dl_mail.attachments)

        if "STP_DISB_RULE_MISMATCH" in query and year in query[:-8]:
            do_query(query, date + " Item Type Setup Mismatch " + year + ".xls", directory,
                     sys_mail.attachments)

        if "ussfa037" in query and year in query:
            do_query(query, query, t_directory,
                     null_mail.attachments)

        if "ussfa035-" in query and year in query:
            do_query(query, query, t_directory,
                     null_mail.attachments)

        if "MR_EXPIRED_MPN_CURR_LN" in query and year in query[:-8]:
            do_query(query, date + " Expired MPN Current LN " + year + ".xls", directory,
                    loans_mail.attachments)

        if "_MR_RNDM_GRAD_ENR_CHANG_" in query and year in query[:-8]:
            do_query(query, date + " Audit Graduation & Enrollment Change " + year + ".xls", directory,
                     loans_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Monthly Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]
    

#End of Term Queries
def do_end_of_term_queries():
    global aid_year
    term = 'F'
    strm = "Error"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_EOT_ACAD_PLAN_RVW"):
            prompt = "Enter Term: (e.g. 2016U, 2017F, or 2032S):"
            while True:
                strm = str.upper(raw_input(prompt))
                aid_year = strm[2:4]
                term = strm[4]
                if term == 'S' or term == 'U' or term == 'F':
                    break

    year = aid_year[-2:]
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/SAP/', "20" + year + term))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems\QUERIES/SAP/', "20" + year + term))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        if query.startswith("UUFA_EOT_ACAD_PLAN_RVW_FRAP"):
            do_query(query, date + " Academic plan RVW FRAP " + year + ".xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_EOT_CBA_ACAD_PRG_BELOW_FT"):
            do_query(query, date + " CBA Acad Prog below FT.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_EOT_CBA_UNDISBURSED"):
            do_query(query, date + " CBA UNDISBURSED.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_EOT_ALT_LOAN_SAP"):
            do_query(query, date + " Alt Loans Awards with SAP Holds.xls", directory,
                     loans_kkt_mail.attachments)

        if query.startswith("UUFA_EOT_LN_ORIG_FAIL_PENDING"):
            do_query(query, date + " Loans Originated Failed Pending.xls", directory,
                     loans_kt_mail.attachments)

        if query.startswith("UUFA_EOT_LN_ACD_PRG_BLW_HT_UND"):
            do_query(query, date + " Loan Acad Prog below HT Undisbursed.xls", directory,
                     loans_kt_mail.attachments)

        if query.startswith("UUFA_EOT_LN_ACD_PRG_BLW_HT_SUB"):
            do_query(query, date + " Loan Acad Prog below HT Subsq Disb.xls", directory,
                     loans_kt_mail.attachments)

        if query.startswith("UUFA_EOT_PRO_STDNT_SAP_WARNING"):
            do_query(query, date + " SAP Warning Professional Students.xls", directory,
                     loans_kt_mail.attachments)

        if query.startswith("UUFA_EOT_MED_SAP"):
            do_query(query, date + " Medical SAP Review.xls", directory,
                     prof_k_mail.attachments)

        if query.startswith("UUFA_EOT_FWS_WITH_NSI_HOLD"):
            do_query(query, date + " FWS with NSI Holds.xls", directory,
                     ben_mail.attachments)

        if query.startswith("UUFA_EOT_PELL_ACD_PRG_LES_THN"):
            do_query(query, date + " PELL Less Than FT.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_EOT_PELL_OFFERED_NOT_DIS"):
            do_query(query, date + " Pell Offered Not Disbursed.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_EOT_THESIS_STUDENT_NONRES"):
            do_query(query, date + " Thesis Students Non-Resident.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_EOT_SCH_ACAD_PROG_REVIEW"):
            do_query(query, date + " Scholarship Academic Progress Review " + year + ".xls", directory,
                     jja_mail.attachments)

        if query.startswith("UUFA_EOT_SCH_ACD_PRG_REVIEW"):
            do_query(query, date + " All 70000792 Item types Academic Progress Review.xls", directory,
                     jja_mail.attachments)

        if query.startswith("UUFA_EOT_SAP_AGGCP_DENTAL"):
            do_query(query, date + " SAP Aggregate Dental Career.xls", directory,
                     prof_rk_mail.attachments)

        if query.startswith("UUFA_EOT_SAP_AGGCP_LAW"):
            do_query(query, date + " SAP Aggregate Law Career.xls", directory,
                     prof_rk_mail.attachments)

        if query.startswith("UUFA_EOT_SAP_AGGCP_MED"):
            do_query(query, date + " SAP Aggregate Med Career.xls", directory,
                     prof_rk_mail.attachments)

        if query.startswith("UUFA_EOT_EU_FALL_GRADE"):
            do_query(query, date + " EU Grade Fall.xls", directory,
                     kklt_mail.attachments)

        if query.startswith("UUFA_EOT_EU_SPRING_GRADE"):
            do_query(query, date + " EU Grade Spring.xls", directory,
                     kklt_mail.attachments)

        if query.startswith("UUFA_EOT_EU_SUMMER_GRADE"):
            do_query(query, date + " EU Grade Summer.xls", directory,
                     kklt_mail.attachments)

        if query.startswith("UUFA_EOT_PELL_ELG_ENRLL_NO_AWD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_EOT_SAP_FSAP"):
            do_query(query, date + " FSAP Students.xls", directory,
                     kkt_mail.attachments)

        if query.startswith("UUFA_EOT_SCHOLAR_LEADER_CGPA"):
            do_query(query, date + " Scholarship CGPA Leadership.xls", directory,
                     jja_mail.attachments)

        if "EOT_WUE_ACAD_PROG_REV" in query:
            do_query(query, date + " WUE AcadProg Rvw 700007880038.xls", directory,
                     jja_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", "End of Term Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Disbursement Queries
def do_disb_queries():
    global aid_year
    year = "17"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB"):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
        
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    disb_date = str(raw_input("Enter Date the most recent Disbursement ran in 'MM-DD-YY' format:"))
    disb_date = disb_date[:-1]
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Disbursement',
                                                  aid_year, month_folder))
    else:
        directory = os.path.realpath(os.path.join('O:\Systems\QUERIES\Disbursement', aid_year, month_folder))

    if not os.path.isdir(directory):
        os.makedirs(directory)

    for query in os.listdir("."):
        if query.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB_") and (year in query[:-8]) :
            do_query(query, disb_date + " Item Types Authorized Not Disbursed " + year + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("UUFA_DQ_ALL_DISBURSED") and (year in query[:-8]) :
            do_query(query, disb_date + " All Item Types Disbursed " + year + ".xls", directory,
                     disb_mail.attachments)
                     
        if query.startswith("FA_DQ_ATHLETE_RM_BD_") and (year in query[:-8]) :
            do_query(query, disb_date + " Athlete Room and Board " + year + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("FA_DQ_ATH_OFF_SCHED_RM_BD_") and (year in query[:-8]) :
            do_query(query, disb_date + " Athlete Off Schedule R&B " + year + ".xls", directory,
                     disb_mail.attachments)

        if ("_CASH_DISB_TOTALS") in query and (year in query[:-8]) :
            do_query(query, disb_date + " Cash Disbursement Totals " + year + ".xls", directory,
                     disb_tot_mail.attachments)

        if ("_DQ_DISB_TOTALS") in query and (year in query[:-8]) :
            do_query(query, disb_date + " Disbursement Totals " + year + ".xls", directory,
                     disb_tot_mail.attachments)

        if query.startswith("UUFA_DQ_FALL_" + year) :
            do_query(query, disb_date + " DL Fall Awards 20" + year + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("UUFA_DQ_FALL_SPRING_") and (year in query[:-8]) :
            do_query(query, disb_date + " DL Fall Spring Awards 20" + year + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("UUFA_DQ_DL_SUMMER_") and (year in query[:-8]) :
            do_query(query, disb_date + " DL Summer Awards 20" + year + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("UUFA_DQ_SPRING_") and (year in query[:-8]) :
            do_query(query, disb_date + " DL Spring Awards 20" + year + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("UUFA_DQ_UG_PLUS_REFUND_IA_") and (year in query[:-8]) :
            do_query(query, disb_date + " DL UG PLUS Refund Borrower " + year + ".xls", directory,
                     sl_mail.attachments)

        if ("DQ_MISC_CARES_DISB") in query and (year in query[:-8]) :
            do_query(query, disb_date + " Misc CARES Disbursement Total " + year + ".xls", directory,
                     sl_mail.attachments)

        if ("DQ_MISC_CARES_RES_DISB") in query and (year in query[:-8]) :
            do_query(query, disb_date + " Misc CARES Resources Disb " + year + ".xls", directory,
                     sl_mail.attachments)

        if ("MISC_DISB_TOTALS") in query and (year in query[:-8]) :
            do_query(query, disb_date + " Misc Disbursement Totals " + year + ".xls", directory,
                     disb_tot_mail.attachments)

        if query.startswith("UUFA_DQ_MISC_RESOURCE_DISB_") and (year in query[:-8]) :
            do_query(query, disb_date + " Misc Resources Disbursement " + year + ".xls", directory,
                     disb_mail.attachments)

        if ("NONCASH_DISB_TOTALS") in query and (year in query[:-8]) :
            do_query(query, disb_date + " Non-cash Disbursement Totals " + year + ".xls", directory,
                     disb_tot_mail.attachments)

        if query.startswith("UUFA_DQ_PELL_ACPT_GR8_DISB_") and (year in query[:-8]) :
            do_query(query, disb_date + " Pell Accepted Awards Greater Disb " + year + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("UUFA_DQ_SF_ITEM_TYPE_ERROR"):
            do_query(query, disb_date + " FA SF Item Type Error " + year + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("UUFA_DQ_TEACH_GRANT_" + str(int(year) - 1)):
            do_query(query, disb_date + " Teach Grant Recipients 20" + str(int(year) - 1) + ".xls", directory,
                     disb_mail.attachments)

        if query.startswith("UUFA_DQ_TEACH_GRANT_") and (year in query[:-8]) :
            do_query(query, disb_date + " Teach Grant Recipients 20" + year + ".xls", directory,
                     disb_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Disbursement Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#2nd LDR Queries
def do_2nd_ldr():
    global aid_year
    year = "18"
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
                     kkt_mail.attachments)

        if query.startswith("UUFA_HRS_DECREASE_ATH"):
            do_query(query, date + " Hours Decrease Athlete.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_HRS_DECREASE_FC"):
            do_query(query, date + " Hours Decrease FC.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_HRS_DECREASE_SV"):
            do_query(query, date + " Hours Decrease SV.xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_PELL_ELIG_ENROLL_NO_AWARD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PELL_OFFERED_NOT_DISB"):
            do_query(query, date + " Pell Offered Not Disbursed.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PELL_DL_MATH990"):
            do_query(query, date + " Pell-DL Math 990.xls", directory,
                     kkmt_mail.attachments)

        if "PELL_DL_MATH980" in query:
            do_query(query, date + " Pell-DL Math 980.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PELL_DL_ELI575_ELI685"):
            do_query(query, date + " Pell-DL ELI0075 ELI0085.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_SF_DISB_ATH_AWD_NOPOST"):
            do_query(query, date + " Athlete Awd Disb not Posted.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_SF_DISB_WAIVER_AWD_NOPOST"):
            do_query(query, date + " Waiver Awd Disb not Posted.xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_PRT_PELL_ELG_NO_PELL_") and (year in query[:-8]) :
            do_query(query, date + " Pell Eligible No Pell 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Second Session LDR Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Day After LDR Queries
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

    # Change File_Name to be file as it is received and _new_file_name to what
    # the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_LDR_MIN_ENROLLMENT_ATH"):
            do_query(query, date + " Minimum Enrollment Athlete.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_LDR_MINIMUM_ENROLLMENT_FC"):
            do_query(query, date + " Minimum Enrollment FC (Federal & Campus Based Aid).xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_LDR_MINIMUM_ENROLLMENT_SV"):
            do_query(query, date + " Minimum Enrollment SV (Scholarships & Waivers).xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_LDR_PELL_AWARDS"):
            do_query(query, date + " Pell Awards.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_ATHLETE_AWARD_DISBURSED"):
            do_query(query, date + " Athlete Award Disbursed.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_CBA_UNDISBURSED"):
            do_query(query, date + " CBA Undisbursed.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PELL_DL_MATH990"):
            do_query(query, date + " Pell-DL MATH 990.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PELL_DL_ELI575_ELI685"):
            do_query(query, date + " Pell-DL ELI575 ELI685.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PELL_ELIG_ENROLL_NO_AWARD"):
            do_query(query, date + " Pell Eligible Enrolled NO Award.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PELL_OFFERED_NOT_DISB"):
            do_query(query, date + " Pell Offered Not Disbursed.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PELL_SUMMER_ENROLLMENT"):
            do_query(query, date + " Pell Summer Enrollment Check.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_THESIS_STUDENTS_NONRES"):
            do_query(query, date + " Thesis Students Non-Resident.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_REGISTERED_CENSUS_DATE"):
            do_query(query, date + " LDR FA Load Check.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_SF_DISB_ATH_AWD_NOPOST"):
            do_query(query, date + " Athlete Awd Disb not Posted.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_SF_DISB_WAIVER_AWD_NOPOST"):
            do_query(query, date + " Waiver Awd Disb Not Posted.xls", directory,
                     hj_mail.attachments)

        if query.startswith("UUFA_PRT_AWD_PLL_ELG_NO_PLL_") and (year in query[:-8]) :
            do_query(query, date + " Pell Eligible No Pell 20" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_NO_MTRC_STU_ATH_BAL_OWING"):
            do_query(query, date + " Non-Matric Stu Athlete Balance Owing.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_FATERM_SOURCE_N_AWD_ATH"):
            do_query(query, date + " NonTerm Source N Cancel Award Rvw - Athlete.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_FATERM_SOURCE_N_AWD_SV"):
            do_query(query, date + " Term Source N Cancel Award Rvw - Scholarships.xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_PRT_PELL_ELG_NO_PELL_"):
            do_query(query, date + " Pell Eligible No Pell " + year + ".xls", directory,
                     ath_mail.attachments)

        if query.startswith("UUFA_ATH_NO_MTRC_STU_BAL_OWING"):
            do_query(query, date + " Non-Matric Stu Athlete Balance Owing " + year + ".xls", directory,
                     ath_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Day After LDR Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Direct Loan Pre-Outbound Queries
def dl_pre_outbound():
    global aid_year
    year = '17'
    for query_name in os.listdir("."):
        if "DLR_LOAN_ORIG_EDIT_ERR" in query_name:
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year
    orig_file_doc = date + " DL ORIG 20" + year + ".doc"
    orig_file_doc_2 = date + " DL ORIG 20" + year + " (2).doc"
    orig_file_docx = date + " DL ORIG 20" + year + ".docx"
    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Direct Loans', 'DL Pre-Outbound'))
        orig_doc = os.path.realpath(os.path.join('C:\Testing Bob\Direct Loans', 'Origination', orig_file_doc))
        orig_docx = os.path.realpath(os.path.join('C:\Testing Bob\Direct Loans', 'Origination', orig_file_docx))

    else:
        directory = os.path.realpath(os.path.join('O:/Systems/Direct Loans', 'DL Pre-Outbound'))
        orig_doc = os.path.realpath(os.path.join('O:/Systems/Direct Loans', 'Origination', orig_file_doc))
        orig_doc_2 = os.path.realpath(os.path.join('O:/Systems/Direct Loans', 'Origination', orig_file_doc_2))
        orig_docx = os.path.realpath(os.path.join('O:/Systems/Direct Loans', 'Origination', orig_file_docx))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what
    # the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_DLR_ORIG_TRNS_PEND_") and (year in query[:-8]) :
            do_query(query, date + " DL Orig-Trans Pending " + year + ".xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_ENTRANCE_COUNSEL_") and (year in query[:-8]) :
            do_query(query, date + " DL Entrance Counseling I  " + year + ".xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_LOAN_EFT_DT_LNDR_ERR"):
            do_query(query, date + " Loan EFT Date Lender Error.xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_LOAN_ORIG_FA_LOAD_"):
            do_query(query, date + " DL ORIG FA Load " + year + ".xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_LOAN_NO_NSLDS_") and (year in query[:-8]) :
            do_query(query, date + " Loan No NSLDS " + year + ".xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_LOAN_ORIG_ACAD_LVL_") and (year in query[:-8]) :
            do_query(query, date + " Loans Academic Level " + year + ".xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_LOAN_ORIG_EDIT_ERR"):
            do_query(query, date + " Loan Originate Edit Errors.xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_LOAN_ORIG_SPLT_CDS_") and (year in query[:-8]) :
            do_query(query, date + " Loan Split Codes " + year + ".xls", directory,
                     dl_mail.attachments)

        if "LN_ORIG_VLOAN_RSN" in query and (year in query[:-8]) :
            do_query(query, date + " Loan ORIG VLOAN Reasons.xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_LOAN_SPC_NEED_OVWD_") and (year in query[:-8]) :
            do_query(query, date + " Loan Overaward Special Need " + year + ".xls", directory,
                     dl_mail.attachments)

        if "DLR_LN_ACPT_STAF_31_32_" in query and year in query[:-8] :
            do_query(query, date + " Stafford Accept Offer " + year + ".xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_NOT_DISB_") and (year in query[:-8]) :
            do_query(query, date + " DL Disb Errors " + year + ".xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_NOT_DISBURSED_") and (year in query[:-8]) :
            do_query(query, date + " DL Disbursement Errors " + year + "(Old).xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_NOT_DISBURSED_20") and (year in query[:-8]) :
            do_query(query, date + " DL Disbursement Errors " + year + ".xls", directory,
                     dl_mail.attachments)

        if query.startswith("UUFA_DLR_UG_PLUS_REFND_IND"):
            do_query(query, date + " DL UG PLUS Refund Indicator.xls", directory,
                     dl_mail.attachments)

    if not test:
        while True:
            if os.path.isfile(orig_doc):
                dl_mail.attachments.append(orig_doc)
                if os.path.isfile(orig_doc_2):
                    dl_mail.attachments.append(orig_doc_2)
                break
            if os.path.isfile(orig_docx):
                dl_mail.attachments.append(orig_docx)
                if os.path.isfile(orig_doc_2):
                    dl_mail.attachments.append(orig_doc_2)
                break
            else:
                raw_input("\nCould not locate DL ORIG 20" + year + ".doc\nMake sure it is located in O:/Systems/Direct Loans/Origination\n\nPress Enter when ready.")

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Pre-Outbound Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Alt Loan Pre-Outbound Queries
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

    # Change File_Name to be file as it is received and _new_file_name to what
    # the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_ALR_110_CHNG_PDG_TRANS_"):
            do_query(query, date + " Loan Pending Change Transactions.xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_CL_APP_RSPNS_ERR_"):
            do_query(query, date + " CL Response Load Errors.xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LN_FA907_1_REVISE_"):
            do_query(query, date + " Loan Disbursed Report " + year + ".xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LN_FA907_2_REVISE_"):
            do_query(query, date + " CLoan Not Disbursed Report " + year + ".xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LN_SENT_NO_RESP_"):
            do_query(query, date + " CLoan Sent No Response " + year + ".xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LN_EFT_DETAIL_ERR_"):
            do_query(query, date + " Loans EFT Detail Error.xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LN_EFT_DT_LNDR_ERR_"):
            do_query(query, date + " Loan EFT Date Lender Errors.xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LOAN_ORIG_ACAD_LVL_") and (year in query[:-8]) :
            do_query(query, date + " Loans Academic Level " + year + ".xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LN_ORIG_EDIT_ERR_"):
            do_query(query, date + " Loan Originate Edit Errors.xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LOAN_ORIG_FA_LOAD_") and (year in query[:-8]) :
            do_query(query, date + " Loan ORIG FA Load " + year + ".xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LOAN_ORG_LND_NT_CK_") and (year in query[:-8]) :
            do_query(query, date + " Loan ORIG Lender Note " + year + ".xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LOAN_ORIG_SPLT_CDS_") and (year in query[:-8]) :
            do_query(query, date + " Loan ORIG Split Codes " + year + ".xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LN_ORIG_VLOAN_RSN_"):
            do_query(query, date + " Loan ORIG VLOAN Reasons.xls", directory,
                     alt_mail.attachments)

        if query.startswith("UUFA_ALR_LOAN_SPC_NEED_OVWD_") and (year in query[:-8]) :
            do_query(query, date + " Loan Overaward Special Need " + year + ".xls", directory,
                     alt_mail.attachments)

    if not test:
        while True:
            if os.path.isfile(orig_doc):
                alt_mail.attachments.append(orig_doc)
                break
            if os.path.isfile(orig_docx):
                alt_mail.attachments.append(orig_docx)
                break
            if str.capitalize(skip) == "Y":
                break
            else:
                skip = raw_input("\nCould not locate ALT Loan ORIG 20" + year + ".doc\nMake sure it is located in O:/Systems/Queries/ALT Loans/" + aid_year + "\n\nPress Enter when ready.")

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Alt Loan Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Pre-Repackaging Queries
def do_pre_repackaging():
    global aid_year
    strm = "no STRM found"
    prompt = "What STRM is this for? (ex. 1154 = Spring 2015:"
    aid_year = "not defined"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_PP_"):
            while True:
                strm = raw_input(prompt)
                if (len(strm) == 5) and ((0 / int(strm) == 0)):
                    break
                else:
                    print ("\nOnly STRMs in the form of 1(year)4/6/8 are acceptable. Ex: 1168 is Fall of 16.")
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
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PP_RPKG_AWD_AY_NO_BDGT"):
            do_query(query, date + " Pell Award  AY One STRM Budget.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PP_RPKG_AWD_STRM_INACTIVE"):
            do_query(query, date + " Pell Award STRM Inactive.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PP_RPKG_AWRD_LOCK"):
            do_query(query, date + " Pell Award Lock No FPEL.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PP_RPKG_COA_DOUBLE"):
            do_query(query, date + " Pell COA Double.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PP_RPKG_LTHT_PELL_COA"):
            do_query(query, date + " Pell COA LTHT.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PP_RPKG_RPKG_NO_BUDGET"):
            do_query(query, date + " Pell Repackaging No Budget.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_PP_RPKG_RPKG_SNAPSHOT"):
            do_query(query, date + " Pell Repackage Snapshot.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_PP_RPKG_TOTAL_WDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop.xls", directory,
                     kkmt_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Pre-Pell Only Repackaging Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Mid-Repackaging Queries
def do_mid_repack_queries():
    global aid_year
    strm = "no STRM found"
    prompt = "What STRM is this for? (EX. 1154 = Spring 2015:)"
    aid_year = "not defined"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_MP_"):
            while True:
                strm = raw_input(prompt)
                if len(strm) == 5 and (0 / int(strm) == 0):
                    break
                else:
                    print ("\nOnly STRMs in the form of 1(year)4/6/8 are acceptable.")
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
                     kkmt_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_AWD_STRM_INACTIVE"):
            do_query(query, date + " Pell Award STRM Inactive.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_FCIT"):
            do_query(query, date + " Pell Repackage FCIT" + year + " FDEG" + year + ".xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_FDR"):
            do_query(query, date + " Pell Repackage FDR" + year + " Initiated.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_FOVP_FARC_I"):
            do_query(query, date + " Pell REPKG FOVP FACR Initiated.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_ISIR_CMT_346_347"):
            do_query(query, date + " Pell Repackage ISIR CMT 346 347.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_LEAVE"):
            do_query(query, date + " Pell Repackaging Leave Absense.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_TOTAL_WDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop.xls", directory,
                     sys_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_TOTAL_WTHDRN_DRP"):
            do_query(query, date + " Pell Total Withdrawn Drop (old).xls", directory,
                     sys_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_VAR_1_2"):
            do_query(query, date + " Pell Repackage Flags.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_MP_RPKG_VER_UNFLAG"):
            do_query(query, date + " Pell Verification Flag Unchecked.xls", directory,
                     null_mail.attachments)

 
    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Pell Repackaging Mid-Packaging Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#After Repackaging Queries
def do_after_repackaging():
    global aid_year
    strm = "no STRM found"
    prompt = "What STRM is this for? (ex. 1154 = Spring 2015:"
    aid_year = "not defined"
    for query_name in os.listdir("."):
        if query_name.startswith("UUFA_AP_"):
            while True:
                strm = raw_input(prompt)
                if len(strm) == 5 and (0 / int(strm) == 0):
                    break
                else:
                    print ("\nOnly STRMs in the form of 1(year)4/6/8 are acceptable.")
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
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_ACTN"):
            do_query(query, date + " Pell Award Activity.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_AWACT_C"):
            do_query(query, date + " Pell Awards Cancelled.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_AW_ACT"):
            do_query(query, date + " Pell Repackage Activity.xls", directory,
                     v_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_FPEL_AWARD_LCK"):
            do_query(query, date + " Pell Award Lock FPEL" + year + ".xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_PLAN_ID_BLANK"):
            do_query(query, date + " Pell Repackaging Plan ID Blank.xls", directory,
                     v_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_RPKG_SNAPSHOT"):
            do_query(query, date + " Pell Repackage Snapshot.xls", directory,
                     null_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_SAP_HOLD_DELETED"):
            do_query(query, date + " Pell SAP Holds.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_SKIP"):
            do_query(query, date + " Pell Repackage Skip.xls", directory,
                     ml_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_TERM_FT"):
            do_query(query, date + " Term Pell Awards FT.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_TERM_HT"):
            do_query(query, date + " Term Pell Awards HT.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_TERM_LH"):
            do_query(query, date + " Term Pell Awards LH.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_TERM_NL"):
            do_query(query, date + " Term Pell Awards NL.xls", directory,
                     kkmt_mail.attachments)

        if query.startswith("UUFA_AP_RPKG_TERM_TQ"):
            do_query(query, date + " Term Pell Awards TQ.xls", directory,
                     kkmt_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Pell Only Repackaging Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Daily Scholarships Queries
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

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what
    # the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if query.startswith("UUFA_SCHOLAR_DISB_ZERO_") and (year in query[:-8]) :
            do_query(query, date + " Scholarships awarded not disbursed " + year + ".xls", directory,
                     ss_mail.attachments)

        if query.startswith("UUFA_SCHOLAR_TWO_CAREERS_") and (year in query[:-8]) :
            do_query(query, date + " Scholarship Award with Two Careers " + year + ".xls", directory,
                     maisa_mail.attachments)

        if query.startswith("UUFA_SCHOLAR_AUTH_NOT_DISB_") and (year in query[:-8]) :
            do_query(query, date + " Scholar Authorized Not Disbursed " + year + ".xls", directory,
                     maisa_mail.attachments)

 
    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Daily Scholarship Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#Weekly Scholarships Queries
def do_weekly_scholarships():
    global aid_year
    global year
    for query_name in os.listdir("."):
        if ("UUFA_WS_" in query_name):
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year

    if test:
        directory = os.path.realpath(os.path.join('C:\Testing Bob/Scholarships', aid_year + ' Scholar\Queries'))
        directory_save = os.path.realpath(os.path.join('C:\Testing Bob/Save', aid_year))
        directory_errors = os.path.realpath('C:\Systems\External Awards\Errors')
    else:
        directory = os.path.realpath(os.path.join('O:\Systems\Scholarships', aid_year + ' Scholar\Queries'))
        directory_save = os.path.realpath(os.path.join('O:\Systems\Queries\Save', aid_year))
        directory_errors = os.path.realpath('O:\Systems\External Awards\Errors')

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(directory_save):
        os.makedirs(directory_save)
    if not os.path.isdir(directory_errors):
        os.makedirs(directory_errors)

    # Change File_Name to be file as it is received and _new_file_name to what
    # the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if ("DEPT_POST_WRNG_ITEM" in query) and (year in query[:-8]) :
            do_query(query, date + " Depts posting to the wrong item type  " + year + ".xls", directory,
                     bjns_mail.attachments)

        if ("MISC_TOTAL" in query) and (year in query[:-8]) :
            do_query(query, date + " IT Dept Misc Awards Total (10001) " + year + ".xls", directory_save,
                     null_mail.attachments)

        if ("CARES_HEERF" in query):
            do_query(query, date + " Cares & HEERF Awards by Term.xls", directory,
                     schol_mail.attachments)

        if ("MISC_NOPOST" in query) and (year in query[:-8]) :
            do_query(query, date + " 7880013 Dept Misc Awards No Post " + year + ".xls", directory_save,
                     hjr_mail.attachments)

        if ("BOOKS_TOTAL" in query) and (year in query[:-8]) :
            do_query(query, date + " 788 IT Dept Books Awards Total (10002) " + year + ".xls", directory_save,
                     null_mail.attachments)

        if ("BOOKS_NOPOST" in query) and (year in query[:-8]) :
            do_query(query, date + " 7880015 Dept Books No Post " + year + ".xls", directory_save,
                     hjr_mail.attachments)

        if ("ROOMBOARD_TOTAL" in query) and (year in query[:-8]) :
            do_query(query, date + " 788 IT Dept Room & Board Awards Total (10003) " + year + ".xls", directory_save,
                     null_mail.attachments)

        if ("ROOMBOARD_NOPOST" in query) and (year in query[:-8]) :
            do_query(query, date + " 7880029 Dept Room & Board No Post " + year + ".xls", directory_save,
                     hjr_mail.attachments)

        if ("TRAVEL_TOTAL" in query) and (year in query[:-8]) :
            do_query(query, date + " 788 IT Dept Travel Awards Total (10004) " + year + ".xls", directory_save,
                     null_mail.attachments)

        if ("TRAVEL_NOPOST" in query) and (year in query[:-8]) :
            do_query(query, date + " 7880033 Dept Travel No Post " + year + ".xls", directory_save,
                     hjr_mail.attachments)

        if (("TRAINEESHIP" in query) and ("TOTAL" in query)) and (year in query[:-8]) :
            do_query(query, date + " 788 IT Dept Traineeship Awards Total (10005) " + year + ".xls", directory_save,
                     null_mail.attachments)

        if ("TRAINEESHIP_NOPOST" in query) and (year in query[:-8]) :
            do_query(query, date + " 7880034 Dept Traineeship No Post " + year + ".xls", directory_save,
                     hjr_mail.attachments)

        #if( year == int(date[-2:])+1):
        if ("SCH_ALL_NEED" in query) and (year in query[:-8]):
            do_query(query, date + " All Scholarships Need Based " + year + ".xls", directory,
                     bhjjns_mail.attachments)

        if ("SCH_ALL_NRFRESH" in query) and (year in query[:-8]):
            do_query(query, date + " All Scholarships Non Res Freshman " + year + ".xls", directory,
                     bhjjns_mail.attachments)

        if ("SCH_ALL_NRTRAN" in query) and (year in query[:-8]):
            do_query(query, date + " All Scholarships Non Res Transfer " + year + ".xls", directory,
                     bhjjns_mail.attachments)

        if ("SCH_ALL_RESFRESH" in query) and (year in query[:-8]):
            do_query(query, date + " All Scholarships Res Freshman " + year + ".xls", directory,
                     bhjjns_mail.attachments)

        if ("SCH_ALL_RESTRAN" in query) and (year in query[:-8]):
            do_query(query, date + " All Scholarships Res Transfer " + year + ".xls", directory,
                     bhjjns_mail.attachments)

        if ("WS_UT_PROMISE_CHKLST" in query) and (year in query[:-8]):
                do_query(query, date + " Utah Promise Checklist " + year + ".xls", directory,
                     jj_mail.attachments)

        if ("WS_SCH_ALL_EA_ERRORS" in query):
                do_query(query, date + " Weekly Error Query " + year + ".xls", directory_errors,
                     schol_mail.attachments)

        if ("WS_SCHOLAR_ADM_REVIEW" in query) and (year in query[:-8]):
                do_query(query, date + " Scholarship Admin Review " + year + ".xls", directory,
                     schol_mail.attachments)

        if ("WS_FAC_STAFF" in query) and (year in query[:-8]):
                do_query(query, date + " Faculty-Staff Not Posted " + year + ".xls", directory,
                     schol_mail.attachments)
            
        if ("WS_SCHOLAR_STATUS" in query):
                do_query(query, date + " Scholarship Offer Status DEIN " + year + ".xls", directory,
                     schol2_mail.attachments)


    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", aid_year + " Weekly Scholarships Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#ATB, FBILL, 3C Queries
def do_atb_fb_3c_queries():
    global aid_year
    year = "19"
    aid_year = "20" + str(int(year) - 1) + "-20" + year

    if test:
        directory = os.path.realpath(os.path.join('C:/Testing Bob/QUERIES/3C Queries'))
        atb_directory = os.path.realpath(os.path.join('C:/Testing Bob/QUERIES\ATB'))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/3C Queries'))
        atb_directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/ATB'))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)
    if not os.path.isdir(atb_directory):
        os.makedirs(atb_directory)

    # Change File_Name to be file as it is received and _new_file_name to what
    # the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if ("UUFA_ADD_FDLP" in query) and (year in query[:-8]) :
            do_query(query, date + " ADD FDLP" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_ADD_FGLO" in query) and (year in query[:-8]) :
            do_query(query, date + " ADD FGLO" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_ADD_FHST" in query) and (year in query[:-8]) :
            do_query(query, date + " ADD FHST" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_ADD_FMPN" in query) and (year in query[:-8]) :
            do_query(query, date + " ADD FMPN" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_ADD_FNON" in query) and (year in query[:-8]) :
            do_query(query, date + " ADD FNON" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_ADD_FTYN" in query) and (year in query[:-8]) :
            do_query(query, date + " ADD FTYN" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_ADD_FULO" in query) and (year in query[:-8]) :
            do_query(query, date + " ADD FULO" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_ADD_FLPR" in query) and (year in query[:-8]) :
            do_query(query, date + " ADD FLPR" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_COMPLETE_FDLP" in query) and (year in query[:-8]) :
            do_query(query, date + " COMPLETE FDLP" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_COMPLETE_FGLO" in query) and (year in query[:-8]) :
            do_query(query, date + " COMPLETE FGLO" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_COMPLETE_FHST" in query) and (year in query[:-8]) :
            do_query(query, date + " COMPLETE FHST" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_COMPLETE_FMPN" in query) and (year in query[:-8]) :
            do_query(query, date + " COMPLETE FMPN" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_COMPLETE_FNON" in query) and (year in query[:-8]) :
            do_query(query, date + " COMPLETE FNON" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_COMPLETE_FULO" in query) and (year in query[:-8]) :
            do_query(query, date + " COMPLETE FULO" + year + ".xls", directory,
                     null_mail.attachments)

        if ("UUFA_HS_06_AFTER_SEC" in query):
            do_query(query, date + " SEC High School Code 06.xls", atb_directory,
                     null_mail.attachments)

        if ("UUFA_HS_04_AFTER_SEC" in query):
            do_query(query, date + " SEC Home School Code 04.xls", atb_directory,
                     null_mail.attachments)

        if ("UUFA_GED_07_AFTER_SEC" in query):
            do_query(query, date + " SEC New ISIRs GED 07.xls", atb_directory,
                     null_mail.attachments)

        if ("UUFA_ATB_ISIR_NOT_MATCH" in query):
            do_query(query, date + " ISIR Not Match SEC.xls", atb_directory,
                     kkmt_mail.attachments)

        if ("UUFA_ATB_SEQUENCE_DIFFERENCE" in query):
            do_query(query, date + " SEC Sequence Difference.xls", atb_directory,
                     kkmt_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", " ATB Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

#TSM Queries            
def do_tsm_queries():
    global aid_year
    year = "19"
    for query_name in os.listdir("."):
        if "_NSLDS_" in query_name:
            year = str(int(re.search(r'\d+', query_name).group()))
            break
    aid_year = "20" + str(int(year) - 1) + "-20" + year

    if test:
        directory = os.path.realpath(os.path.join('C:/QUERIES/TSM/NSLDS TSM Request'))
    else:
        directory = os.path.realpath(os.path.join('O:/Systems/QUERIES/TSM/NSLDS TSM Request'))

    # the list 'my_path' should be populated with the FOLDER variables above.
    if not os.path.isdir(directory):
        os.makedirs(directory)

    # Change File_Name to be file as it is received and _new_file_name to what
    # the new file should be.  Prefix date
    # will be added.
    for query in os.listdir("."):
        if ("NSLDS_REQUEST_DN" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Request DN " + year + ".xls", directory,
                     null_mail.attachments)

        if ("NSLDS_REQUEST_GR" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Request DN " + year + ".xls", directory,
                     null_mail.attachments)

        if ("NSLDS_REQUEST_LW" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Request DN " + year + ".xls", directory,
                     null_mail.attachments)

        if ("NSLDS_REQUEST_MD" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Request DN " + year + ".xls", directory,
                     null_mail.attachments)

        if ("NSLDS_REQUEST_UG" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Request DN " + year + ".xls", directory,
                     null_mail.attachments)

        if ("_NSLDS_VAR_FLAG9_DN" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Var Flag9 DN " + year + ".xls", directory,
                     null_mail.attachments)

        if ("_NSLDS_VAR_FLAG9_GR" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Var Flag9 GR " + year + ".xls", directory,
                     null_mail.attachments)

        if ("_NSLDS_VAR_FLAG9_LW" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Var Flag9 LW " + year + ".xls", directory,
                     null_mail.attachments)

        if ("_NSLDS_VAR_FLAG9_MD" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Var Flag9 MD " + year + ".xls", directory,
                     null_mail.attachments)

        if ("_NSLDS_VAR_FLAG9_UG" in query) and (year in query[:-8]) :
            do_query(query, date + " NSLDS Var Flag9 UG " + year + ".xls", directory,
                     null_mail.attachments)

    for mail_group in query_groups:
        if(mail_group.attachments):
            mailer("", " TSM Queries", mail_group.recipients,"", mail_group.attachments)
            del mail_group.attachments[:]

def main():
    for filename in os.listdir("."):
    # Daily Queries
        if filename.startswith("UUFA_IL_ISIR_SSN_MISMATCH"):
            do_dailies()
    # Monday Weekly Queries
        if filename.startswith("UUFA_WR_"):
            do_monday_weeklies()
    # Budget Queries
        if "UUFA_BR_" in filename:
            do_budget_queries()
    # Packaging Queries
        if filename.startswith("UUFA_PRT_ACAD_PROG_REVIEW"):
            do_packaging_queries()
    # Monthly Queries
        if "MR_PELL_SSN_MISMATCH" in filename:
            do_monthlies()
    # Disbursement Queries
        if filename.startswith("UUFA_DQ_AUTHORIZED_NOT_DISB"):
            do_disb_queries()
    #2nd LDR Queries
        if filename.startswith("UUFA_PRT_PELL_ELG_NO_PELL"):
            do_2nd_ldr()
    # End of Term Queries
        if filename.startswith("UUFA_EOT_ACAD_PLAN_RVW"):
            do_end_of_term_queries()
    # Day After LDR Queries
        if filename.startswith("UUFA_LDR_MIN_ENROLLMENT_ATH"):
            do_day_after_ldr()
    # Direct Loans Pre-Outbound Queries
        if "DLR_LOAN_ORIG_EDIT_ERR" in filename:
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
        if ("UUFA_WS" in filename):
            do_weekly_scholarships()
    # Budget Testing Queries
        if filename.startswith("UUFA_BUDGET_20"):
            do_budget_test_queries()
    # ATB and 3C Queries
        if "UUFA_ATB" in filename:
            do_atb_fb_3c_queries()
    # Remove extra files 
        if "FASTDVER" in filename or "FINAID_Checklist_" in filename  or "ussfa09" in filename or "USSFA090 Reset" in filename or "O-A" in filename:
            os.remove(filename)
            print("Removed " + filename)
    # Transfer Student Monitoring
        if "_NSLDS_" in filename:
            do_tsm_queries()

if __name__ == "__main__":
    main()


        # TEMPLATE
        # Change File_Name to be file as it is received and _new_file_name to what the new file should be.  Prefix date
        # will be added.
        # for query in os.listdir("."):
        # if query.startswith("____________________"):
        #        do_query(query, date + " ________________" + year + ".xls", directory,
        #                 lkj_mail.attachments)

        #for mail_group in mail_groups:
        #if(mail_group.attachments):
        #    mailer("", aid_year + " Queries", mail_group.recipients,"", mail_group.attachments)
        #    del mail_group.attachments[:]

raw_input("So Long, and Thanks for All the Fish.\nPRESS ENTER TO CLOSE.")