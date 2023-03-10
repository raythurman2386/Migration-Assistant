# UPDATE EACH VARIABLE NAME WITH THE PROPER COLUMN FROM THE SPREADSHEET
# MAIN.PY IS WHERE CLIENT FILE IS CONVERTED TO OUR TEMPLATE
# DO NOT CHANGE MAIN.PY FILE, ALL CHANGES RUN THROUGH THIS FILE
# ENSURE STARTING ROWS ARE SET AND BEGIN YOUR COUNT FOR EACH VARIABLE
# FROM THAT STARTING ROW, INDEXING STARTS AT 0
# THESE VARIABLES COVER THE CASE AND CLASS TABS, MINUS OUTCOMES
# ONCE WE HAVE REGULAR AGENCIES SENDING OUTCOMES THAT WILL BE BUILT AS WLL
# MINOR UPDATES TO DATA CORRECTION OR UTILS FILE MAY BE NECESSARY

# AGENCY SHOULD JUST BE THE AGENCY NAME TO SEND THE PROPER VARIABLES FOR THE TEMPLATES
# AGENCY IS FOR THE OUTGOING FILE NAME
# MULTIPLE FILE VARIABLES JUST IN CASE THERE ARE MULTIPLE FILES FROM AGENCY
CASE_VARS = {
    # ALL VARIABLES FOR GRABBING CASE DATA, SOME FIELDS GO TO CASE AS WELL
    # STARTING ROWS AND COLUMNS ARE USED INSIDE OF EACH FOR LOOP, ENSURE THESE ARE CORRECT
    # STARTING ROW AND COLUMN ARE 1 INDEX BASED!!!!!
    "STARTINGROW": 2,
    "STARTINGCOL": 1,
    # ALL VARIABLES SHOULD BE DOUBLE CHECKED BEFORE RUNNING TO ENSURE SCRIPT
    # WILL IDEALLY ONLY NEED TO BE RUN ONCE, IF AGENCY STICKS TO OUR TEMPLATE
    # ALL UPDATES SHOULD GO SMOOTHLY AND EFFICIENTLY
    # ALL OF THESE FIELDS ARE 0 INDEX BASED!!!!!!!
    "CLIENTID":0,
    "FIRSTNAME":1,
    "LASTNAME":2,
    "PHONE":3,
    "ADDRESSNUMBER":4,
    "STREETNAME":5,
    "CITY":6,
    "STATE":7,
    "COUNTY":8,
    "ZIP":9,
    "RACE":10,
    "ETHNICITY":11,
    "ENGLISHPROF":12,
    "HOUSEHOLDINCOME":13,
    "HOUSEHOLDSIZE":14,
    "APTNUMBER":15,
    "EMAIL":16,
    "HOUSEHOLDTYPE":48,
    "RURALSTATUS":49,
    "INITIALCASETYPE":17,
    "DATEOFBIRTH":18,
    "COUNSELOR":19,
    "INTAKEDATE":20,
    "CASETYPE":21,
    "CASESTARTDATE":22,
    "CASEID":23,
    "SESSIONID":24,
    "DATE":25,
    "NOFAGRANT":26,
    "NOTES":47,
}


# if the agency is maha return their class variable
MAHA_VARS = {
    # MAHA DO NOT CHANGE, UNCOMMENT FOR WHICH AGENCY IS BEING MIGRATED
    # STARTING ROW AND COLUMN ARE 1 INDEX BASED!!!!!
    "STARTINGROW":2,
    "STARTINGCOL":1,
    # ALL OF THESE FIELDS ARE 0 INDEX BASED!!!!!!!
    "CLIENTID":3,
    "FIRSTNAME":1,
    "LASTNAME":2,
    "PHONE":25,
    "ADDRESSNUMBER":5,
    "STREETNAME":6,
    "CITY":8,
    "STATE":11,
    "COUNTY":10,
    "ZIP":9,
    "RACE":20,
    "ETHNICITY":19,
    "ENGLISHPROF":21,
    "HOUSEHOLDINCOME":18,
    "HOUSEHOLDSIZE":24,
    "APTNUMBER":7,
    "EMAIL":4,
    "INITIALCASETYPE":22,
    "DATEOFBIRTH":15,
    "COUNSELOR":26,
    "INTAKEDATE":15,
    "RURALSTATUS":23,
    "HOUSEHOLDTYPE":17,
    "COURSEID":13,
    "CLASSDATE":16,
    "ATTENDANCESTATUS":14,
}

# if the agency is abcdc return their class variables
ABCDC_VARS = {
    # ABCDC DO NOT CHANGE, UNCOMMENT FOR WHICH AGENCY IS BEING MIGRATED
    # STARTING ROW AND COLUMN ARE 1 INDEX BASED!!!!!
    "STARTINGROW":2,
    "STARTINGCOL":1,
    # ALL OF THESE FIELDS ARE 0 INDEX BASED!!!!!!!
    "CLIENTID":0,
    "FIRSTNAME":1,
    "LASTNAME":2,
    "PHONE":3,
    "ADDRESSNUMBER":4,
    "STREETNAME":5,
    "CITY":6,
    "STATE":7,
    "COUNTY":8,
    "ZIP":9,
    "RACE":10,
    "ETHNICITY":11,
    "ENGLISHPROF":12,
    "HOUSEHOLDINCOME":13,
    "HOUSEHOLDSIZE":14,
    "APTNUMBER":15,
    "EMAIL":16,
    "INITIALCASETYPE":19,
    "DATEOFBIRTH":20,
    "COUNSELOR":21,
    "INTAKEDATE":22,
    "RURALSTATUS":18,
    "HOUSEHOLDTYPE":17,
    "COURSEID":23,
    "CLASSDATE":24,
    "ATTENDANCESTATUS":25,
}


# THESE SHOULD BE THE ONLY ADDITIONAL VARIABLES THAT ARE NEEDED FOR THE CLASS
# PORTION OF OUR TEMPLATE
# THESE ARE DUPLICATES ONLY BECAUSE WE GET TWO FILES OCCASIONALLY
CLASS_VARS = {
    # BASE TEMPLATE and FEA DO NOT CHANGE, UNCOMMENT FOR WHICH AGENCY IS BEING MIGRATED
    # STARTING ROW AND COLUMN ARE 1 INDEX BASED!!!!!
    "STARTINGROW":2,
    "STARTINGCOL":1,
    # ALL OF THESE FIELDS ARE 0 INDEX BASED!!!!!!!
    "CLIENTID":0,
    "FIRSTNAME":1,
    "LASTNAME":2,
    "PHONE":3,
    "ADDRESSNUMBER":4,
    "STREETNAME":5,
    "CITY":6,
    "STATE":7,
    "COUNTY":8,
    "ZIP":9,
    "RACE":10,
    "ETHNICITY":11,
    "ENGLISHPROF":12,
    "HOUSEHOLDINCOME":13,
    "HOUSEHOLDSIZE":14,
    "APTNUMBER":15,
    "EMAIL":16,
    "INITIALCASETYPE":17,
    "DATEOFBIRTH":18,
    "COUNSELOR":19,
    "INTAKEDATE":20,
    "RURALSTATUS":25,
    "HOUSEHOLDTYPE":24,
    "COURSEID":21,
    "CLASSDATE":22,
    "ATTENDANCESTATUS":23,
}

