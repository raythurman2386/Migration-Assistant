from utils import options
from uszipcode import SearchEngine
import datetime


# Shorten Legacy ID
def correct_legacy_id(id):
    return str(id)[:10].replace('W', '')


# Correct Name
def correct_name(name):
    return ''.join(e for e in name if e.isalnum() or e.isspace())


#  Clean and Correct Phone
def clean_phone(phone):
    phone_str = str(phone).replace(' ', '').replace('-', '').replace('(', '').replace(')', '').replace(',', '')
    return phone_str[-10:] if len(phone_str) >= 10 else None


def correct_address_num(num):
    if type(num) is not int:
        return None

    return num


#  Split Address
def split_address(address, client):
    # Split address if address=True
    po_variations = ['PO', 'P.O.', 'P', 'HOMELESS', 'Alden', 'COMFORT', 'POBOX', 'NO', 'Homeless', 'Chicago',
                     'ss', 'NAOMI', 'Thresholds', 'NULL', 'Courtyard', 'YMCA', 'NA', 'CORNERSTONE', 'no',
                     'POBox', 'HOMLESS', 'SALVATION', 'POBOX368463', 'Homeless-', 'homeless', 'Homeless', 'unk', 'unknown', 'UNK', 'Unknown', 'UNKNOWN']
    directions = ['N', 'S', 'E', 'W', 'NW', 'NE', 'SW', 'SE', '#']
    if address is not None:
        try:
            street_arr = address.split()
            if street_arr[0] not in po_variations:
                street_num = street_arr[0]
                new_address = ' '.join(street_arr[1:3])
                client.update({"AddressNumber": street_num,
                               "StreetName": new_address})
        except (AttributeError, IndexError):
            pass


def remove_spaces(item):
    if item is not None:
        return item.replace(' ', '').strip()


def get_zipcode(city, state, client):
    search = SearchEngine()
    count = 0
    if city is not None:
        try:
            zip = search.by_city_and_state(city=city, state=state, returns=1)
            new_zipcode = zip[0].zipcode
            new_county = zip[0].county
            client.update({"Zip": new_zipcode,
                           "ClientCounty": new_county})
        except IndexError:
            count += 1

    print(f'This many clients did not convert: {0}'.format(count))


# Correct Dates
def correct_date(date):
    try:
        if date is None:
            return None
        if isinstance(date, str):
            date = datetime.datetime.strptime(date, '%m/%d/%Y').date()
        if isinstance(date, datetime.datetime):
            date = date.date()
        return date.strftime('%m/%d/%Y')
    except (AttributeError, ValueError, TypeError):
        return date


# correct Race
def correct_race(race):
    race_mapping = {
        'American Indian/Alaskan Native': options['race']['a'],
        'a. American Indian/Alaskan Native': options['race']['a'],
        'Asian': options['race']['b'],
        'b. Asian': options['race']['b'],
        'Black or African American': options['race']['c'],
        'c. Black or African American': options['race']['c'],
        'k. chose not to respond': options['race']['d'],
        'Chose not to respond': options['race']['d'],
        'Unknown': options['race']['d'],
        'None': options['race']['d'],
        'k. Chose not to respond': options['race']['d'],
        'Native Hawaiian or Other Pacific Islander': options['race']['e'],
        'd. Native Hawaiian or Other Pacific Islander': options['race']['e'],
        'Asian and White': options['race']['f'],
        'b. Asian and White': options['race']['f'],
        'g. Asian and White': options['race']['f'],
        'Black or African American and White': options['race']['f'],
        'h. Black or African American and White': options['race']['f'],
        'i. American Indian or Alaska Native and Black or African American': options['race']['f'],
        'More than one race': options['race']['f'],
        'j. Other multiple race': options['race']['f'],
        'Other multiple race': options['race']['f'],
        'White': options['race']['g'],
        'e. White': options['race']['g']
    }
    return race_mapping.get(race, options['race']['d'])



# Correct Ethnicity
def correct_ethnicity(eth):
    ethnicity_mapping = {
        'a. hispanic': options['ethnicity']['a'],
        'a. Hispanic': options['ethnicity']['a'],
        'Hispanic': options['ethnicity']['a'],
        'a.  Hispanic': options['ethnicity']['a'],
        'b. not hispanic': options['ethnicity']['b'],
        'b. Not Hispanic': options['ethnicity']['b'],
        'b. not hispanic': options['ethnicity']['b'],
        'Not Hispanic': options['ethnicity']['b'],
        'b.  Not Hispanic': options['ethnicity']['b'],
        'Chose not to respond': options['ethnicity']['c'],
        'c. Chose not to respond': options['ethnicity']['c'],
        'c. Chose not to respond': options['ethnicity']['c'],
        'c.  Chose not to respond': options['ethnicity']['c']
    }
    return ethnicity_mapping.get(eth, options['ethnicity']['c'])


# Correct Email if needed
def correct_email(email):
    if email is not None:
        test = email.split(']')
        return test[0].replace('[', '')


# Correct English Proficiency
def correct_eng_prof(engProf):
    eng_prof_mapping = {
        'a. Household is Limited English Proficient': options['englishProficiency']['a'],
        'Household is Limited English Proficient': options['englishProficiency']['a'],
        'Limited English Proficient': options['englishProficiency']['a'],
        'Is not English proficient': options['englishProficiency']['a'],
        'b. Household is not Limited English Proficient': options['englishProficiency']['b'],
        'b . Household is not Limited English Proficient': options['englishProficiency']['b'],
        'Household is not Limited English Proficient': options['englishProficiency']['b'],
        'Not Limited English Proficient': options['englishProficiency']['b'],
        'Is English proficient': options['englishProficiency']['b'],
        'Chose not to respond': options['englishProficiency']['c'],
    }
    return eng_prof_mapping.get(engProf, options['englishProficiency']['c'])


# Correct Case Type
def correct_case_type(caseType):
    case_type_mapping = {
        'Home Purchase': options['caseType']['a'],
        'Pre-purchase/Homebuying': options['caseType']['a'],
        'Financial Capability': options['caseType']['a'],
        'Seeking Shelter or Homeless Srvcs': options['caseType']['b'],
        'Homeowner Services': options['caseType']['c'],
        'Non-Delinquency Post-Purchase': options['caseType']['c'],
        'Mortgage Default/Early Delinquency': options['caseType']['d'],
        'Foreclosure Counseling': options['caseType']['d'],
        'Rental Counseling': options['caseType']['e'],
        'Education Services': options['caseType']['f'],
    }
    return case_type_mapping.get(caseType, options['caseType']['a'])


# Correct Rural Status
def correct_rural(ruralStatus):
    rural_status_mapping = {
        'Household lives in a rural area': options['RuralAreaStatus']['a'],
        'Lives in a rural area': options['RuralAreaStatus']['a'],
        'Yes': options['RuralAreaStatus']['a'],
        'Household does not live in a rural area': options['RuralAreaStatus']['b'],
        'Does not live in a rural area': options['RuralAreaStatus']['b'],
        'No': options['RuralAreaStatus']['b'],
        'Chose not to respond': options['RuralAreaStatus']['c'],
        'c. Chose not to respond': options['RuralAreaStatus']['c'],
    }
    return rural_status_mapping.get(ruralStatus, options['RuralAreaStatus']['c'])


# Correct household type
def correct_household(house):
    household_mapping = {
        'Single Adult': options['HouseholdType']['a'],
        'Female-Single Parent': options['HouseholdType']['b'],
        'Female headed single parent household': options['HouseholdType']['b'],
        'Male headed single parent household': options['HouseholdType']['c'],
        'Married without Children': options['HouseholdType']['d'],
        'Married with Children': options['HouseholdType']['e'],
        'Two or more unrelated adults': options['HouseholdType']['f'],
    }
    return household_mapping.get(house, options['HouseholdType']['g'])


# correct class attendance type
def correct_attendance(attendance):
    attendance_mapping = {
        'Graduated': options['Attended']['a'],
        'No Show': options['Attended']['b'],
        'Enrolled': options['Attended']['c'],
    }
    return attendance_mapping.get(attendance, options['Attended']['a'])


# Correct NOFA
def correct_nofa(nofa):
    nofa_mapping = {
        'NOFA-2019': options['NOFA']['b'],
        'NOFA-2020': options['NOFA']['c'],
        'NOFA-2021': options['NOFA']['d'],
    }
    return nofa_mapping.get(nofa, nofa)


def correct_case_id(id):

    return id


# Check class name and update class id number
def get_class(date):
    try:
        if date is not None:
            # print(date)
            if date == 'Buying A Home: Starts 08-17-2021 VIRTUAL':
                return '10400'
            if date == 'Buying A Home: Starts 08-03-2021 VIRTUAL- SPANISH ONLY':
                return '10401'
            if date == 'Homeowner 201: Starts 08-07-2021 VIRTUAL':
                return '10402'
            if date == 'Condo Owner Class: STARTS 08-21-21 VIRTUAL':
                return '10403'
            if date == 'Foreclosure-Refinance Counseling August 2021':
                return '10404'
            if date == 'Framework: 08-26-2021 VIRTUAL':
                return '10405'

            return None
    finally:
        pass


# Get impacts from a separate file
def get_impacts(client, impact_data):
    for row in impact_data.active.iter_rows(min_row=3, min_col=2, values_only=True):
        try:
            impact_case_id = row[0]
            impact_session_id = row[1]
            impact_date = correct_date(row[2])
            impact_arr = row[3].split('.')
            impact = impact_arr[0].replace('-', '')
            if client['CaseID'] == impact_case_id and client['SessionID'] == impact_session_id:
                client.update({impact: impact_date})
        except AttributeError:
            pass


# get class number from a separate file
def get_class_no(classID, sheet):
    for row in sheet.iter_rows(min_row=2, values_only=True, min_col=3):
        if row[0] == classID:
            return row[3]



