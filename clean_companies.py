import locale
import pandas as pd
import numpy as np
from dateutil import parser
from datetime import datetime
import re

#---------------REGION FORMATS-------------------
locale.setlocale(locale.LC_NUMERIC, '') 
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

#--------------------DATA------------------------
# default missing values  
missing_currency = 'EUR'
missing_amount = '0'
missing_cp = '00000'
missing_activity_code = '0000'
missing_NAICS_activity_code = '000000'
missing_number_of_directors = 0
missing_employees_total= 0
missing_online_shop= 0
missing_number_of_reviews = 0
missing_average_rating = 0
missing_year = '-'
missing_status = 'Inactive'
missing_mode = 'Estimated'
missing_phone = '-'
missing_email = '-'
missing_activity = '-'
missing_web = '-'
missing_hours = '-'

# to delete in 'amount'
currency_symbols = ['$', 'USD', '€', 'EUR', '£']

# to unify 'status'
status_for_Active = ['activa', 'active']

# to unify 'mode'
mode_for_Estimated = ['estimated', 'estimate']


cities = [
    ['12400', 'Segorbe', 'Castellon'], 
    ['46930', 'Quart de Poblet', 'Valencia'], 
    ['31500', 'Tudela', 'Navarra'], 
    ['28823', 'Coslada', 'Madrid'], 
    ['08008', 'Barcelona','Barcelona'], 
    ['28040', 'Aravaca', 'Madrid'],
    ['17430', 'Santa Coloma de Farners','Girona'], 
    ['28232', 'Las Rozas de Madrid', 'Madrid'], 
    ['46185', 'La Pobla de Vallbona', 'Valencia'],
    ['18006', 'Ronda', 'Granada'], 
    ['30009', 'Murcia','Murcia'], 
    ['08243', 'Manresa','Barcelona']
]

df_cities = pd.DataFrame(cities, columns=['postal_code', 'city_name', 'province_name'])

countries = [
    ['España', 'ES', '+34'],
    ['Spain', 'ES', '+34'],
    ['Germany', 'DE', '+49'],
    ['Alemania', 'DE', '+49'],
    ['Great Britain', 'GB', '+44'],
    ['Reino Unido', 'GB', '+44']
]

df_countries = pd.DataFrame(countries, columns=['country_name', 'country_code', 'phone_prefix'])

type_of_streets = [
    ['C.', 'Calle'],
    ['C/', 'Calle'],
    ['Av.', 'Avenida'],
    ['Av ', 'Avenida']
]

df_type_of_streets = pd.DataFrame(type_of_streets, columns=['abbreviation', 'type_street'])


#-------------SOME FUNCTIONS----------------------

# DELETE A CHAR IF EXISTS FUNCTION
def delete_char(in_word, char_to_delete):
    if char_to_delete in in_word:
        result = in_word.replace(char_to_delete,'').strip()
    else:
        result = in_word.strip()
    return result


# CLEAN 'amount' FUNCTION
def clean_amount(column_amount, currency_symbols):    
    column_amount = column_amount.astype(str).str.strip()
    column_amount = column_amount.str.replace('.', '')      
                                                     
    # remove symbols
    for symbol in currency_symbols:
        column_amount = column_amount.str.replace(symbol, '')

    # 'amount' String to int
    column_amount = column_amount.apply(lambda x: int(locale.atof(x)))

    return column_amount



# EXTRACT YEAR FROM A DATE
def extract_year(in_date):
    try:
        # do your best
        parsed_year = parser.parse(in_date, fuzzy=True)
        return parsed_year.year
    except ValueError:   # in letters
        parsed_year = in_date.split()
        parsed_year = parsed_year[-1]  #last
        if len(parsed_year) > 4:
            return None
        else:
            return parsed_year



# CLEAN 'year' AND 'incorporated' FUNCTION
def clean_years(df):
    columns = ['year', 'incorporated']
    for col in columns:
        df[col] = df[col].str.strip()
        symbols_to_delete = ['.', ',']
        for symb in symbols_to_delete:
            df[col] = df[col].str.replace(symb, '')
        df[col] = df[col].apply(extract_year)  
    
    return df


# UNIFY 'status' FUNCTION
def unify_status(column_status, status_for_Active):
    column_status = column_status.str.strip().str.lower()

    condition = column_status.isin(status_for_Active)
    column_status[condition] = 'Active'
    column_status[~condition] = 'Inactive'
    
    return column_status


# UNIFY 'mode' FUNCTION
def unify_mode(column_mode, mode_for_Estimated):
    column_mode = column_mode.str.strip().str.lower()
    
    condition = column_mode.isin(mode_for_Estimated)
    column_mode[condition] = 'Estimated'
    column_mode[~condition] = 'Real'

    return column_mode

# FORMAT NAME FUNCTION
def format_name(column_name):
    column_name = column_name.str.strip().str.upper()
    column_name = column_name.str.replace('.', '')      
    column_name = column_name.str.replace('SOCIEDAD LIMITADA', 'SL')
    column_name = column_name.str.replace('SOCIEDAD ANONIMA', 'SA') 

    # conditional replacement  ~=not
    column_name = column_name.where(~column_name.str.endswith(' SA'), column_name.str.replace(' SA', ', S.A.'))
    column_name = column_name.where(~column_name.str.endswith(' SL'), column_name.str.replace(' SL', ', S.L.'))
    column_name = column_name.where(~column_name.str.endswith(' SAU'), column_name.str.replace(' SAU', ', S.A.U.'))
    column_name = column_name.where(~column_name.str.endswith(' SCCL'), column_name.str.replace(' SCCL', ', S.C.C.L.'))
    column_name = column_name.where(~column_name.str.endswith(' BV'), column_name.str.replace(' BV', ', B.V.'))

    return column_name



# FORMAT STREET TYPES FUNCTION
def street_type(address, df_type_of_streets):
    for index, row in df_type_of_streets.iterrows():
        abbreviation = row['abbreviation']
        type_street = row['type_street']
        if abbreviation in address:
            address = address.replace(abbreviation, type_street)
            return address #only one replacement
    return address



# EXTRACT COUNTRY FUNCTION
def extract_country(address, countries_list):
    address_lower = address.lower()
    for country in countries_list:
        country_lower = country.lower()
        if address_lower.startswith(country_lower) or address_lower.endswith(country_lower):
            return country.strip()

    return None


# EXTRACT CITY FUNCTION
def extract_city(address, cities_list):
    address_lower = address.lower()
    for city in cities_list:
        city_lower = city.lower()
        if city_lower in address_lower:
            return city.strip()

    return None



# CAPITALIZE FUNCTION
# Función para capitalizar palabras específicas
def words_caps(text_to_capitalize):
    no_capitalize = ['el', 'la', 'los', 'las', 'de', 'del', 'en', 'y']
    words = text_to_capitalize.split() 
    
    result = [words[0].capitalize()]
    for idx, element in enumerate(words[1:]):
        if element.lower() in no_capitalize:
            result.append(element.lower())
        else:
            result.append(element.capitalize())
    
    return ' '.join(result)


# FORMAT ADDRESS FUNCTION
def format_address(df, df_type_of_streets, df_cities, df_countries):
    
    # preformatting
    df['address'] = df['address'].str.strip()
    df['address'] = df['address'].str.replace(' - ', ', ')
    df['address'] = df['address'].str.replace('- ', ', ')
    df['address'] = df['address'].str.replace(';', ',')
    df['address'] = df['address'].apply(lambda x: street_type(x, df_type_of_streets))
    df['address'] = df['address'].apply(words_caps)
    
    #---------------------------Postal Code------------------------------------
    
    # extract, save and delete from address
    df['cp'] = df['address'].str.extract(r'(\d{5})')  # r=raw literal, 5 digits
    df['address'] = df['address'].str.replace(r'\b\d{5}\b', '', regex=True).str.strip()

    #----------------------------Country---------------------------------------
    
    # extract at the begining or end, save and delete from address
    countries_list = df_countries['country_name'].tolist()
    
    df['country_name'] = df['address'].apply(lambda x: extract_country(x, countries_list)).str.capitalize()
    df['address'] = df.apply(lambda row: row['address'].lower().replace(row['country_name'].lower(), '').strip(', ').capitalize(), axis=1)
    
    #------------------------------City----------------------------------------

    # search city by known cp
    df = df.merge(df_cities[['postal_code', 'city_name']], left_on='cp', right_on='postal_code', how='left')
    df.rename(columns={'city_name': 'city'}, inplace=True)
    df.drop('postal_code', axis=1, inplace=True)
    df['address'] = df.apply(lambda row: row['address'].lower().replace(row['city'].lower(), '').strip(', ').capitalize(), axis=1)   
    df['address'] = df['address'].apply(words_caps)
    
    #-----------------------------Province-------------------------------------
    
    # search province by known cp
    df = df.merge(df_cities[['postal_code', 'province_name']], left_on='cp', right_on='postal_code', how='left')
    df.rename(columns={'province_name': 'province'}, inplace=True)
    df.drop('postal_code', axis=1, inplace=True)
    
    #delete from address if exists
    df['province'] = df['province'].astype(str)
    for province in df['province']:
        df['address'] = df['address'].str.replace(province, '').str.strip(', ')
        
        
    # some cleaning
    df['address'] = df['address'].str.replace(' ,', ',')
    df['address'] = df['address'].str.replace(',,', ',')
    df['address'] = df['address'].str.replace(', ,', ',')
    df['address'] = df['address'].str.replace(',,', ',')
    #df['address'] = df['address'].apply(words_caps)
    
    return df



# FORMAT PHONE NUMBER FUNCTION
# add the national telephone prefix and to-String
# country needed
def format_phone(df, df_countries):
    merged_df = df.merge(df_countries, how='left', left_on='country_name', right_on='country_name')
    
    for index, row in merged_df.iterrows():
        phone_number = str(row['phone_number'])
        phone_number = phone_number.split('.')[0]  # no decimals
        phone_prefix = str(row['phone_prefix'])
        if not pd.isnull(phone_prefix) and phone_number != '0':
            if not phone_number.startswith(phone_prefix):
                df.at[index, 'phone_number'] = phone_prefix + phone_number
    
    return df['phone_number']
        
 


# JOIN CONFLICTING VALUES FUNCTION
def join_conflicting(x):
    
    # 'mode' and 'hiring' all values needed, except if they're all empty
    if x.name == 'mode' or x.name == 'hiring':  
        non_empty_values = [str(val) for val in x if pd.notnull(val)]
        unique_values = x.unique()
        if len(unique_values) == 1:
            return unique_values[0]
        if len(non_empty_values) > 0:
            return '//'.join(non_empty_values)
        else:
            return ''
    else:
        # non repeteated values to be joined
        unique_values = pd.unique(x)  
        
        # iterates on unique values but non empty
        non_empty_values = [str(val) for val in unique_values if pd.notnull(val)] 
        
        # returns empty if all values are empty
        return '//'.join(non_empty_values) if non_empty_values else ''  # empty if all values are empty



#----------------------------LOAD DATA-----------------------------------------

# load data from Excel file to Pandas dataframe
excel_file = 'C:/Users/Carmelo/Downloads/original.xlsx'  
df = pd.read_excel(excel_file)


#-------------------------DATAFRAME STRUCTURE----------------------------------

# add new columns
df['cp'] = missing_cp
df['country_name']=''

# drop redundant columns
df = df.drop('country_user', axis=1)    # ='country'
df = df.drop('name.1', axis=1)         # ='name'

# columns names to lowercase
df = df.rename(columns={'Longitude': 'longitude', 'Latitude': 'latitude'})


#-----------------------------PREPARE DATA-------------------------------------

# drop error observations: 'name' empty
df = df.dropna(subset=['name'], axis=0).reset_index(drop=True)

# drop columns with no values
df = df.dropna(axis=1, how='all')

# missing_values and change type 
df['number_of_directors'].fillna(missing_number_of_directors, inplace=True)
df['number_of_directors'] = df['number_of_directors'].astype('int')

df['employees_total'].fillna(missing_employees_total, inplace=True)
df['employees_total'] = df['employees_total'].astype('int')

df['online_shop'].fillna(missing_online_shop, inplace=True)
df['online_shop'] = df['online_shop'].astype('int')

df['number_of_reviews'].fillna(missing_number_of_reviews, inplace=True)
df['number_of_reviews'] = df['number_of_reviews'].astype('int')

df['average_rating'].fillna(missing_average_rating, inplace=True)  #float

df['amount'] = df['amount'].astype(str)
df['amount'] = df['amount'].replace('nan', np.nan)
df['amount'].fillna(missing_amount, inplace=True)

df['currency'] = df['currency'].replace('nan', np.nan)
df['currency'].fillna(missing_currency, inplace=True)

df['status'].fillna(missing_status, inplace=True)
df['mode'].fillna(missing_mode, inplace=True)

df['phone_number'] = df['phone_number'].fillna(missing_phone).astype(str)

df['year'] = df['year'].fillna(missing_year).astype(str)
df['incorporated'] = df['incorporated'].fillna(missing_year).astype(str)

df['NAICS_activity_secondary'] = df['NAICS_activity_secondary'].fillna(missing_activity).astype(str)
df['Secondary_Activity_Description'] = df['Secondary_Activity_Description'].fillna(missing_activity).astype(str)
df['web_description'] = df['web_description'].fillna(missing_web).astype(str)
df['web_description'] = df['web_description'].fillna(missing_web).astype(str)
df['website'] = df['website'].fillna(missing_web).astype(str)
df['social_link_facebook'] = df['social_link_facebook'].fillna(missing_web).astype(str)
df['social_link_twitter'] = df['social_link_twitter'].fillna(missing_web).astype(str)
df['social_link_linkedin'] = df['social_link_linkedin'].fillna(missing_web).astype(str)
df['social_link_youtube'] = df['social_link_youtube'].fillna(missing_web).astype(str)
df['social_link_instagram'] = df['social_link_instagram'].fillna(missing_web).astype(str)

df['branches_addresses'] = df['branches_addresses'].fillna(missing_web).astype(str)
df['hours'] = df['hours'].fillna(missing_hours).astype(str)

df['activity_code'] = df['activity_code'].fillna(missing_activity_code).astype(int).astype(str).str.zfill(4)
df['NAICS_activity_code'] = df['NAICS_activity_code'].fillna(missing_NAICS_activity_code).astype(int).astype(str).str.zfill(6)
df['NAICS_activity_secondary_code'] = df['NAICS_activity_secondary_code'].fillna(missing_NAICS_activity_code).astype(int).astype(str).str.zfill(6)
df['activity_code_secondary'] = df['activity_code_secondary'].fillna(missing_activity_code).astype(int).astype(str).str.zfill(4)
df['activity_code_other'] = df['activity_code_other'].fillna(missing_activity_code).astype(str)


# some cleaning and formatting
df['longitude'] = df['longitude'].astype(float).round(8).map('{:.8f}'.format)
df['latitude'] = df['latitude'].astype(float).round(8).map('{:.8f}'.format)
df['email'] = df['email'].fillna(missing_email).astype(str)
df['email'] = df['email'].apply(lambda x: str(x).split('?')[0])
df['hiring'] = df['hiring'].fillna(False).astype(int)  # TRUE-> 1, empty-> 0

# transformations
df['amount'] = clean_amount(df['amount'], currency_symbols)
df = clean_years(df)
df['status'] = unify_status(df['status'], status_for_Active)
df['mode'] = unify_mode(df['mode'], mode_for_Estimated)
df['name'] = format_name(df['name'])
df = format_address(df, df_type_of_streets, df_cities, df_countries)
df['phone_number'] = format_phone(df, df_countries)

#-------------------------------SAVE INTERMEDIATE DATA---------------------------------------

# save intermediate result to Excel

df_intermediate = df[['id','country','code',	'name',	'address.1','address','cp','city','province','country_name','activity_code','activity','NAICS_activity','NAICS_activity_code','NAICS_activity_secondary',	'NAICS_activity_secondary_code',	'activity_code_secondary',	'Secondary_Activity_Description','activity_code_other',	'web_description','legal_form',	'employees_total',	'number_of_directors',	'year',	'amount',	'currency',	'mode',	'company_number',	'incorporated',	'status','longitude','latitude','phone_number','email','website',	'branches_addresses',	'online_shop',	'number_of_reviews',	'average_rating',	'social_link_facebook',	'social_link_twitter','social_link_linkedin','social_link_youtube',	'social_link_instagram',	'hiring',	'hours',	'confidence_score'
]].copy()
df_intermediate.to_excel('C:/Users/Carmelo/Downloads/intermediate.xlsx', index=False)



#-------------------------------UNIFY ROWS-------------------------------------

# apply an aggregation function (join_conflicting) to each group of rows with the same id
grouped_df = df.groupby('id').agg(lambda x: join_conflicting(x)).reset_index()




#-------------------------------SAVE FINAL RESULT------------------------------

# save result to Excel

final_df = grouped_df[['id','country','code',	'name',	'address.1','address','cp','city','province','country_name','activity_code','activity','NAICS_activity','NAICS_activity_code','NAICS_activity_secondary',	'NAICS_activity_secondary_code',	'activity_code_secondary',	'Secondary_Activity_Description','activity_code_other',	'web_description','legal_form',	'employees_total',	'number_of_directors',	'year',	'amount',	'currency',	'mode',	'company_number',	'incorporated',	'status','longitude','latitude','phone_number','email','website',	'branches_addresses',	'online_shop',	'number_of_reviews',	'average_rating',	'social_link_facebook',	'social_link_twitter','social_link_linkedin','social_link_youtube',	'social_link_instagram',	'hiring',	'hours',	'confidence_score'
]].copy()
final_df.to_excel('C:/Users/Carmelo/Downloads/final_result.xlsx', index=False)
