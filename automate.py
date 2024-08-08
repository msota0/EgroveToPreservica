import pandas as pd
import os
import json

exhibit = input('Enter the collection name: ')
filepath = f'./{exhibit}_metadata_egrove.xlsx'

# df1 for the egrove spreadsheet
df1 = pd.read_excel(filepath)

# df2 for the preservica spreadsheet (assuming this is where you want to map data)
df2 = pd.DataFrame(columns=[
    'title', 'submitted_by', 'document_type', 
    'author1_fname', 'author1_mname', 'author1_lname', 'author1_suffix', 'author1_email', 'author1_institution',
    'author2_fname', 'author2_mname', 'author2_lname', 'author2_suffix', 'author2_email', 'author2_institution',
    'author3_fname', 'author3_mname', 'author3_lname', 'author3_suffix', 'author3_email', 'author3_institution',
    'author4_fname', 'author4_mname', 'author4_lname', 'author4_suffix', 'author4_email', 'author4_institution',
    'abstract', 'embargo_date', 'publication_date', 'keywords', 'comments', 'access_digital_format', 'city', 'county', 'date_info',
    'digital_collection', 'extent', 'finding_aid', 'hasversion', 'identifier', 'language', 'latitude', 'location', 'longitude', 'lyrics', 
    'multimedia_format', 'original_collection', 'original_format', 'physical_description', 'other_form_name', 'publisher', 'relational_format', 
    'reverse_side', 'rights', 'subject_headings', 'transcript', 'upload_cover_image', 'url', 'master_file_size', 'master_file_format', "Notes"
    'supplemental_filenames', 'supplemental_file_sizes', 'supplemental_file_types', 'Notes'
])


# Comments field compilation function
def comments_compilation(row):
    complete_field = []
    # comment = ''
    # artist_inscription = ''
    # cartoon_text = ''
    # visual_elements = ''
    # olemiss_authordates = ''
    # recipient = ''
    
    if 'comments' in row and pd.notnull(row['comments']):
        complete_field .append('comments: ' + str(row['comments']))
    if 'artist_inscription' in row and pd.notnull(row['artist_inscription']):
        complete_field .append('artist_inscription: ' + str(row['artist_inscription']))
        # artist_inscription = 'artist_inscription: ' + str(row['artist_inscription'])
    if 'cartoon_text' in row and pd.notnull(row['cartoon_text']):
        complete_field .append('artist_inscription: ' + str(row['artist_inscription']))
        # cartoon_text = 'cartoon_text: ' + str(row['cartoon_text'])
    if 'visual_elements' in row and pd.notnull(row['visual_elements']):
        complete_field .append('artist_inscription: ' + str(row['artist_inscription']))
        # visual_elements = 'visual_elements: ' + str(row['visual_elements'])
    if 'olemiss_authordates' in row and pd.notnull(row['olemiss_authordates']):
        complete_field .append('artist_inscription: ' + str(row['artist_inscription']))
        # olemiss_authordates = 'olemiss_authordates: ' + str(row['olemiss_authordates'])
    if 'recipient' in row and pd.notnull(row['recipient']):
        complete_field .append('artist_inscription: ' + str(row['artist_inscription']))
        # recipient = 'recipient: ' + str(row['recipient'])

    
    return '; '.join(complete_field)


# Date fields compilation
def date_compilation(row):
    date = []
    if 'date_info' in row and pd.notnull(row['date_info']):
        date.append(str(row['date_info']))
    if 'custom_date' in row and pd.notnull(row['custom_date']):
        date.append(str(row['custom_date']))
    return ';'.join(date)


# original_format fields compilation
def original_format_compilation(row):
    format = []
    if 'original_format' in row and pd.notnull(row['original_format']):
        format.append(str(row['original_format']))
    if 'orignal_format' in row and pd.notnull(row['orignal_format']):
        format.append(str(row['orignal_format']))
    return '; '.join(format)

# Headings fields compilation
def headings_compilation(row):
    headings = []
    if 'subject_headings' in row and pd.notnull(row['subject_headings']):
        headings.append(str(row['subject_headings']))
    if 'subject_heading' in row and pd.notnull(row['subject_heading']):
        headings.append(str(row['subject_heading']))
    if 'lcsh' in row and pd.notnull(row['lcsh']):
        headings.append(str(row['lcsh']))
    return '; '.join(headings)

def author_check(row):
    if 'authors' in row:
        # Convert the string JSON data to a Python list of dictionaries
        authors_list = json.loads(row['authors'])
        
        # Iterate over each author dictionary
        for author in authors_list:
            if author.get('IS_CORPORATE_AUTHOR') == '1':
                corporate_author = author.get('CORPORATE_AUTHOR')
                return True, corporate_author
        
    return False, None


rows = []
for index, row in df1.iterrows():
    interim_row = {}

    title = row.get('title', '')
    submitted_by = 'Abbie Norris-Davidson'
    document_type = row.get('document_type', '')

    is_corporate, corporate_author = author_check(row)
    if not is_corporate:
        author1_fname = row.get('author1_fname', '')
        author1_mname = row.get('author1_mname', '')
        author1_lname = row.get('author1_lname', '')
        author1_suffix = row.get('author1_suffix', '')
        author1_email = row.get('author1_email', '')
        author1_institution = row.get('author1_institution', '')
    else:
        # Handle corporate author case
        author1_fname = corporate_author
        author1_mname = ''
        author1_lname = ''
        author1_suffix = ''
        author1_email = ''
        author1_institution = ''

    author2_fname = row.get('author2_fname', '')
    author2_mname = row.get('author2_mname', '')
    author2_lname = row.get('author2_lname', '')
    author2_suffix = row.get('author2_suffix', '')
    author2_email = row.get('author2_email', '')
    author2_institution = row.get('author2_institution', '')

    author3_fname = row.get('author3_fname', '')
    author3_mname = row.get('author3_mname', '')
    author3_lname = row.get('author3_lname', '')
    author3_suffix = row.get('author3_suffix', '')
    author3_email = row.get('author3_email', '')
    author3_institution = row.get('author3_institution', '')

    author4_fname = row.get('author4_fname', '')
    author4_mname = row.get('author4_mname', '')
    author4_lname = row.get('author4_lname', '')
    author4_suffix = row.get('author4_suffix', '')
    author4_email = row.get('author4_email', '')
    author4_institution = row.get('author4_institution', '')

    notes = 'There are more than 5 authors' if row.get('author5_fname') else ''

    abstract = row.get('abstract', '')
    embargo_date = row.get('embargo_date', '')
    publication_date = row.get('publication_date', '')

    comments = comments_compilation(row)
    access_digital_format = row.get('native_file_type', '')
    city = row.get('city', '')
    county = row.get('county', '')

    date_info = date_compilation(row)
    digital_collection = row.get('digital_collection', '')
    extent = row.get('extent', '')
    finding_aid = row.get('finding_aid', '')
    hasversion = row.get('hasversion', '')
    identifier = row.get('identifier', '')
    language = row.get('language', '')
    latitude = row.get('latitude', '')
    location = row.get('location', '')
    longitude = row.get('longitude', '')
    lyrics = row.get('lyrics', '')
    multimedia_format = row.get('multimedia_format', '')

    original_collection = row.get('original_collection', '')
    original_format = original_format_compilation(row)
    other_form_name = row.get('other_form_name', '')
    physical_description = row.get('physical_description', '')
    publisher = row.get('publisher', '')
    relational_format = row.get('relational_format', '')
    reverse_side = row.get('reverse_side', '')
    rights = row.get('rights', '')

    subject_headings = headings_compilation(row)
    transcript = row.get('transcript', '')
    upload_cover_image = row.get('upload_cover_image', '')
    url = row.get('url', '')

    master_file_size = row.get('master_file_size', '')
    master_file_format = row.get('master_file_format', '')

    supplemental_filenames = row.get('supplemental_filenames', '')
    supplemental_file_sizes = row.get('supplemental_file_sizes', '')
    supplemental_file_types = row.get('supplemental_file_types', '')

    interim_row = {
        'title': title, 'submitted_by': submitted_by, 'document_type': document_type,
        'author1_fname': author1_fname, 'author1_mname': author1_mname, 'author1_lname': author1_lname,
        'author1_suffix': author1_suffix, 'author1_email': author1_email, 'author1_institution': author1_institution,
        'author2_fname': author2_fname, 'author2_mname': author2_mname, 'author2_lname': author2_lname,
        'author2_suffix': author2_suffix, 'author2_email': author2_email, 'author2_institution': author2_institution,
        'author3_fname': author3_fname, 'author3_mname': author3_mname, 'author3_lname': author3_lname,
        'author3_suffix': author3_suffix, 'author3_email': author3_email, 'author3_institution': author3_institution,
        'author4_fname': author4_fname, 'author4_mname': author4_mname, 'author4_lname': author4_lname,
        'author4_suffix': author4_suffix, 'author4_email': author4_email, 'author4_institution': author4_institution,
        'abstract': abstract, 'embargo_date': embargo_date, 'publication_date': publication_date,
        'comments': comments, 'access_digital_format': access_digital_format, 'city': city, 'county': county,
        'date_info': date_info, 'digital_collection': digital_collection, 'extent': extent, 'finding_aid': finding_aid,
        'hasversion': hasversion, 'identifier': identifier, 'language': language, 'latitude': latitude,
        'location': location, 'longitude': longitude, 'lyrics': lyrics, 'multimedia_format': multimedia_format,
        'original_collection': original_collection, 'original_format': original_format, 'other_form_name': other_form_name,
        'physical_description': physical_description, 'publisher': publisher, 'relational_format': relational_format,
        'reverse_side': reverse_side, 'rights': rights, 'subject_headings': subject_headings, 'transcript': transcript,
        'upload_cover_image': upload_cover_image, 'url': url, 'master_file_size': master_file_size,
        'master_file_format': master_file_format, 'supplemental_filenames': supplemental_filenames,
        'supplemental_file_sizes': supplemental_file_sizes, 'supplemental_file_types': supplemental_file_types,
        'Notes': notes
    }

    rows.append(interim_row)

df2 = pd.DataFrame(rows)

output_filepath = f'./{exhibit}_metadata_preservica.xlsx'
df2.to_excel(output_filepath, index=False)
print(f"Data successfully exported to {output_filepath}")