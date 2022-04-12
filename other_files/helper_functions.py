# helper_functions.py

import os
from os import path
import re
import logging
import shutil
import ocrmypdf
from io import StringIO
from pdfminer.high_level import extract_text_to_fp

ocrmypdf.configure_logging(-1)

''' Composes the all document number '''


def format_doc_numer(doc, nr, year):
    # Composes the all document number
    formated_doc_nr = 'CGO-' + doc + nr + '-' + year
    return formated_doc_nr


''' Prompts the user to prees ENTER after showing an error. '''


def prompt_user(msg):
    inp = ''
    while inp != 'y':
        inp = input(msg)
        inp = 'y'


''' returns details of PEI or ECI document '''


def prep_peci_data(peci_desc):
    # Extracts the doc type, *EI or ECI
    doc = peci_desc[1]

    # Extracts and fomats the document number(0000)
    peci_nr = str.zfill((peci_desc[3]).strip(' '), 4)

    # Extracts the year of the document
    peci_year = peci_desc[2]

    # Composes the all document number
    peci_full_file_name = format_doc_numer(doc, peci_nr, peci_year)

    # returns a list with the document number and document year
    lst_pco = [(peci_nr, peci_year)]

    return doc, lst_pco, peci_full_file_name


''' returns details of RI document '''


def prep_ri_data(ri_desc):
    # Retuns RI
    ri_doc = 'RI'

    # Extracts and fomats the ri document number(0000)
    ri_nr = str.zfill((ri_desc[1]).strip(' '), 4)

    # Extracts the year of the ri document
    ri_year = ri_desc[2]

    # Composes the all document number
    ri_full_file_name = format_doc_numer(ri_doc, ri_nr, ri_year)

    # returns a list with the document number and document year
    lst_ri = [(ri_nr, ri_year)]

    return ri_doc, lst_ri, ri_full_file_name


''' Returns true if the current file extention is ".pdf", is a pdf file '''


def ispdf(cur_file):
    if cur_file.endswith(".pdf"):
        return True


''' Creates and returns a PDF/A file '''


def creates_pdfa_file(img_file, pdfa_file):

    # informs if the current file is NOT a scanned file.
    isscanned = False

    # checks if the path and the file exists and are accessible.
    if os.path.isfile(img_file) and os.access(img_file, os.R_OK):
        try:

            # if the original file is a scanned image pdf, makes a new PDF/A file, with a text layer that can be extracted.
            # if __name__ == '__main__':  # To ensure correct behavior on Windows and macOS
            ocrmypdf.ocr(img_file, pdfa_file, deskew=True,
                         progress_bar=False, language='por')

            # informs if the current file is a scanned file.
            isscanned = True

        # if the original file is already a PDF/A file, a exception is thrown where we create a orc file
        # by copying and rename the original original file.
        except TypeError:

            logging.exception(f'The file {img_file} is not a scanned file!')
            # not an image file, already a PDF/A file
            shutil.copy(img_file, pdfa_file)

    return pdfa_file, isscanned


''' extracts text from pdfa file and returns it assigned to the string read_pdf '''


def extracts_text(pdfa_file):
    # Extracts text from pdfa file
    try:
        output_string = StringIO()
        with open(pdfa_file, 'rb') as read_pdf:
            extract_text_to_fp(read_pdf, output_string)

        # returns the assigned text extracted from pdfa file to the string read_pdf
        read_pdf = output_string.getvalue().strip()
    except NotImplementedError as e:
        # log the exception
        logging.exception(e)
        print(f'Exception {e} on extract_text funtion from {pdfa_file}.')

    return read_pdf


''' Returns a list of all pco numbers on the email'''


def list_docs_from_email(pfi_pco, txt_ext):
    if pfi_pco == 'PCOs':
        nrpco_to_pco_re = re.compile(r'PCO CGO/(\d{4})/(\d{4})')
        lst_pcos_to_pco = nrpco_to_pco_re.findall(txt_ext)
        list_pcos = lst_pcos_to_pco

    elif pfi_pco == 'PFIs':
        nrpco_to_pfi_re = re.compile(r'CGO/PCO(\d{4})/(\d{4})')
        lst_pcos_to_pfi = nrpco_to_pfi_re.findall(txt_ext)
        list_pcos = lst_pcos_to_pfi
    return list_pcos


'''
The file is an email
looks for and gets the subject text and pco or pfis details of an email including a list of all pcos of the email.
'''


def prep_email_data(txt_ext):

    subj_re = re.compile(
        r'(PFIs|PCOs) Aprovad(a|o)s em (\d\d-\d\d-\d\d\d\d) (<> )?Lote(:)? (\d{5})( <>)? ')
    subject = subj_re.search(txt_ext)

    if not subject is None:
        # the doc is an email from DSL acknowledging reception of PCOs or PFIs
        # Gets the the type of document con firmed, PCOs or PFIs fro email subject
        pfi_pco = subject.group(1)

        # Gets lote number from email subject
        lote_nr = subject.group(6)

        # Gets the date transmited from the email subject
        date_transm = subject.group(3)

        # nReturns a list of all pco numbmbers on the email
        lst_of_pcos_from_email = list_docs_from_email(pfi_pco, txt_ext)

        # Composes the complete ri file mane
        email_full_file_name = 'RE DSL email receiving ' + \
            pfi_pco + ' lote ' + lote_nr + ' date ' + date_transm

        # Returns the complete archive file number
        return pfi_pco, lst_of_pcos_from_email, email_full_file_name


''' looks for and return the details of RI, PEI and ECI documents. '''


def get_type_doc_details(txt_ext):
    # looks for and gets the details of scanned RI and returns the RI details
    ocr_ri_re = re.compile(r'R?\s?-\s?(\d{1,4})\s?/\s?(\d{4})')
    ri_desc = ocr_ri_re.search(txt_ext)
    if not ri_desc is None:
        ri_details = prep_ri_data(ri_desc)
        return ri_details

    # looks for and gets the details of either scanned PEI or ECI and returns either PEI or ECI
    ocr_peci_re = re.compile(r'(PEI|ECI) (\d{4})\s?/\s?(\d{1,4})')
    peci_desc = ocr_peci_re.search(txt_ext)
    if not peci_desc is None:
        peci_details = prep_peci_data(peci_desc)
        return peci_details

    # looks for and gets the subject text and pco or pfis detail sof an email related with PFIs or PCOs sent to DSL.
    subj_re = re.compile(
        r'(PFIs|PCOs) Aprovad(a|o)s em (\d\d-\d\d-\d\d\d\d) (<> )?Lote(:)? (\d{5})( <>)? ')
    subject = subj_re.search(txt_ext)
    if not subject is None:
        email_details = prep_email_data(txt_ext)
        return email_details


''' returns the complete folter and complete file name to arquive the document. '''


def get_comp_filename(server, doc_type, nr, pco_year, dest_file_name):
    if (doc_type == 'PCOs') or (doc_type == 'ECI'):
        sub_folder = 'ORDER'
        if doc_type == 'PCOs':
            pco = 'PCO-'
        elif doc_type == 'ECI':
            pco = 'ECI-'
    if doc_type == 'PFIs':
        pco = 'PCO-'
        sub_folder = 'PFI'
    if doc_type == 'PFI':
        pco = 'PFIs'
        nr = ''
        sub_folder = ''
    if doc_type == 'PEI':
        pco = 'PEIs'
        nr = ''
        sub_folder = ''
    if doc_type == 'RI':
        pco = 'RIs'
        nr = ''
        sub_folder = ''

    path_to_folder = server+'\\PCOs\\ORDERS ' + \
        pco_year + '\\' + pco + nr + '\\' + sub_folder
    path_to_file = server+'\\PCOs\\ORDERS ' + pco_year + '\\' + \
        pco + nr + '\\' + sub_folder + '\\' + dest_file_name + '.pdf'

    return path_to_folder, path_to_file


''' copies the files to the correspondent folders. '''


def copy_files_to_folders(ocr_file, path_to_folder, path_to_file):

    # if the destination folder exists
    if os.path.exists(path_to_folder):

        # if file exists
        if os.path.exists(path_to_file):

            # append '(1)' to the file name
            path_to_file = os.path.splitext(
                path_to_file)[0] + '(1)' + os.path.splitext(path_to_file)[1]

        # copy and rename the ocr file to the final archive folder
        shutil.copy(ocr_file, path_to_file)

        # inform user the file is archived and ocr file was removed from scanned files
        print(
            f'The original file: {os.path.basename(path_to_file)} was duly archived.')

    else:
        # Inform user the folder extracted from the file is not valid for archive.
        print(
            'The folder extracted from the file,\n{os.path.basename(path_to_folder)}\nis not valid for archive.')

        # Log the folder extracted from the file is not valid for archive.
        logging.info(
            f'The folder extracted from the file {path_to_folder} does not exist!.')
