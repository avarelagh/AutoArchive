# she-bang for Python3 scripts:
#!/usr/bin/env python3

"""
ArchiveScannedDocs
This Module is the main app.
"""

import os
import logging
from tkinter import Y
from other_files import helper_functions

logging.basicConfig(filename=r'\\ARBACKUP-PC\Arquive\Scanned Docs\Log\jobdone.log',
                    format='%(asctime)s %(message)s', level=logging.DEBUG)

pc_asserver = r'\\ARBACKUP-PC'
archive = r'\Arquive'
pc_asserver_arch = pc_asserver + archive
scanned_docs_dir = pc_asserver_arch + r'\Scanned Docs'
tmp_dir = scanned_docs_dir + r'\tmp'

''' for all files in scanned docs directory '''
for file in os.listdir(scanned_docs_dir):

    # check if is a pdf file
    if not helper_functions.ispdf(file):
        continue

    # builds the orc file name adding "ocr_" at the beggining of the original file name
    ocr_file = "ocr_" + file

    # builds the name including the PATH of the original, scanned or printed file.
    orig_file = os.path.join(scanned_docs_dir, file)

    # makes the original and ocr the same folder (this way they can be separated any time)
    ocr_scanned_docs_dir = scanned_docs_dir

    # builds the name including the PATH of the OCR file.
    ocr_file = os.path.join(ocr_scanned_docs_dir, ocr_file)

    # Creates and returns a PDF/A file, ocr_file, and if the original file is scanned or printed.
    ocr_file, file_is_scanned = helper_functions.creates_pdfa_file(
        orig_file, ocr_file)

    # extracts text from pdfa file and returns it assigned to the string read_pdf
    text_extracted = helper_functions.extracts_text(ocr_file)

    try:
        # Move original files to tmp folder. These files must be deleted later, after a few days.
        os.replace(orig_file, os.path.join(tmp_dir, file))

    # if the file is being used by another process.
    except PermissionError as e:
        msg = f'Error copying file to tmp folder,\n{e}.\nPlease press ENTER to continue'

        # Prompts the user to prees ENTER after showing an error.
        helper_functions.prompt_user(msg)

        # log the error
        logging.exception(f'Error copying file to tmp folder, {e}.')

        # resume to the next document.
        continue

    # Gets all details of the document to be archived, including a list of all pcos from the emails.
    doc_ext = helper_functions.get_type_doc_details(text_extracted)

    if not doc_ext is None:
        # unboxex document details returned
        doctype, lstofdoc, dest_archive_file_name = doc_ext
    else:
        # if doc_ext is empty, inform user there a ocr error and resume to next document.
        msg = f'ocr error on file {os.path.basename(ocr_file)}.\nPlease press ENTER to continue!'

        # Prompts the user to prees ENTER after showing an error.
        helper_functions.prompt_user(msg)

        # resume to the next document.
        continue

    # if the list of pcos is not empty.
    if not lstofdoc is None:
        for doc in lstofdoc:
            nr, pco_year = doc
            # gets the complete file name, with destination path and the destination file
            file_name = helper_functions.get_comp_filename(
                pc_asserver_arch, doctype, nr, pco_year, dest_archive_file_name)

            # unboxex path and file name returned
            path_to_folder, path_to_file = file_name

            # print(file_name)

            # helper_functions.copy_files_to_folders(lstofdoc, path_to_folder, path_to_file, orig_file_name)
            helper_functions.copy_files_to_folders(
                ocr_file, path_to_folder, path_to_file)

            # log the ocr file was just duly filed on the correspondent folder.
            logging.info(
                f'The original file: {os.path.basename(path_to_file)} was duly archived.')

        if os.path.exists(path_to_folder):
            # remove the just copied ocr file from scanned folder
            os.remove(ocr_file)

            # inform user and log the ocr file was just removed from scanned files
            logging.info(
                f'The ocr file file, {os.path.basename(ocr_file)} was just removed from scanned files folder.\n\n *** ========================================================= ***\n\n')
            print(
                f'The ocr file file, {os.path.basename(ocr_file)} was just removed from scanned files folder.')

        else:
            # creates the new name of the ocr_file not filed.
            dest_filename = scanned_docs_dir + '\\PNA_' + \
                os.path.basename(path_to_file)

            # renames and prefixes 'PNA_' path not available to ocr file
            os.rename(ocr_file, dest_filename)

            # inform user and log the ocr file was not removed from scanned files because the path to archive was not found.
            logging.info(
                f'The ocr file file, {os.path.basename(ocr_file)} was NOT removed from scanned files folder.\n\n *** ========================================================= ***\n\n')
            print(
                f'The ocr file file, {os.path.basename(ocr_file)} was NOT removed from scanned files folder.')

    else:
        # inform user the email was not valid for archive.
        print('File not valid for arquive.')
