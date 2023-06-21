import fitz
import io
import sys
import os
import re
from Crypto.Cipher import AES
from time import time
from datetime import datetime
from re import sub, search, findall, IGNORECASE
from time import time
from PyPDF2 import PdfReader
from PIL import Image


def find_text(doc, searchterm):
    result_list = []
    pages = []
    reader = PdfReader(doc)
    # print(searchterm)
    # print(reader.numPages)
    for page_number in range(0, reader.numPages):
        page = reader.getPage(page_number)
        page_content = page.extractText()
        # counter = 0
        # print(page_content)
        if searchterm in page_content:
            # print("found")
            # counter += 1
            pages.append(page_number)

    if pages:
        result = {
            "search string": searchterm,
            "pages": pages
        }
        result_list.append(result)

    return result_list


# # open the file
# pdf_file = fitz.open(file)
#
# # iterate over PDF pages
# for page_index in range(len(pdf_file)):
#     # get the page itself
#     page = pdf_file[page_index]
#     image_list = page.get_images()
#     # printing number of images found in this page
#     if image_list:
#         print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
#     else:
#         print("[!] No images found on page", page_index)
#     for image_index, img in enumerate(page.get_images(), start=1):
#         # get the XREF of the image
#         xref = img[0]
#         # extract the image bytes
#         base_image = pdf_file.extract_image(xref)
#         image_bytes = base_image["image"]
#         # get the image extension
#         image_ext = base_image["ext"]
#         # get the image height
#         image_height = base_image["height"]
#         # get the image width
#         image_width = base_image["width"]
#         # load it to PIL
#         image = Image.open(io.BytesIO(image_bytes))
#         # save it to local disk
#         if image_height > 2:
#             print(image_height)
#             image.save(open(f"images/image{page_index+1}_{image_index}.{image_ext}", "wb"))


def process_file(file_in, config, log):
    '''
    Replaces the old BHGE logo and copyright information

    Parameters
    ----------
    file_in: string
        the path to the input file

    file_out: string
        the path to the output file

    config: dictionary
        content of the configuration file as dictonary with the tag as
        key and the information as value
    '''
    file_path = file_in
    print(file_path)
    output_log = []
    no_match = "No Matches"
    if file_path.endswith('.pdf'):

        for replace_duo in config["ReplaceString"]:
            # file path you want to extract images from "pdfs/100042-1.pdf"
            # file = "pdfs/QA-GLB-En-103321.pdf"
            output = find_text(file_path, replace_duo[0])
            if output:
                # print(output)
                output_log.append(output)
        # print(output_log)
        if output_log:
            log.write(f"{file_in};{output_log}\n")
        else:
            log.write(f"{file_in};{no_match}\n")


def main():

    if len(sys.argv) == 2:
        # Put elements of config file into a dictionary
        config = map_config(sys.argv[1])

        # Check if log folder exists and create it if neccessary
        if not os.path.isdir(config["LogFolder"]):
            os.mkdir(config["LogFolder"])

        # Check if input directory exists
        if os.path.isdir(config["InputFolder"]):
            log = prepare_log(config)

            for file in os.listdir(config["InputFolder"]):
                # Start timer
                start_time = time()

                # Create input and output path and start file processing
                file_in = os.path.join(config["InputFolder"], file)

                # Check if current file is a path to a folder
                if os.path.isdir(file_in):
                    # Ignore folder and continue
                    continue

                # Process the file
                print(f"Processing {file}")
                # process_file(file_in, file_out, config, log)
                process_file(file_in, config, log)

                # End timer and output time
                print(f"Processing took {time() - start_time:.3f} seconds")

            log.close()


def prepare_log(config):

    log_time = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    log_name = f"PDF_Validation_{log_time}.csv"
    log_path = os.path.join(config["LogFolder"], log_name)

    # pylint: disable=R1732
    # Disables consider using with for resource allocation
    # Log needs to be closed in main function after processing all files
    log = open(log_path, "w+", encoding="utf-8")

    log.write("sep=;\n")
    log.write("Inputfile;Notes\n")

    print(f"Logging in file {log_name}")
    return log


def map_config(configfile):
    """
    This function creates a dictionary from the information gathered form
    the configuration file.

    Parameters
    ----------
    configfile: string
        the path to the configuration file

    Returns
    -------
    cfg_dict: dictonary
        content of the configuration file as dictonary with the tag as
        key and the information as value

    """
    cfg_dict = {}
    cfg_dict["ReplaceString"] = []
    cfg_dict["case-unsensitive"] = []

    with open(configfile, encoding="utf-8") as conf:
        for line in conf:
            if not line.startswith("//"):
                if search(r'[\S\s]*? = [\S\s]*?', line) is not None:
                    (key, val)= line.split(' = ')
                    cfg_dict[key] = val.strip()
                elif search(r'[\S\s]*? -> [\S\s]*?', line) is not None:
                    line = str(line)
                    'line = bytes(line,  "utf-8")'
                    (old, new) = line.split(' -> ')
                    # pylint: disable=W1401
                    # Disable error from using \S\s in a binary string which is needed for regex
                    # if search(b'[\S\s]*? // case-unsensitive', new) is not None:
                    #    new = (new[0:new.find(b" // case-unsensitive")])
                    #    cfg_dict["case-unsensitive"].append(new)
                    cfg_dict["ReplaceString"].append([old.strip(), new.strip()])

    return cfg_dict


if __name__ == "__main__":
    main()


