#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Christian Rojas"

import os
import sys
import zipfile
import argparse

DOC_FILENAME = 'word/document.xml'

def create_parser():
    parser = argparse.ArgumentParser(description='Search for text withing dotm files')
    parser.add_argument('--dir', help='directory to search for dotm files')
    parser.add_argument('text', help='text to search within dotm file')
    return parser

def zip_scanner(z, text_search, full_path):
    with z.open(DOC_FILENAME) as doc:
        xml_text = doc.read()
    xml_text = xml_text.decode('utf-8')
    text_loc = xml_text.find(text_search)
    if text_loc >= 0:
        print('Match found in file {}'.format(full_path))
        print(' ...' + xml_text[text_loc-40:text_loc+40] + '...')
        return True

    return False


def main():
    parser = create_parser()
    args = parser.parse_args()

    text_search = args.text
    path_search = args.dir

    if not text_search:
        parser.print_usage()
        sys.exit(1)
    
    print("Searching directory {} from dotm files with text '{}' ...".format(path_search, text_search) )

    file_list = os.listdir(path_search)
    match = 0
    searched = 0

    for file in file_list:
        if not file.endswith('.dotm'):
            print("Disregarding file: " + file)
            continue
        else: 
            searched += 1


        full_path = os.path.join(path_search, file)

        if zipfile.is_zipfile(full_path):
            with zipfile.ZipFile(full_path) as z:
                names = z.namelist()

                if DOC_FILENAME in names:
                    if zip_scanner(z, text_search, full_path): 
                        match += 1
        else:
            print("Not a zipfile: " + full_path)
        
    print('Total dotm files searched: {}'.format(searched))
    print('Total dotm files matched: {}'.format(match))


if __name__ == '__main__':
    main()
