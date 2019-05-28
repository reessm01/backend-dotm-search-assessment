#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import zipfile

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Scott Reese"
__github__ = "reessm01"

def check_files(pattern, file_path = './'):
    dotm_files = os.listdir(file_path)
    found = 0
    file_count = 0

    find_pattern(pattern, dotm_files, found, file_count)

def find_pattern(pattern, dotm_files, found, file_count):
    for dotm in dotm_files:
        if str(dotm)[-4:] == 'dotm':
            file_count += 1
            doc = zipfile.ZipFile('./dotm_files/' + dotm)
            xml_content = str(doc.read('word/document.xml'))
            index = xml_content.find(pattern)
            if index != -1 : found+=1
            print_result(dotm, xml_content, index)

    print('Total dotm files searched: ' + str(file_count))
    print('Total dotm files matched: ' + str(found))

def print_result(dotm, xml_content, index):
    while index != -1:
        print('Match found in file ./dotm_files/' + str(dotm))
        print('...' + xml_content[index-40:index+40] + '...')
        index = xml_content.find('$', index+1)

def main():
    args = sys.argv[1:]

    if args[0] == '--dir' and len(args) == 3:
        check_files(args[2], args[1])
    else:
        check_files(args[0])

if __name__ == '__main__':
    main()
